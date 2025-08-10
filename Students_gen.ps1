
Param(
    [Parameter(Mandatory=$false)] [string] $SearchBase = 'OU=groups,OU=SVPO,OU=KRIGT,DC=krsk,DC=irgups,DC=ru',
    [Parameter(Mandatory=$false)] [string] $ExcludeOUNameLike = 'groups',
    [Parameter(Mandatory=$false)] [string] $OutputDirectory = 'C:\vpo',
    [Parameter(Mandatory=$false)] [int] $PassValidityDays = 1865,
    [Parameter(Mandatory=$false)] [string] $TypeValue = 'институт',
    [Parameter(Mandatory=$false)] [string] $CompanyValue = 'SVPO',
    [Parameter(Mandatory=$false)] [switch] $NoGui
)

#Requires -Modules ActiveDirectory

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-ActiveDirectoryModule {
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
        throw 'Модуль ActiveDirectory не найден. Установите RSAT и модуль ActiveDirectory.'
    }
    Import-Module ActiveDirectory -ErrorAction Stop | Out-Null
}

function Ensure-OutputDirectoryExists {
    Param([string] $DirectoryPath)
    if (-not (Test-Path -LiteralPath $DirectoryPath)) {
        New-Item -ItemType Directory -Path $DirectoryPath | Out-Null
    }
}

function Get-TargetOrganizationalUnits {
    Param(
        [string] $SearchBaseDn,
        [string] $ExcludeNameLike
    )
    $filterString = "Name -notlike '$ExcludeNameLike'"
    Get-ADOrganizationalUnit -Filter $filterString -SearchBase $SearchBaseDn -Properties DistinguishedName, Name |
        Select-Object -Property DistinguishedName, Name
}

function Get-UsersForOrganizationalUnit {
    Param(
        [string] $OrganizationalUnitDn,
        [string] $CompanyValueParam
    )
    Get-ADUser -SearchBase $OrganizationalUnitDn -Filter * -Properties givenName, sn, displayName, sAMAccountName, Department |
        Select-Object -Property @(
            @{ Name = 'Name'; Expression = { ($_.givenName -split '\\s+')[0] } },
            @{ Name = 'MidName'; Expression = { ($_.displayName -split '\\s+')[ -1 ] } },
            @{ Name = 'LastName'; Expression = { $_.sn } },
            @{ Name = 'TabNumber'; Expression = { $_.sAMAccountName } },
            'Department',
            @{ Name = 'WorkPhone'; Expression = { '' } },
            @{ Name = 'HomePhone'; Expression = { '' } },
            @{ Name = 'BirthDate'; Expression = { '' } },
            @{ Name = 'Address'; Expression = { '' } },
            @{ Name = 'Post'; Expression = { '' } },
            @{ Name = 'Company'; Expression = { $CompanyValueParam } },
            @{ Name = "'"; Expression = { 'Нет' } },
            @{ Name = 'Status'; Expression = { 'Хозорган' } }
        )
}

function Export-UsersCsv {
    Param(
        [System.Collections.IEnumerable] $Users,
        [string] $FilePath
    )
    $Users | ConvertTo-Csv -NoTypeInformation -Delimiter ';' |
        ForEach-Object { $_ -replace '"', '' } |
        Out-File -FilePath $FilePath -Encoding Default
}

function Export-PassesCsv {
    Param(
        [System.Collections.IEnumerable] $Users,
        [datetime] $StartDate,
        [datetime] $EndDate,
        [string] $TypeValueParam,
        [string] $FilePath
    )
    $Users | Select-Object -Property @(
        @{ Name = 'blan'; Expression = { '' } },
        @{ Name = 'blank2'; Expression = { '' } },
        'TabNumber',
        @{ Name = 'type'; Expression = { $TypeValueParam } },
        @{ Name = 'start_date'; Expression = { $StartDate.ToString('dd.MM.yyyy') } },
        @{ Name = 'end_date'; Expression = { $EndDate.ToString('dd.MM.yyyy') } },
        @{ Name = 'blank3'; Expression = { '' } }
    ) |
    ConvertTo-Csv -NoTypeInformation -Delimiter ';' |
    ForEach-Object { $_ -replace '"', '' } |
    Out-File -FilePath $FilePath -Encoding Default
}

function Get-SafeFileNameFromOuName {
    Param([string] $OuName)
    $name = $OuName -replace '\\.', '-'
    $invalid = [System.IO.Path]::GetInvalidFileNameChars() -join ''
    $regex = "[" + [System.Text.RegularExpressions.Regex]::Escape($invalid) + "]"
    return ($name -replace $regex, '_')
}

function Generate-ForAllOrganizationalUnits {
    Param(
        [string] $SearchBaseDn,
        [string] $ExcludeNameLike,
        [string] $OutputDir,
        [int] $ValidityDays,
        [string] $TypeValueParam,
        [string] $CompanyValueParam,
        [System.Windows.Forms.TextBox] $LogTextBox = $null
    )

    Ensure-ActiveDirectoryModule
    Ensure-OutputDirectoryExists -DirectoryPath $OutputDir

    $startDateObj = Get-Date
    $endDateObj = (Get-Date).AddDays($ValidityDays)

    $organizationalUnits = Get-TargetOrganizationalUnits -SearchBaseDn $SearchBaseDn -ExcludeNameLike $ExcludeNameLike

    $index = 0
    foreach ($ou in $organizationalUnits) {
        $index++
        $safeName = Get-SafeFileNameFromOuName -OuName $ou.Name
        $users = Get-UsersForOrganizationalUnit -OrganizationalUnitDn $ou.DistinguishedName -CompanyValueParam $CompanyValueParam

        $usersFile = Join-Path $OutputDir "$safeName.csv"
        $passesFile = Join-Path $OutputDir "$safeName-pass.csv"

        Export-UsersCsv -Users $users -FilePath $usersFile
        Export-PassesCsv -Users $users -StartDate $startDateObj -EndDate $endDateObj -TypeValueParam $TypeValueParam -FilePath $passesFile

        $message = "[$index] OU: $($ou.Name) -> $([System.IO.Path]::GetFileName($usersFile)), $([System.IO.Path]::GetFileName($passesFile))"
        if ($LogTextBox) { $LogTextBox.AppendText($message + [Environment]::NewLine) } else { Write-Host $message }
    }

    $doneMessage = "Готово. Файлы сохранены в: $OutputDir"
    if ($LogTextBox) { $LogTextBox.AppendText($doneMessage + [Environment]::NewLine) } else { Write-Host $doneMessage }
}

function Start-Gui {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    Add-Type -AssemblyName System.Drawing | Out-Null

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Генератор CSV студентов/пропусков'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(720, 540)

    $labelWidth = 160
    $inputWidth = 460
    $leftMargin = 16
    $top = 16
    $rowHeight = 28

    function Add-LabelAndTextBox([string] $labelText, [string] $defaultText) {
        param([int] $rowIndex)
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $labelText
        $label.AutoSize = $true
        $label.Location = New-Object System.Drawing.Point($leftMargin, $top + $rowIndex * $rowHeight)
        $form.Controls.Add($label)

        $textbox = New-Object System.Windows.Forms.TextBox
        $textbox.Width = $inputWidth
        $textbox.Location = New-Object System.Drawing.Point($leftMargin + $labelWidth, $top + $rowIndex * $rowHeight - 3)
        $textbox.Text = $defaultText
        $form.Controls.Add($textbox)
        return $textbox
    }

    $tbSearchBase = Add-LabelAndTextBox -labelText 'SearchBase DN:' -defaultText $SearchBase -rowIndex 0
    $tbExclude = Add-LabelAndTextBox -labelText 'Исключить OU (Name -notlike):' -defaultText $ExcludeOUNameLike -rowIndex 1
    $tbOutput = Add-LabelAndTextBox -labelText 'Папка вывода:' -defaultText $OutputDirectory -rowIndex 2
    $tbDays = Add-LabelAndTextBox -labelText 'Срок действия (дней):' -defaultText $PassValidityDays -rowIndex 3
    $tbType = Add-LabelAndTextBox -labelText 'Значение поля type:' -defaultText $TypeValue -rowIndex 4
    $tbCompany = Add-LabelAndTextBox -labelText 'Компания:' -defaultText $CompanyValue -rowIndex 5

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = '...'
    $btnBrowse.Width = 32
    $btnBrowse.Location = New-Object System.Drawing.Point($leftMargin + $labelWidth + $inputWidth + 6, $top + 2 * $rowHeight - 6)
    $btnBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.SelectedPath = $tbOutput.Text
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $tbOutput.Text = $dlg.SelectedPath
        }
    })
    $form.Controls.Add($btnBrowse)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = 'Сгенерировать'
    $btnRun.Width = 160
    $btnRun.Location = New-Object System.Drawing.Point($leftMargin, $top + 7 * $rowHeight)
    $form.Controls.Add($btnRun)

    $tbLog = New-Object System.Windows.Forms.TextBox
    $tbLog.Multiline = $true
    $tbLog.ScrollBars = 'Vertical'
    $tbLog.ReadOnly = $true
    $tbLog.WordWrap = $false
    $tbLog.Location = New-Object System.Drawing.Point($leftMargin, $top + 8.5 * $rowHeight)
    $tbLog.Size = New-Object System.Drawing.Size(670, 320)
    $form.Controls.Add($tbLog)

    $btnRun.Add_Click({
        try {
            $btnRun.Enabled = $false
            $tbLog.AppendText("Старт..." + [Environment]::NewLine)
            $genParams = @{
                SearchBaseDn      = $tbSearchBase.Text
                ExcludeNameLike   = $tbExclude.Text
                OutputDir         = $tbOutput.Text
                ValidityDays      = [int]$tbDays.Text
                TypeValueParam    = $tbType.Text
                CompanyValueParam = $tbCompany.Text
                LogTextBox        = $tbLog
            }
            Generate-ForAllOrganizationalUnits @genParams
        }
        catch {
            $tbLog.AppendText('Ошибка: ' + $_.Exception.Message + [Environment]::NewLine)
        }
        finally {
            $btnRun.Enabled = $true
        }
    })

    $form.Add_Shown({ $form.Activate() })

    # Проверка STA для WinForms
    $isSta = [System.Threading.Thread]::CurrentThread.ApartmentState -eq 'STA'
    if (-not $isSta) {
        [System.Windows.Forms.MessageBox]::Show('Для GUI-требуется запуск в STA. Запустите PowerShell с параметром -STA.', 'Требуется STA') | Out-Null
    }

    [void] $form.ShowDialog()
}

# CLI / GUI entry
if ($NoGui) {
    $params = @{
        SearchBaseDn      = $SearchBase
        ExcludeNameLike   = $ExcludeOUNameLike
        OutputDir         = $OutputDirectory
        ValidityDays      = $PassValidityDays
        TypeValueParam    = $TypeValue
        CompanyValueParam = $CompanyValue
    }
    Generate-ForAllOrganizationalUnits @params
}
else {
    Start-Gui
}



