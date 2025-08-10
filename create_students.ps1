Param(
    [Parameter(Mandatory=$false)] [ValidateSet('SPO','VO','VOZO')] [string] $Type,
    [Parameter(Mandatory=$false)] [string] $InputCsv,
    [Parameter(Mandatory=$false)] [string] $OutputCsv,
    [Parameter(Mandatory=$false)] [string] $Delimiter = ';',
    [Parameter(Mandatory=$false)] [string] $UpnDomain = 'krsk.irgups.ru',
    [Parameter(Mandatory=$false)] [switch] $UseCsvPass,
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

function Get-PasswordString {
    Param(
        [hashtable] $Row,
        [switch] $PreferCsv
    )
    $passValue = $null
    if ($PreferCsv) {
        if ($Row.ContainsKey('Pass')) {
            $candidate = [string]$Row['Pass']
            if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                $passValue = $candidate
            }
        }
    }
    if (-not $passValue) {
        $upper = -join ((65..90) | Get-Random -Count 1 | ForEach-Object { [char]$_ })
        $lower = -join ((97..122) | Get-Random -Count 5 | ForEach-Object { [char]$_ })
        $digit = -join ((48..57) | Get-Random -Count 1 | ForEach-Object { [char]$_ })
        $passValue = "$upper$lower$digit"
    }
    return $passValue
}

function Get-LoginForType {
    Param(
        [ValidateSet('SPO','VO','VOZO')] [string] $Type,
        [hashtable] $Row
    )
    switch ($Type) {
        'SPO' { return 's' + ([string]$Row['SamA']).Trim() }
        'VO'  { return 'v' + ([string]$Row['SamA']).Trim() }
        'VOZO' { return ([string]$Row['SamA']).Trim() }
    }
}

function Get-StudentNumberForType {
    Param(
        [ValidateSet('SPO','VO','VOZO')] [string] $Type,
        [hashtable] $Row
    )
    switch ($Type) {
        'VOZO' { return ([string]$Row['NumberStud']).Trim() }
        default { return ([string]$Row['SamA']).Trim() }
    }
}

function Get-PathForType {
    Param(
        [ValidateSet('SPO','VO','VOZO')] [string] $Type,
        [string] $Department
    )
    $dept = $Department.Trim()
    switch ($Type) {
        'SPO'  { return "OU=$dept,OU=groups,OU=SSPO,OU=KRIGT,DC=krsk,DC=irgups,DC=ru" }
        'VO'   { return "OU=$dept,OU=groups,OU=SVPO,OU=KRIGT,DC=krsk,DC=irgups,DC=ru" }
        'VOZO' { return "OU=$dept,OU=zo,OU=SVPO,OU=KRIGT,DC=krsk,DC=irgups,DC=ru" }
    }
}

function Ensure-OrganizationalUnitExists {
    Param([string] $OuDn, [ValidateSet('SPO','VO','VOZO')] [string] $Type)

    $exists = $false
    try {
        $exists = [adsi]::Exists("LDAP://$OuDn")
    } catch {
        $exists = $false
    }

    if (-not $exists) {
        switch ($Type) {
            'SPO'  { New-ADOrganizationalUnit -Name (([regex]::Match($OuDn, '^OU=([^,]+),')).Groups[1].Value) -Path 'OU=groups,OU=SSPO,OU=KRIGT,DC=krsk,DC=irgups,DC=ru' -ErrorAction Stop }
            'VO'   { New-ADOrganizationalUnit -Name (([regex]::Match($OuDn, '^OU=([^,]+),')).Groups[1].Value) -Path 'OU=groups,OU=SVPO,OU=KRIGT,DC=krsk,DC=irgups,DC=ru' -ErrorAction Stop }
            'VOZO' { New-ADOrganizationalUnit -Name (([regex]::Match($OuDn, '^OU=([^,]+),')).Groups[1].Value) -Path 'OU=zo,OU=SVPO,OU=KRIGT,DC=krsk,DC=irgups,DC=ru' -ErrorAction Stop }
        }
    }
}

function Build-DisplayName {
    Param([string] $Surname, [string] $GivenName, [string] $MiddleName)
    $given = ($GivenName, $MiddleName) -join ' '
    $given = $given.Trim()
    return ("$Surname $given").Trim()
}

function Convert-DictionaryToObject {
    Param([hashtable] $h)
    $obj = New-Object psobject
    foreach ($k in $h.Keys) { Add-Member -InputObject $obj -NotePropertyName $k -NotePropertyValue $h[$k] }
    return $obj
}

function Process-Students {
    Param(
        [ValidateSet('SPO','VO','VOZO')] [string] $Type,
        [string] $CsvPath,
        [string] $OutCsvPath,
        [string] $Delimiter,
        [string] $UpnDomain,
        [switch] $UseCsvPass,
        [System.Windows.Forms.TextBox] $LogTextBox = $null
    )

    Ensure-ActiveDirectoryModule

    if (-not (Test-Path -LiteralPath $CsvPath)) {
        throw "CSV не найден: $CsvPath"
    }

    $rowsRaw = Import-Csv -Path $CsvPath -Delimiter $Delimiter
    $rows = @()
    foreach ($r in $rowsRaw) {
        # Convert to hashtable for robust key checks
        $h = @{}
        foreach ($p in $r.PSObject.Properties) { $h[$p.Name] = $p.Value }
        $rows += ,$h
    }

    if ($rows.Count -eq 0) {
        throw 'CSV пустой или не содержит строк.'
    }

    $outputRecords = New-Object System.Collections.Generic.List[Object]

    $i = 0
    foreach ($row in $rows) {
        $i++
        $surname = [string]$row['Name1']
        $name2 = [string]$row['Name2']
        $name3 = [string]$row['Name3']
        $given = @($name2, $name3) -join ' '
        $department = [string]$row['Department']

        $login = Get-LoginForType -Type $Type -Row $row
        $tabNumber = Get-StudentNumberForType -Type $Type -Row $row
        $passStr = Get-PasswordString -Row $row -PreferCsv:$UseCsvPass
        $password = ConvertTo-SecureString $passStr -AsPlainText -Force

        $path = Get-PathForType -Type $Type -Department $department
        Ensure-OrganizationalUnitExists -OuDn $path -Type $Type

        $displayName = Build-DisplayName -Surname $surname -GivenName $name2 -MiddleName $name3
        $userPrincipalName = "$login@$UpnDomain"

        $usr = Get-ADUser -Filter "(SamAccountName -eq '$login')" -ErrorAction SilentlyContinue
        if (-not $usr) {
            New-ADUser -Name $displayName `
                       -GivenName $given `
                       -Surname $surname `
                       -SamAccountName $login `
                       -UserPrincipalName $userPrincipalName `
                       -DisplayName $displayName `
                       -AccountPassword $password `
                       -Department $department `
                       -ChangePasswordAtLogon $false `
                       -PasswordNeverExpires $true `
                       -CannotChangePassword $true `
                       -Description $tabNumber `
                       -Path $path | Out-Null
            Enable-ADAccount -Identity $login
            $msg = "Создан пользователь $login"
        }
        else {
            Enable-ADAccount -Identity $login
            $usr | Set-ADAccountPassword -Reset -NewPassword $password -PassThru | Set-ADUser -ChangePasswordAtLogon $false -PasswordNeverExpires $true -CannotChangePassword $true -Description $tabNumber -Department $department -PassThru | Out-Null
            $usr | Move-ADObject -TargetPath $path -ErrorAction SilentlyContinue
            $msg = "Существующий пользователь $login"
        }

        Set-ADUser -Identity $login -Clear info -ErrorAction SilentlyContinue
        Set-ADUser -Identity $login -Replace @{ info = $passStr }

        $rowObj = [pscustomobject]@{
            department     = $department
            Name           = $displayName
            SamAccountName = $login
            info           = $passStr
        }
        [void]$outputRecords.Add($rowObj)

        if ($LogTextBox) { $LogTextBox.AppendText($msg + [Environment]::NewLine) } else { Write-Host $msg }
    }

    $csvLines = $outputRecords | ConvertTo-Csv -NoTypeInformation -Delimiter $Delimiter | ForEach-Object { $_ -replace '"', '' }
    $outDir = Split-Path -Path $OutCsvPath -Parent
    if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -Path $outDir -ItemType Directory | Out-Null }
    $csvLines | Set-Content -LiteralPath $OutCsvPath -Encoding UTF8

    $doneMessage = "Готово. Сохранено: $OutCsvPath"
    if ($LogTextBox) { $LogTextBox.AppendText($doneMessage + [Environment]::NewLine) } else { Write-Host $doneMessage }
}

function Start-Gui {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    Add-Type -AssemblyName System.Drawing | Out-Null

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Создание студентов в AD'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(760, 560)

    $labelWidth = 180
    $inputWidth = 460
    $leftMargin = 16
    $top = 16
    $rowHeight = 28

    function Add-LabelAndTextBox([string] $labelText, [string] $defaultText, [int] $rowIndex) {
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

    $labelType = New-Object System.Windows.Forms.Label
    $labelType.Text = 'Тип списка:'
    $labelType.AutoSize = $true
    $labelType.Location = New-Object System.Drawing.Point($leftMargin, $top)
    $form.Controls.Add($labelType)

    $cbType = New-Object System.Windows.Forms.ComboBox
    $cbType.Items.AddRange(@('SPO','VO','VOZO'))
    $cbType.DropDownStyle = 'DropDownList'
    $cbType.SelectedIndex = 0
    $cbType.Location = New-Object System.Drawing.Point($leftMargin + $labelWidth, $top - 2)
    $form.Controls.Add($cbType)

    $tbCsv = Add-LabelAndTextBox -labelText 'Входной CSV:' -defaultText '' -rowIndex 1
    $tbOut = Add-LabelAndTextBox -labelText 'Выходной CSV:' -defaultText '' -rowIndex 2
    $tbDel = Add-LabelAndTextBox -labelText 'Разделитель:' -defaultText $Delimiter -rowIndex 3
    $tbUpn = Add-LabelAndTextBox -labelText 'Домен (UPN):' -defaultText $UpnDomain -rowIndex 4

    $chkUseCsvPass = New-Object System.Windows.Forms.CheckBox
    $chkUseCsvPass.Text = 'Использовать пароль из CSV (колонка Pass)'
    $chkUseCsvPass.Checked = [bool]$UseCsvPass
    $chkUseCsvPass.Location = New-Object System.Drawing.Point($leftMargin + $labelWidth, $top + 5 * $rowHeight - 6)
    $form.Controls.Add($chkUseCsvPass)

    $btnBrowseIn = New-Object System.Windows.Forms.Button
    $btnBrowseIn.Text = '...'
    $btnBrowseIn.Width = 32
    $btnBrowseIn.Location = New-Object System.Drawing.Point($leftMargin + $labelWidth + $inputWidth + 6, $top + 1 * $rowHeight - 6)
    $btnBrowseIn.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = 'CSV (*.csv)|*.csv|Все файлы (*.*)|*.*'
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $tbCsv.Text = $dlg.FileName }
    })
    $form.Controls.Add($btnBrowseIn)

    $btnBrowseOut = New-Object System.Windows.Forms.Button
    $btnBrowseOut.Text = '...'
    $btnBrowseOut.Width = 32
    $btnBrowseOut.Location = New-Object System.Drawing.Point($leftMargin + $labelWidth + $inputWidth + 6, $top + 2 * $rowHeight - 6)
    $btnBrowseOut.Add_Click({
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter = 'CSV (*.csv)|*.csv|Все файлы (*.*)|*.*'
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $tbOut.Text = $dlg.FileName }
    })
    $form.Controls.Add($btnBrowseOut)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = 'Запустить'
    $btnRun.Width = 160
    $btnRun.Location = New-Object System.Drawing.Point($leftMargin, $top + 7 * $rowHeight)
    $form.Controls.Add($btnRun)

    $tbLog = New-Object System.Windows.Forms.TextBox
    $tbLog.Multiline = $true
    $tbLog.ScrollBars = 'Vertical'
    $tbLog.ReadOnly = $true
    $tbLog.WordWrap = $false
    $tbLog.Location = New-Object System.Drawing.Point($leftMargin, $top + 8.5 * $rowHeight)
    $tbLog.Size = New-Object System.Drawing.Size(710, 360)
    $form.Controls.Add($tbLog)

    $btnRun.Add_Click({
        try {
            $btnRun.Enabled = $false
            $tbLog.AppendText('Старт...' + [Environment]::NewLine)
            if ([string]::IsNullOrWhiteSpace($tbCsv.Text)) { throw 'Укажите входной CSV.' }
            if ([string]::IsNullOrWhiteSpace($tbOut.Text)) { throw 'Укажите выходной CSV.' }
            $params = @{
                Type       = $cbType.SelectedItem
                CsvPath    = $tbCsv.Text
                OutCsvPath = $tbOut.Text
                Delimiter  = $tbDel.Text
                UpnDomain  = $tbUpn.Text
                UseCsvPass = $chkUseCsvPass.Checked
                LogTextBox = $tbLog
            }
            Process-Students @params
        }
        catch {
            $tbLog.AppendText('Ошибка: ' + $_.Exception.Message + [Environment]::NewLine)
        }
        finally {
            $btnRun.Enabled = $true
        }
    })

    $form.Add_Shown({ $form.Activate() })

    $isSta = [System.Threading.Thread]::CurrentThread.ApartmentState -eq 'STA'
    if (-not $isSta) {
        [System.Windows.Forms.MessageBox]::Show('Для GUI требуется запуск PowerShell с параметром -STA.', 'Требуется STA') | Out-Null
    }

    [void] $form.ShowDialog()
}

# CLI / GUI entry
if ($NoGui) {
    if (-not $PSBoundParameters.ContainsKey('Type')) { throw 'Укажите -Type: SPO|VO|VOZO' }
    if (-not $PSBoundParameters.ContainsKey('InputCsv')) { throw 'Укажите -InputCsv' }
    if (-not $PSBoundParameters.ContainsKey('OutputCsv')) { throw 'Укажите -OutputCsv' }
    $procParams = @{
        Type       = $Type
        CsvPath    = $InputCsv
        OutCsvPath = $OutputCsv
        Delimiter  = $Delimiter
        UpnDomain  = $UpnDomain
        UseCsvPass = $UseCsvPass
    }
    Process-Students @procParams
}
else {
    Start-Gui
}