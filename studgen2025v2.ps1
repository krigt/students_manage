Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Функция: генерация надёжного пароля ---
function New-RandomPassword {
    param(
        [int]$Length = 9,
        [int]$MinUppercase = 1,
        [int]$MinLowercase = 1,
        [int]$MinDigits = 1
    )

    $upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ'.ToCharArray()
    $lower = 'abcdefghjkmnpqrstuvwxyz'.ToCharArray()
    $digits = '23456789'.ToCharArray()
    $allChars = ($upper + $lower + $digits)
    $password = New-Object System.Collections.Generic.List[Char]

    # Гарантируем минимум
    for ($i = 0; $i -lt $MinUppercase; $i++) { $password.Add($upper[(Get-Random $upper.Length)]) }
    for ($i = 0; $i -lt $MinLowercase; $i++) { $password.Add($lower[(Get-Random $lower.Length)]) }
    for ($i = 0; $i -lt $MinDigits; $i++) { $password.Add($digits[(Get-Random $digits.Length)]) }

    # Дополняем до нужной длины
    for ($i = $password.Count; $i -lt $Length; $i++) {
        $password.Add($allChars[(Get-Random $allChars.Length)])
    }

    # Перемешиваем
    ($password | Sort-Object { Get-Random }) -join ''
}

# Попытка загрузить модуль AD
try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Не удалось загрузить модуль ActiveDirectory.`r`nУстановите RSAT или запустите на контроллере домена.",
        "Ошибка",
        "OK",
        "Error"
    )
    exit
}

# --- Автоопределение домена ---
try {
    $domainInfo = Get-ADDomain
    $domainDNS = $domainInfo.DNSRoot
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Не удалось получить информацию о домене.`r`nПроверьте подключение к сети и вход в домен.",
        "Ошибка домена",
        "OK",
        "Error"
    )
    exit
}

# --- Список доступных OU ---
$availableOUs = @(
    "OU=2025,OU=groups,OU=SVPO,OU=KRIGT,DC=krsk,DC=irgups,DC=ru"
    "CN=Users,$($domainInfo.DistinguishedName)"
    "OU=2025,OU=zo,OU=SVPO,OU=KRIGT,DC=krsk,DC=irgups,DC=ru"
    "OU=2025,OU=groups,OU=SSPO,OU=KRIGT,DC=krsk,DC=irgups,DC=ru"
)

# --- Создание формы ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Создание пользователей в Active Directory"
$form.Size = New-Object System.Drawing.Size(720, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.Icon = [System.Drawing.SystemIcons]::User

# --- Метка: файл ---
$labelFile = New-Object System.Windows.Forms.Label
$labelFile.Location = New-Object System.Drawing.Point(20, 20)
$labelFile.Size = New-Object System.Drawing.Size(500, 23)
$labelFile.Text = "Выберите CSV-файл с пользователями:"
$form.Controls.Add($labelFile)

# --- Поле ввода файла ---
$txtFile = New-Object System.Windows.Forms.TextBox
$txtFile.Location = New-Object System.Drawing.Point(20, 50)
$txtFile.Size = New-Object System.Drawing.Size(500, 23)
$form.Controls.Add($txtFile)

# --- Кнопка "Обзор" ---
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Location = New-Object System.Drawing.Point(530, 50)
$btnBrowse.Size = New-Object System.Drawing.Size(100, 23)
$btnBrowse.Text = "Обзор..."
$btnBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "CSV-файлы (*.csv)|*.csv|Все файлы (*.*)|*.*"
    $dialog.Title = "Выберите CSV-файл"
    if ($dialog.ShowDialog() -eq "OK") {
        $txtFile.Text = $dialog.FileName
        $btnCreate.Enabled = $true
        LoadPreview
    }
})
$form.Controls.Add($btnBrowse)

# --- Метка: выбор OU ---
$labelOU = New-Object System.Windows.Forms.Label
$labelOU.Location = New-Object System.Drawing.Point(20, 80)
$labelOU.Size = New-Object System.Drawing.Size(300, 23)
$labelOU.Text = "Выберите подразделение (OU):"
$form.Controls.Add($labelOU)

# --- Выпадающий список OU ---
$cbOU = New-Object System.Windows.Forms.ComboBox
$cbOU.Location = New-Object System.Drawing.Point(20, 110)
$cbOU.Size = New-Object System.Drawing.Size(660, 23)
$cbOU.DropDownStyle = "DropDownList"
$availableOUs | ForEach-Object { $cbOU.Items.Add($_) }
$cbOU.SelectedIndex = 0
$form.Controls.Add($cbOU)

# --- Метка: предпросмотр ---
$labelPreview = New-Object System.Windows.Forms.Label
$labelPreview.Location = New-Object System.Drawing.Point(20, 140)
$labelPreview.Size = New-Object System.Drawing.Size(500, 23)
$labelPreview.Text = "Предварительный просмотр (первые 5 строк):"
$form.Controls.Add($labelPreview)

# --- Таблица-просмотр ---
$dgvPreview = New-Object System.Windows.Forms.DataGridView
$dgvPreview.Location = New-Object System.Drawing.Point(20, 170)
$dgvPreview.Size = New-Object System.Drawing.Size(660, 150)
$dgvPreview.AutoSizeColumnsMode = "Fill"
$form.Controls.Add($dgvPreview)

# --- Метка: лог ---
$labelLog = New-Object System.Windows.Forms.Label
$labelLog.Location = New-Object System.Drawing.Point(20, 330)
$labelLog.Size = New-Object System.Drawing.Size(500, 23)
$labelLog.Text = "Лог операций:"
$form.Controls.Add($labelLog)

# --- Поле лога ---
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 360)
$txtLog.Size = New-Object System.Drawing.Size(660, 160)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.BackColor = "White"
$form.Controls.Add($txtLog)

# --- Кнопка "Создать" ---
$btnCreate = New-Object System.Windows.Forms.Button
$btnCreate.Location = New-Object System.Drawing.Point(580, 530)
$btnCreate.Size = New-Object System.Drawing.Size(100, 30)
$btnCreate.Text = "Создать"
$btnCreate.Enabled = $false
$form.AcceptButton = $btnCreate
$form.Controls.Add($btnCreate)

# --- Функция: предпросмотр CSV ---
function LoadPreview {
    $dgvPreview.Rows.Clear()
    $dgvPreview.Columns.Clear()
    try {
        $data = Import-Csv -Path $txtFile.Text -Delimiter ";" -Encoding UTF8 | Select-Object -First 5
        if ($data) {
            $data[0].PSObject.Properties | ForEach-Object {
                $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
                $col.Name = $_.Name
                $col.HeaderText = $_.Name
                $dgvPreview.Columns.Add($col)
            }
            foreach ($row in $data) {
                $values = $row.PSObject.Properties | ForEach-Object { $_.Value }
                $dgvPreview.Rows.Add($values)
            }
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка чтения CSV: $($_.Exception.Message)", "Ошибка", "OK", "Error")
    }
}

# --- Функция: запись в лог ---
function Log {
    param([string]$Message)
    $txtLog.AppendText("$(Get-Date -Format 'HH:mm:ss') - $Message`r`n")
    $txtLog.ScrollToCaret()
}

# --- Обработчик кнопки "Создать" ---
$btnCreate.Add_Click({
    $csvPath = $txtFile.Text.Trim()
    if (-not (Test-Path $csvPath)) {
        [System.Windows.Forms.MessageBox]::Show("Файл не найден.", "Ошибка", "OK", "Error")
        return
    }

    try {
        # Читаем CSV и делаем редактируемую копию
        $usersRaw = Get-Content -Path $csvPath -Encoding UTF8 | ConvertFrom-Csv -Delimiter ";"
        $usersWithPasswords = @()

        foreach ($user in $usersRaw) {
            $userObj = $user.PSObject.Copy()
            if (-not $userObj.PSObject.Properties.Match('Password')) {
                Add-Member -InputObject $userObj -MemberType NoteProperty -Name "Password" -Value ""
            }
            $usersWithPasswords += $userObj
        }
    }
    catch {
        Log ("❌ Ошибка чтения CSV: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show("Ошибка чтения CSV: $($_.Exception.Message)", "Ошибка", "OK", "Error")
        return
    }

    if ($usersWithPasswords.Count -eq 0) {
        Log "⚠ CSV пустой."
        [System.Windows.Forms.MessageBox]::Show("CSV пустой.", "Внимание", "OK", "Warning")
        return
    }

    # Логируем заголовки
    $headers = $usersWithPasswords[0].PSObject.Properties.Name -join ", "
    Log ("📋 Загружено. Столбцы: {0}" -f $headers)
    Log ("📊 Найдено записей: {0}" -f $usersWithPasswords.Count)

    $selectedOU = $cbOU.SelectedItem.ToString()

    $result = [System.Windows.Forms.MessageBox]::Show(
        "Создать $($usersWithPasswords.Count) пользователей в:`r`n$selectedOU`r`nПродолжить?",
        "Подтверждение",
        "YesNo",
        "Question"
    )
    if ($result -ne "Yes") { return }

    Log ("🚀 Начинаем создание пользователей в OU: {0}" -f $selectedOU)

    foreach ($user in $usersWithPasswords) {
        # Чтение и очистка полей
        $lastName     = if ($user.LastName)     { $user.LastName.Trim() }     else { $null }
        $firstName    = if ($user.FirstName)    { $user.FirstName.Trim() }    else { $null }
        $middleName   = if ($user.MiddleName)   { $user.MiddleName.Trim() }   else { $null }
        $login        = if ($user.Login)        { $user.Login.Trim() }        else { $null }
        $department   = if ($user.Department)   { $user.Department.Trim() }   else { $null }

        Log ("📄 Читаем строку: Фамилия='{0}', Имя='{1}', Отчество='{2}', Логин='{3}', Подразделение='{4}'" -f `
            $lastName, $firstName, $middleName, $login, $department)

        # Проверка обязательных полей
        if (-not $login) {
            Log "❌ Пропущено: пустой логин"
            continue
        }

        # Проверка существования пользователя
        try {
            Get-ADUser -Identity $login -ErrorAction Stop | Out-Null
            Log ("⚠ Существует: {0}" -f $login)
            continue
        }
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            # Пользователя нет — всё в порядке
        }
        catch {
            Log ("❌ Ошибка проверки {0}: {1}" -f $login, $_.Exception.Message)
            continue
        }

        # Формируем ФИО
        $fullName = $lastName
        if ($firstName)  { $fullName += " $firstName" }
        if ($middleName) { $fullName += " $middleName" }

        # Генерация пароля, если слабый или пустой
        $providedPassword = $user.Password
        if ($providedPassword) { $providedPassword = $providedPassword.Trim() }

        $hasUpper = $providedPassword -cmatch "[A-Z]"
        $hasLower = $providedPassword -cmatch "[a-z]"
        $hasDigit = $providedPassword -match "\d"
        $isStrong = $providedPassword -and $providedPassword.Length -ge 8 -and $hasUpper -and $hasLower -and $hasDigit

        if ($isStrong) {
            $password = $providedPassword
            Log "🔐 Используем пароль из CSV для $login"
        } else {
            $password = New-RandomPassword -Length 10
            $user.Password = $password  # Обновляем в объекте
            Log "🔐 Сгенерирован пароль для ${login}: $password"
        }

        # Description = логин, info = пароль
        $description = $login
        $notes = $password

        # Создание пользователя
        try {
            New-ADUser `
                -Name $fullName `
                -GivenName $firstName `
                -Surname $lastName `
                -DisplayName $fullName `
                -SamAccountName $login `
                -Path $selectedOU `
                -Department $department `
                -Description $description `
                -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) `
                -Enabled $true `
                -ChangePasswordAtLogon $false `
                -CannotChangePassword $true `
                -PasswordNeverExpires $true `
                -OtherAttributes @{
                    "info" = $notes
                }

            Log ("✅ Создан: {0} ({1})" -f $login, $fullName)
        }
        catch {
            Log ("❌ Ошибка при создании {0}: {1}" -f $login, $_.Exception.Message)
        }
    }

    # --- Сохранение обновлённого CSV с паролями ---
    try {
        $outputCsv = $usersWithPasswords | ConvertTo-Csv -Delimiter ";" -NoTypeInformation
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllLines($csvPath, $outputCsv, $utf8NoBom)
        Log "💾 Обновлённый CSV сохранён с паролями: $csvPath"
    }
    catch {
        Log "❌ Не удалось сохранить CSV: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Не удалось сохранить CSV с паролями.", "Ошибка", "OK", "Error")
    }

    [System.Windows.Forms.MessageBox]::Show("Готово! Проверьте лог и обновлённый CSV.", "Завершено", "OK", "Information")
})

# --- Запуск формы ---
$form.ShowDialog() | Out-Null