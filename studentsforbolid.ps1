# Убедитесь, что установлен модуль ImportExcel:
# Install-Module -Name ImportExcel -Scope CurrentUser -Force

$ExcelPath = "C:\vpo\users_2025_sorted.xlsx"
$OutputPath = "C:\vpo"
$FileNameBase = "SVPO"

# Настройки
$FixedDepartment = "СВПО ЗО 2025"
$PassType = "Студент без общежития"
$GenderDefault = "Мужчина"

# Создаём папку вывода, если не существует
if (-not (Test-Path -Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath | Out-Null
}

# Проверяем наличие Excel-файла
if (-not (Test-Path $ExcelPath)) {
    throw "Файл не найден: $ExcelPath"
}

# Читаем данные из Excel
Write-Host "Чтение данных из: $ExcelPath"
$Users = Import-Excel -Path $ExcelPath -WorksheetName "Пользователи"

if (-not $Users) {
    throw "В Excel-файле нет данных на листе 'Пользователи'"
}

# Проверка обязательных столбцов
$requiredColumns = 'Фамилия', 'Имя', 'Отчество', 'Логин', 'rawcard'
$missing = $requiredColumns | Where-Object { $_ -notin $Users[0].PSObject.Properties.Name }
if ($missing) {
    throw "В Excel отсутствуют столбцы: $($missing -join ', ')"
}

# ==================================================
# 1. Файл сотрудников (SVPO.csv) — без кавычек, в ANSI
# ==================================================
$HeaderEmployee = @(
    'Name', 'MidName', 'LastName', 'TabNumber', 'Department',
    'WorkPhone', 'HomePhone', 'BirthDate', 'Address', 'Post',
    'Company', "'", 'Status', 'Weight', 'WeightDelta',
    'TypeDocum', 'DocSeries', 'DocNumber', 'INN', 'DocDate',
    'SubDivisionCode', 'WhoIssued', 'DocDateFinish', 'BirthPlace', 'Gender',
    'ReceiveName', 'ReceiveLastName', 'ReceiveMidName', 'CompanyReceive', 'SectionReceive',
    'ReceiveGuestRoom', 'Auto', 'AutoNumber', 'Automarka', 'AutoColor',
    'DateVisit', 'DateVisitFinish', 'VisitGoal', 'VIN', 'BlackList',
    'BlackListReason', 'Fired', 'FireReason', 'IndexForContactId', 'EmailList'
) -join ';'

$EmployeeLines = @($HeaderEmployee)

foreach ($user in $Users) {
    $line = @(
        $user.Имя
        $user.Отчество
        $user.Фамилия
        $user.Логин
        $FixedDepartment
        '' # WorkPhone
        '' # HomePhone
        '' # BirthDate
        '' # Address
        '' # Post
        '' # Company
        '' # Поле "'"
        '5' # Status
        '' # Weight
        '' # WeightDelta
        '' # TypeDocum
        '' # DocSeries
        '' # DocNumber
        '' # INN
        '' # DocDate
        '' # SubDivisionCode
        '' # WhoIssued
        '' # DocDateFinish
        '' # BirthPlace
        $GenderDefault
        '' # ReceiveName
        '' # ReceiveLastName
        '' # ReceiveMidName
        '' # CompanyReceive
        '' # SectionReceive
        '' # ReceiveGuestRoom
        '' # Auto
        '' # AutoNumber
        '' # Automarka
        '' # AutoColor
        '' # DateVisit
        '' # DateVisitFinish
        '' # VisitGoal
        '' # VIN
        '0' # BlackList
        '' # BlackListReason
        '0' # Fired
        '' # FireReason
        '' # IndexForContactId
        '' # EmailList
    ) -join ';'

    $EmployeeLines += $line
}

$EmployeePath = Join-Path $OutputPath "$FileNameBase.csv"
$EmployeeLines | Set-Content -Path $EmployeePath -Encoding Default
Write-Host "✅ Файл сотрудников сохранён (ANSI): $EmployeePath"

# ==================================================
# 2. Файл пропусков (SVPO-pass.csv) — без кавычек, в ANSI
# ==================================================
$StartDate = Get-Date -Format 'dd.MM.yyyy'
$EndDate = (Get-Date).AddDays(1865).ToString('dd.MM.yyyy')

$PassLines = @()
foreach ($user in $Users) {
    $line = @(
        '0'
        '0'
        $user.Логин
        $PassType
        $StartDate
        $EndDate
        $user.rawcard
        '0'
        '0'
        ''
        ''
    ) -join ';'

    $PassLines += $line
}

$PassPath = Join-Path $OutputPath "$FileNameBase-pass.csv"
$PassLines | Set-Content -Path $PassPath -Encoding Default
Write-Host "✅ Файл пропусков сохранён (ANSI):   $PassPath"

Write-Host "`n🎉 Генерация завершена! Файлы в кодировке ANSI (Windows-1251)."
