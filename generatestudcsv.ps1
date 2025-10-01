# Убедитесь, что установлен модуль ImportExcel:
# Install-Module -Name ImportExcel -Scope CurrentUser -Force

$SearchBase = "OU=2025,OU=zo,OU=SVPO,OU=KRIGT,DC=krsk,DC=irgups,DC=ru"

# Получаем пользователей и сортируем по Department
$Users = Get-ADUser -SearchBase $SearchBase -Filter * -Properties givenName, sn, displayName, description, department |
    Where-Object { $_.Enabled -eq $true } |
    Sort-Object -Property Department

$ExportData = foreach ($user in $Users) {
    $MiddleName = ""
    if ($user.displayName) {
        $parts = $user.displayName -split '\s+'
        if ($parts.Count -ge 3) {
            $MiddleName = $parts[2]
        }
    }

    [PSCustomObject]@{
        Фамилия     = $user.sn
        Имя         = $user.givenName
        Отчество    = $MiddleName
        Description = $user.description
        Department  = $user.department
        Логин       = $user.sAMAccountName
        rawcard     = ""  # пустой текстовый столбец
        codecard    = ""  # пустой текстовый столбец
    }
}

$ExcelPath = "C:\vpo\users_2025_sorted.xlsx"
$ExportData | Export-Excel -Path $ExcelPath -WorksheetName "Пользователи" -AutoSize -ClearSheet

Write-Host "Выгрузка завершена и отсортирована по Department: $ExcelPath"
