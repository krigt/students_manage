cls
Import-Module ActiveDirectory
#путь
$filepath = "D:\work\Институт\списки студентов\ВО очное\2021"
#файл для выгрузки логинов и паролей
$output_file = "$filepath\VO_login-password_2021.csv"
#указание, откуда брать данные
$Users = Import-CSV "$filepath\2021.csv" –Delimiter ";"

Foreach($CurrentUser in $Users) {

$GivenName = $CurrentUser.Name2+" "+$CurrentUser.Name3 #Имя - в данном случае "Имя Отчество"
$Surname = $CurrentUser.Name1 #Фамилия
$NumZK = $CurrentUser.SamA #номер зачётки
#$pass_str = $CurrentUser.Pass
#
#$range1 = 65..90 | Where-Object { 73 -notcontains $_ }
#$range2 = 97..122 | Where-Object { 108 -notcontains $_ }
#
$pass_str = $null
#if ($pass_str -like $null) {
    #$pass_str = get-random -count 1 -input $range1 | % -begin { $pass = $null } -process {$pass += [char]$_} -end {$pass}
    $pass_str = get-random -count 1 -input (65..90) | % -begin { $pass = $null } -process {$pass += [char]$_} -end {$pass}
    #$pass_str += get-random -count 5 -input $range2 | % -begin { $pass = $null } -process {$pass += [char]$_} -end {$pass}
    $pass_str += get-random -count 5 -input (97..122) | % -begin { $pass = $null } -process {$pass += [char]$_} -end {$pass}
    $pass_str += get-random -count 1 -input (0..9)
#    }

$password = ConvertTo-SecureString $pass_str -AsPlainText -Force #пароль из файла
$Login = "v"+$NumZK #логин юзера - латинская V + номер зачётки/студика без пробела
$Department = $CurrentUser.Department #"Отдел" - группа студента
$Path = "OU=$Department" #путь UO=Группа
$Path += ",OU=groups,OU=SVPO,OU=KRIGT,DC=krsk,DC=irgups,DC=ru" #полный путь в целевой OU для создания в нём юзера ВО
$Displayname = $Surname + " " + $GivenName  #в кавычках — пробел! Отображаемое имя пользователя - ФИО полностью
$UserPrincipalName = $Login + "@krsk.irgups.ru" #полное имя для авторизации - логин@домен
$OU = Get-ADOrganizationalUnit  -Filter * | Where-Object -FilterScript {$PSItem.distinguishedname -like $Path}
if ($OU -eq $null)
    {
    #создание OU группы, игнорирование ошибки, если OU уже создан - продолжение выполнения скрипта
    New-ADOrganizationalUnit -Name $Department -Path "OU=groups,OU=SVPO,OU=KRIGT,DC=krsk,DC=irgups,DC=ru"
    }
    else
    {}
#проверка существования аккаунта в AD
#получаем учётку из AD (если она существут, значение переменной $USR будет отличаться от null)
$USR = Get-ADUser -filter {(SamAccountName -eq $Login)}
    if ($USR -eq $null)
        {
        #собственно создание юзера, когда все нужные параметры определены
        New-ADUser -Name $Displayname -GivenName $GivenName -Surname $Surname –SamAccountName $Login –UserPrincipalName $UserPrincipalName -DisplayName $DisplayName -AccountPassword $Password -Department $Department -ChangePasswordAtLogon $false -PasswordNeverExpires $true -CannotChangePassword $true -Description $NumZK -Path $Path  #-HomeDrive "Z:" -HomeDirectory "\\sr1\$Login"
        
        #включение созданной учётки
        Enable-ADAccount $Login
        echo "Создан пользователь $Login"
        }
    else
        {
        Enable-ADAccount $Login
        $USR | Set-ADAccountPassword -Reset -NewPassword $password -PassThru | Set-ADuser -ChangePasswordAtLogon $false -PasswordNeverExpires $true -CannotChangePassword $true -Description $NumZK -Department $Department -PassThru
        $USR | Move-ADObject -TargetPath $Path
        echo "Существующий пользователь $Login"
        }

Set-ADUser $Login -Clear info
Set-ADUser $Login -Replace @{info = $pass_str} #внесение в поле "зметки" пароля пользователя
echo $Login
echo $USR
echo '***************************************'
Get-ADUser $Login -Properties info,department |  Select  department,Name,SamAccountName,info |  Export-Csv $output_file -NoTypeInformation -Encoding UTF8 -Append -Delimiter ";" | % {$_ -replace '"',''} 
}
