cls
#chcp 1251
Import-Module ActiveDirectory
$filepath = 'D:\work\Институт\списки студентов\ВО ЗО\2021'
$output_file = "$filepath\1 курс 2021 passwords.csv"
$Users = Import-CSV "$filepath\1курс_ВО_ЗО_2021.csv" –Delimiter ";"

Foreach($CurrentUser in $Users) {

$GivenName = $CurrentUser.Name2+" "+$CurrentUser.Name3 #Имя - в данном случае "Имя Отчество"

$Surname = $CurrentUser.Name1 #Фамилия

$NumZK = $CurrentUser.NumberStud #номер зачётки

$pass_str = $CurrentUser.Pass
#

if ($pass_str -like $null)
    {
    $pass_str = get-random -count 1 -input (65..90) | % -begin { $pass = $null } -process {$pass += [char]$_} -end {$pass}
    $pass_str += get-random -count 5 -input (97..122) | % -begin { $pass = $null } -process {$pass += [char]$_} -end {$pass}
    $pass_str += get-random -count 1 -input (48..57) | % -begin { $pass = $null } -process {$pass += [char]$_} -end {$pass}
    }

$password = ConvertTo-SecureString $pass_str -AsPlainText -Force #пароль из файла

$Login = $CurrentUser.SamA #логин юзера из файла

$Department = $CurrentUser.Department #"Отдел" - группа студента

$Path = "OU=$Department" #путь UO=Гру ппа

$Path += ",OU=zo,OU=SVPO,OU=KRIGT,DC=krsk,DC=irgups,DC=ru" #полный путь в целевой OU для создания в нём юзера

$Displayname = $Surname + " " + $GivenName  #в кавычках — пробел! Отображаемое имя пользователя - ФИО полностью

$UserPrincipalName = $Login + "@krsk.irgups.ru" #полное имя для авторизации - логин@домен

#проверка существования AD Organizational Unit
#$OU = Get-ADOrganizationalUnit  -Filter * | Where-Object -FilterScript {$PSItem.distinguishedname -like $Path} 
$ou_exists = [adsi]::Exists("LDAP://$Path")
# -Identity $Path #-SearchBase 'OU=zo,OU=SVPO,OU=KRIGT,DC=krsk,DC=irgups,DC=ru'

if ($ou_exists -eq $false)
    {
    #создание OU группы, игнорирование ошибки, если OU уже создан - продолжение выполнения скрипта
    New-ADOrganizationalUnit -Name $Department -Path "OU=zo,OU=SVPO,OU=KRIGT,DC=krsk,DC=irgups,DC=ru"
    }
    else
    {}
#проверка существования аккаунта в AD
#получаем учётку из AD (если она существут, значение переменной $USR будет отличаться от null)
#$USR = Get-ADUser -Filter (SamAccountName -eq $Login)
$USR = Get-ADUser -filter {(SamAccountName -eq $Login)}
    if ($USR -eq $null)
        {
        #собственно создание юзера, когда все нужные параметры определены
        New-ADUser -Name $Displayname -GivenName $GivenName -Surname $Surname –SamAccountName $Login –UserPrincipalName $UserPrincipalName -DisplayName $DisplayName -AccountPassword $Password -Department $Department -ChangePasswordAtLogon $false -PasswordNeverExpires $true -CannotChangePassword $true -Description $NumZK -Path $Path  #-HomeDrive "Z:" -HomeDirectory "\\sr1\$Login"
        
        #включение созданной учётки
        Enable-ADAccount $Login
        echo "Создан пользователь $Login"
        
    #Add-ADGroupMember "KRIGT-Internet-Restricted" $Login
    #$pathfolder = "d:\profiles\$Department\$Login"
    #$folder = New-Item -Path $pathfolder  -ItemType "directory"
        }
    else
        {
        Enable-ADAccount $Login
        $USR | Set-ADAccountPassword -Reset -NewPassword $password -PassThru | Set-ADuser -ChangePasswordAtLogon $false -PasswordNeverExpires $true -CannotChangePassword $true -Description $NumZK -Department $Department -PassThru
        $USR | Move-ADObject -TargetPath $Path
        echo '----------------------'
        }

Set-ADUser $Login -Clear info
Set-ADUser $Login -Replace @{info = $pass_str} #внесение в поле "зметки" пароля пользователя
echo $Login
echo $USR
echo '***************************************'
Get-ADUser $Login -Properties info,department |  Select  department,Name,SamAccountName,info |  Export-Csv $output_file -NoTypeInformation -Encoding UTF8 -Append -Delimiter ";" | % {$_ -replace '"',''} 

#создание шары

#net share $Login /delete
#net share $Login=$pathfolder /grant:$Login`,FULL /Users:1 /Remark:"Сетевой диск студента"
}
