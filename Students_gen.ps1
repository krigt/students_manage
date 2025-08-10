
$group=Get-ADOrganizationalUnit -Filter  {(Name -notlike "groups")} -SearchBase 'OU=groups,OU=SVPO,OU=KRIGT,DC=krsk,DC=irgups,DC=ru' -Properties DistinguishedName,Name| select-object DistinguishedName,Name
$date=Get-Date -Format 'dd.MM.yyyy'
#Дата окончания срока карточки 1865=5лет
$end_date= $(get-date).AddDays(1865).ToString("dd.MM.yyyy")
#Генерирует файлы с пользователями и пролями
foreach($groupname in $group)
{

$filename=$groupname.Name.replace('.','-')
    $users=Get-ADUser -SearchBase $groupname.DistinguishedName -Filter * -Properties givenName,sn,displayName,sAMAccountName,Department | Select-Object  @{
    Name = 'Name'
    Expression = {
        ($_.givenName.split()[0])
    }},
     @{ Name = 'MidName'
    Expression = {($_.displayName.split()[-1])}},
    @{
    Name = 'LastName'
    Expression = {
        ($_.sn)
    }},
    @{
    Name = 'TabNumber'
    Expression = {
        ($_.sAMAccountName)
    }},Department,
    @{
    Name = 'WorkPhone'
    Expression = {""""}},
    @{
    Name = 'HomePhone'
    Expression = {""""}},
    @{
    Name = 'BirthDate'
    Expression = {""""}},
    @{
    Name = 'Address'
    Expression = {""""}},
    @{
    Name = 'Post'
    Expression = {""""}},
    @{
    Name = 'Company'
    Expression = {"SVPO"}},
    @{
    Name = "'"
    Expression = {"Нет"}},
    @{
    Name = "Status"
   Expression = {"Хозорган"}
}

$users | convertto-csv -NoTypeInformation -Delimiter ";"|% { $_ -replace '"', ""} |out-file  C:\vpo\$filename.csv -Encoding default
$users |select-Object  -property  @{
    Name = 'blan'
    Expression = {""""}
},@{
    Name = 'blank2'
    Expression = {""""}
}, TabNumber,
@{
    Name = 'type'
    Expression = { "институт" }
}, 
@{
    Name = 'start_date'
    Expression = {
        ($date) 
    }},
    @{
    Name = 'end_date'
    Expression = {
        ($end_date) 
    }},
     @{
    Name = 'blank3'
    Expression = {""""}} | convertto-csv -NoTypeInformation -Delimiter ";" |% { $_ -replace '"', ""} | out-file  C:\vpo\$filename-pass.csv -Encoding default 
 }



