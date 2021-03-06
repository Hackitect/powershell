#Powershell script to retrieve users in O365
#writen by Charles M. Njenga MIS Officer Danish Refugee Council
#Modified 29-04-2015

# The variables which change per location are domain name which are; these are held in the variable name $HOA-Domain
#ddgsom.org
#drcdjibouti.org
#drcethiopia.org
#drchoa.org
#drckenya.org
#drcsomalia.org
#drcyemen.org
#ddghoa.org
#ddgsomalia.org 
$a = (get-date).day
$a = (get-date).dayofweek
$a = (get-date).dayofyear
$a = (get-date).hour
$a = (get-date).millisecond
$a = (get-date).minute
$a = (get-date).month
$a = (get-date).second
$a = (get-date).timeofday
$a = (get-date).year
#initialize domain arrays
#$Fieldstofetch = @(Country, DisplayName, UserPrincipalName, Title, Department, State, Office, PhoneNumber, MobilePhone, LastPasswordChangeTimestamp, FirstName, LastName)
$array = @("drckenya.org", "drchoa.org", "drcsomalia.org", "drcethiopia.org") #you can add more domains here

#$DomainHOA[7] = "ddgyemen.org"

#Construct path to save the output
$myDate = (date).tostring('dd')+"-"+(date).tostring('yyyy')
$FileName2Exportall = "$env:USERPROFILE\Desktop\drckenya.all"+$myDate+".csv"

#Sign in to get authenticated
$LiveCred = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection

Import-PSSession $Session -AllowClobber
Import-Module MsOnline
Connect-MSolservice

#Get specific fields about the users which are the most important



foreach ($element in $array) { 
        $FileName2Export = "$env:USERPROFILE\Desktop\$element"+$myDate+".csv"
        Get-MsolUser -DomainName $element | Select-Object Country, DisplayName, FirstName, LastName, UserPrincipalName, Title, Department, State, Office, PhoneNumber, MobilePhone, LastPasswordChangeTimestamp  | Export-Csv $FileName2Export
}
#Or get All field for that users - uncomment the below line
#Get-MsolUser -DomainName drckenya.org | Export-Csv $FileName2Exportall

#finally unload the user variables and destroy session

#Remove-Variable {array}
Remove-PSSession $Session
