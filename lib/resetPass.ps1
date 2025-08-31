#
#     Author: David Peprah
#

pram (
    [string]$Email,
    [string]$NewPassword
)


import-Module ActiveDirectory
$Time= (Get-Date)

Try {
$workEmail = $Email
$password = $NewPassword

$user = get-aduser -Filter {mail -eq $workEmail} -Properties SamAccountName | select -ExpandProperty SamAccountName


Set-ADAccountPassword -Identity $user -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $password -Force)
"$Time Password Reset Successful for $workEmail" | out-file logs\event_log.log -append
  
  return ("0", $user)

} catch {
   $ErrorMessage = $_.Exception.Message

  "$Time Password Reset failed $workEmail. Error Message: $ErrorMessage" | out-file logs\event_log.log -append
  
  
  # Return this information to python
  return ("1", $Time Password Reset failed $workEmail. Error Message: $ErrorMessage)
  
  Break
}
