#
#     Author: David Peprah
#



import-Module ActiveDirectory
$Time= (Get-Date)

Try {

$workEmail = $args[0].ToString()
$password = $args[1].ToString()

$user = get-aduser -Filter {mail -eq $workEmail} -Properties SamAccountName | select -ExpandProperty SamAccountName


Set-ADAccountPassword -Identity $user -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $password -Force)
"$Time Password Reset Successful for $workEmail" | out-file logs\event_log.log -append
  
  return ("0", $user)

} catch {
   $ErrorMessage = $_.Exception.Message

  "$Time Password Reset failed $workEmail. Error Message: $ErrorMessage" | out-file logs\event_log.log -append
  
  
  # Return this information to python
  return ("1", $user)
  
  Break
}
