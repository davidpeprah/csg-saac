#
#     Author: David Peprah
#

param (
    [string]$Email,
    [string]$NewPassword,
    [string]$testing = "false"
)


import-Module ActiveDirectory
$Time= (Get-Date)

Try {
$workEmail = $Email.ToLower().Trim()

$user = get-aduser -Filter {mail -eq $workEmail} -Properties SamAccountName | select -ExpandProperty SamAccountName

if ($testing -eq "true") {
    "$Time Testing mode enabled. No changes will be made to the AD." | out-file logs\event_log.log -append
    Set-ADAccountPassword -Identity $user -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $NewPassword -Force) -WhatIf
    return ("0", "Testing mode enabled. No changes will be made to the AD.")
}

  Set-ADAccountPassword -Identity $user -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $NewPassword -Force) -WhatIf
  "$Time Password Reset Successful for $workEmail" | out-file logs\event_log.log -append
  
  return ("0", $user)

} catch {
  $ErrorMessage = $_.Exception.Message

  "$Time Password Reset failed $workEmail. Error Message: $ErrorMessage" | out-file logs\event_log.log -append
  
  
  # Return this information to python
  return ("1", "$Time Password Reset failed $workEmail. Error Message: $ErrorMessage")
}
