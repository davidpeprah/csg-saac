##########################################################
#Author     : David Peprah                               #
#                                                        #
#Date       :  10/17/2019                                #
#                                                        #
#Description: For enable users in Active Directory    #
#             using their District Emails                #
##########################################################


$Time= (Get-Date)


function lastFourSSN($lastFourDigitSSN) {
   
   if ($lastFourDigitSSN.length -lt 3) { 
     
       return (get-date).year 
   } else { 
   
       return $lastFourDigitSSN 
   } 
}


try {

# Get email from console
$mail = ($args[0]).ToString()
$lastFourDigitSSN = lastFourSSN $args[1].ToString()


# check account exist
if (get-aduser -Filter {mail -eq $mail}) {
    
    #Search for the Account name in AD using the Email address
    $user_acc = get-aduser -Filter {mail -eq $mail} -Properties SamAccountName | Select -ExpandProperty SamAccountName
   
    #Enable user
    set-ADuser -Identity $user_acc -Enabled $True
    Set-ADAccountPassword -Identity $user_acc -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "milford$lastFourDigitSSN" -Force)
    
    "$Time Account Successfully Activated: $user_acc " | out-file ..\logs\event_log.log -append
    return (6, "Account Successfully Activated")


} else {
    "$Time Account doesn't exist: $mail " | out-file ..\logs\event_log.log -append
    return (2, "Account doesn't exist: $mail")
}
   
} catch {
  
  $ErrorMessage = $_.Exception.Message

  "$Time Password Reset failed $workEmail. Error Message: $ErrorMessage" | out-file ..\logs\event_log.log -append
  return (2, "An error occured while reactivating an account")
  
  Break
}

