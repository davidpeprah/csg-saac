<#

   Author: David Peprah
    
   NOTE: Most of the modification should be done here in this script. Different institutions have different
            Active Directory structure and policies in place when it comes to creating staff account.
            You can add new functions and logic and remove any objects not needed for your environment
#>

# Parameters
param (
    [string]$FirstName,
    [string]$MiddleName,
    [string]$LastName,
    [string]$jobrole,
    [string]$department,
    [string]$adgroups,
    [string]$oupath,
    [string]$jobtitle,
    [string]$testing = "$false"
)


import-Module ActiveDirectory


$Time= (Get-Date)


<#
 This return the username to be used after checking the AD to make sure
 its not being used already by another staff
#>
function Get-SamAccountNm {
    param (
        [string]$lastName,
        [string]$middleName,
        [string]$firstName
    )
    # remove all white spaces in the last name and first name
    $lastName = $lastName.Trim().Replace(" ", "")
  
    $proposeAccName = ($firstName[0] + $lastName).ToLower()

    # this to make sure the samAccountName is not more than 21 characters
  
    if (-Not(get-aduser -Filter {SamAccountName -eq $proposeAccName})) {
        return $proposeAccName
    } 
    else {

    if ($MiddleName) {
        $proposeAccName = ($firstName[0] + $MiddleName[0] + $lastName).ToLower()
        
        if (-Not(get-aduser -Filter {SamAccountName -eq $proposeAccName})){
            return $proposeAccName
        } 
     
    }


    }
    return $false
}

<#
This function returns OU path to be used to create a user.
#>
function check-OUpath {
    param (
        [string]$oupath
    )

    $OU = Get-ADOrganizationalUnit -Filter {Name -eq $oupath} | Select -ExpandProperty DistinguishedName
    if (-not $OU) {
        return $false
    }
    return $OU 
}

function checkGrp{
    param (
        [string]$grpname
    )

   try {
        if ($grpname.endswith("@columbusschoolforgirls.org")) {
            try {
                get-adgroup -Filter "mail -eq '$grpname'"
                return $true
            } catch {
                return $false
            }    
        } 
     Get-ADGroup $grpname
        return $true
   } catch {
     return $false
   }

}

function Get-fullname {
    param (
        [string]$firstName,
        [string]$MiddleName,
        [string]$lastName
    )
    if ($MiddleName) {
        return "$firstName $MiddleName[0] $lastName"
    }
   return "$firstName $lastName"
}

try{

    <# 
    Assign the various values from python to their descriptive variable
    Remove all white spaces in lastname and firstname and social security
    Remove leading and trailing spaces from building, position, and stafftype
    #>
    $SamAccountName = Get-SamAccountNm -firstName $FirstName -middleName $MiddleName -lastName $LastName
    if (-not $SamAccountName) {
        "$Time Could not create an account for $FirstName $MiddleName $LastName. The proposed username is not available" | out-file logs\event_log.log -append
        return (2, ' ', "Could not create account, $Time Could not create an account for $FirstName $MiddleName $LastName. The proposed username is not available")
    }
    
    $emailAddress = "$SamAccountName@columbusschoolforgirls.org"
    $fullName = Get-fullname -firstName $FirstName -middleName $MiddleName -lastName $LastName
    
    $oupath = check-OUpath -oupath $oupath
    if (-not $oupath) {
        "$Time The OU path $oupath does not exist in Active Directory. Please check with your AD Administrator" | out-file logs\event_log.log -append
        return (2, $emailAddress, "Could not create account, $Time The OU path $oupath does not exist in Active Directory. Please check with your AD Administrator")
    }

    $middleInitial = ''
    if ($MiddleName) {
        $middleInitial = $MiddleName[0].ToUpper()
    }

    $password = ("P@ssw0rd@!!").ToString() # This password will change once the account is confirmed in Google console
    $description = $jobtitle
    $userPrincipalName = "$SamAccountName@columbusschoolforgirls.org"
    
    $homeDirectory = "\\csgfs01\administration\$SamAccountName"
    if ($jobrole -eq "faculty") {
        $homeDirectory = "\\csgfs01\faculty\$SamAccountName"
    }

    # Groups
    $ADgrps = $adgroups.split(",")

    #"$Time $fullName, $password, $SamAccountName, $userPrincipalName, $building, $department, $path, " | out-file logs\event_log.log -append
    # Create User Account
    if ($testing -eq "$true") {

        "$Time Testing mode is enabled. No changes will be made to Active Directory" | out-file logs\event_log.log -append
        New-ADUser -Name $fullName -GivenName $FirstName -Surname $LastName -DisplayName $fullName `
        -initials $middleInitial -AccountPassword (ConvertTo-SecureString -AsPlainText $password -Force) `
        -SamAccountName $SamAccountName -UserPrincipalName $userPrincipalName -HomeDrive "H:" -HomeDirectory $homeDirectory `
        -Path $oupath -EmailAddress $emailAddress -Description $description -Company "Columbus School For Girls" `
        -Department $department -Title $jobtitle -PasswordNeverExpires $True -Enabled $True -WhatIf

        $unknowngroups = @()
        # Add user account to default Group
        forEach ($grp in $ADgrps) {
    
            if (-not (checkGrp $grp)) {
                $unknowngroups += $grp
            }
        }   
        return (1, $emailAddress, "Testing mode is enabled. $? details information: $fullname, $SamAccountName, $userPrincipalName, $homeDirectory, $oupath, $emailAddress, $description, $department, $jobtitle, $ADgrps, Unknown Groups: $unknowngroups")
    }

    New-ADUser -Name $fullName -GivenName $FirstName -Surname $LastName -DisplayName $fullName `
    -initials $middleInitial -AccountPassword (ConvertTo-SecureString -AsPlainText $password -Force) `
    -SamAccountName $SamAccountName -UserPrincipalName $userPrincipalName -HomeDrive "H:" -HomeDirectory $homeDirectory `
    -Path $oupath -EmailAddress $emailAddress -Description $description -Company "Columbus School For Girls" `
    -Department $department -Title $jobtitle -PasswordNeverExpires $True -Enabled $True


    $unknowngroups = @()
    # Add user account to default Group
    forEach ($grp in $ADgrps) {
    
    if (checkGrp $grp) {
        
            Add-ADGroupMember $grp -members $SamAccountName
        
        } else {
            $unknowngroups += $grp
        }
    }

    "$Time Account was successfully created for $fullName" | out-file logs\event_log.log -append

    if ($unknowngroups.Count -gt 0) {
        $unknowngroupslist = $unknowngroups -join ","
        "$Time However, the following groups do not exist in Active Directory: $unknowngroupslist. Please check with your AD Administrator" | out-file logs\event_log.log -append
        return (1, $emailAddress, "Account Successfully Created in Active Directory but the following groups do not exist in Active Directory: $unknowngroupslist. Please check with your AD Administrator")
    }
        # Return this information to python
    
    return (1, $emailAddress, "Account Successfully Created in Active Directory")

} catch {

  $ErrorMessage = $_.Exception.Message
  $FailedItem = $_.Exception.ItemName

  "$Time An error occured when trying to create an account for $firstname $lastname. Error Message: $ErrorMessage" | out-file logs\event_log.log -append
  $FailedItem | out-file logs\event_log.log -append
  
  # Return this information to python
  return (2, '', "Could not create account,An error occured when trying to create an account for $firstname $lastname. Error Message: $ErrorMessage, Failed Item: $FailedItem")
  
}
 





