Param(
    [string]$action,
    [string]$file
)


# function for display help
function displayHelp {
    
    "Help for adManage`n"
    "Parameters:(`n"
        "`t -action: specify action to do  **required"
            "`t `t possibles values: 'createUsers','updateUsers','deactivateUser','updateOu','electManager'`n"
        "`t -file: specify path file to use for operate your action   **required`n"
        "`t no parameter will print this help`n"
    "`n)`n"

    "Examples:(`n"
        "`t- Create users with ad_manage"
        "`t `t./ad_manage.ps1 -action createUsers -file path_to_users_to_create.csv`n"

        "`t- Update information of users with ad_manage"
        "`t `t ./ad_manage.ps1 -action updateUsers -file path_to_users_to_update.csv`n"

        "`t- Deactivate users with ad_manage"
        "`t `t ./ad_manage.ps1 -action deactivateUser -file path_to_users_to_deactivate.csv`n"

        "`t- Update Organizational Unit with ad_manage"
        "`t `t ./ad_manage.ps1 -action updateOU -file path_to_ou_to_update.csv`n"

        "`t- Update Organizational Unit with ad_manage"
        "`t `t ./ad_manage.ps1 -action electManager -file path_to_users_to_promote.csv`n"

    "`n)"
}

## export variables 
$domainName = "biodevops.eu"
$company = "BioDevops"

# utilities

# convert domain name to path active directory
function convertDomainNameToPath {
    param (
        [Parameter(Mandatory = $false)]
        [string] $domain = $domainName
    )

    $formatedString = ""

    foreach ($substring in $domain.Split(".")) {
        <# $string iubss the cuf$formatedString #>
        if ($substring -eq $domain.Split(".")[-1]) {
            <# Action to perform if the condition is true #>
            $formatedString += "DC=$substring"
        }
        else {
            $formatedString += "DC=$substring,"
        }
    }

    return $formatedString
}


# load data from csv files.
function loadData {
    param( 
         [Parameter(Mandatory = $true)]
         [string] $pathToCsv,
         [Parameter(Mandatory = $true)]
         [string] $delimiter
    )

    $data = Import-Csv -Path $pathToCsv  -Delimiter "$delimiter"

    return $data
}

function getIdentifiant{
    param (
        [Parameter(Mandatory = $true)]
        [string] $year
    )
    #tant que cet identifiant existe, donner un nouveau identifiant
    do
    {
       $rand = Get-Random -Minimum 0000000001 -Maximum 9999999999
       $samaccountname = "U" + $year + $rand
       $usersAD = Get-ADUser -Filter "samaccountname -eq '$samaccountname'" | Select-Object samaccountname
    }while($usersAD)

    return $samaccountname
}


# check if OU already exists
function isExistsOU {
    param (
        [Parameter(Mandatory = $true)]
        [string] $name
    )

    $filter = Get-ADOrganizationalUnit -filter "Name -eq '$name'"

    if ($filter) {
        <# Action to perform if the condition is true #>
        return $true
    }
    
    return $false
}


# check if user already exists
function isExistsUser {
    param (
        [Parameter(Mandatory = $true)]
        [string] $name
    )
    
    try {
        $filter = Get-ADUser -Identity $name
        if ($filter) {
            <# Action to perform if the condition is true #>
            return $true
        }
    }
    catch {
        return $false
    }
    
    return $false
}


# create organizational unit
function createOU {
    param (
        [Parameter(Mandatory = $true)]
        [string] $promotion_acronym,
        [Parameter(Mandatory = $true)]
        [string] $year
    )

    $name = $year + "_" + $promotion_acronym.ToUpper()
    $domainNamePath = convertDomainNameToPath
    $path = "OU= $name,$domainNamePath"
    $check = isExistsOU -name $name
    if ($check -eq $false) {
        New-ADOrganizationalUnit -Name $name 
        New-ADGroup -Name $name -Path $path `
                    -GroupScope Global -GroupCategory Security
    
        #ajout du nom du groupe de distribution
        $distrib = $name + "_distrib"
        $mail = $promotion_acronym + "." + $year + "@$domainName"
    
        New-ADGroup -name $distrib -Path $path `
                    -GroupCategory Distribution -GroupScope Global `
                    -OtherAttributes @{'mail'= $mail.ToLower()}    
    }
}


# format email and common name
function formEmailAndCommonName {
    param (
        [Parameter(Mandatory = $true)]
        [string] $name,
        [Parameter(Mandatory = $true)]
        [string] $surname
    )

    $email = $($surname + "." + $name + "@" + $domainName).Trim().ToLower().replace(' ','')
    $cname = $surname + " " + $name
    $filter = Get-ADUser -Filter "mail -eq '$email'"
    $cpt = 0
    if ($filter) {
        <# Action to perform if the condition is true #>  
        do {
            $cpt = $cpt + 1
            $email = $($surname + "." + $name + $($cpt) + "@" + $domainName).Trim().ToLower().replace(' ','')
            $cname = $surname + " " + $name + $($cpt)
            $filter = Get-ADUser -Filter "mail -eq '$email'"
            # $filter
            
        } while (
             $filter
        )
    }

    return @($email,$cname)
    
}


# create users from csv
function createUsersFromCsv {
    param( 
         [Parameter(Mandatory = $true)]
         [string] $pathToCsv
    )

    try {
            $users = loadData -pathToCsv $pathToCsv -delimiter ";"

            foreach ($user in $users) {
                <# $user is the current item #>
                $samaccountname = getIdentifiant -year $user.annee
                $emailAndCommonName = formEmailAndCommonName -name $user.nom -surname $user.prenom
                $display = $user.prenom +" "+ $user.nom
                $promotion = $user.annee + "_" + $user.formation
                $path = "OU="+ $promotion + "," + $(convertDomainNameToPath -domain $domainName)
                $distrib = $promotion + "_distrib"
                $manager = (Get-ADOrganizationalUnit -Filter "distinguishedname -eq '$path'" -Properties managedby | `
                Select-Object managedby).managedby     
                
                createOU -promotion_acronym $user.formation -year $user.annee

                New-ADUser  -Name $emailAndCommonName[1] `
                             -GivenName $user.prenom `
                             -Surname $user.nom `
                             -SamAccountName $samaccountname `
                             -DisplayName $display `
                             -Department "$($user.annee)$($user.formation)" `
                             -Title "Student" `
                             -Company $company `
                             -EmailAddress $emailAndCommonName[0] `
                             -UserPrincipalName $emailAndCommonName[0] `
                             -Path $path `
                             -AccountPassword(ConvertTo-SecureString "motdepass@1234698Agzgz" -AsPlainText -Force) `
                             -ChangePasswordAtLogon $true `
                             -Enabled $true `
                             -Manager $manager `
                        
            
                Add-ADGroupMember -Identity $promotion -Members $samaccountname
                Add-ADGroupMember -Identity $distrib -Members $samaccountname  
            }
    }
    catch [System.SystemException] {
        "An error occurred that could not be resolved."
        Write-Host $_
    }
    
    Write-Information "Users creation done !!"
}


# update promotion of users
function updateUserPromotion {
    param (
        [Parameter(Mandatory = $true)]
        [string] $old,
        [Parameter(Mandatory = $true)]
        [string] $new,
        [Parameter(Mandatory = $true)]
        [string] $login
    )
    #recuperation du user dans l'AD a travers son login 
    $user = Get-ADUser -Identity $login -Properties *
    $path = "OU="+ $new + "," + $(convertDomainNameToPath -domain $domainName)
    $old_distribution_group = $old + "_distrib"
    $new_distribution_group = $new + "_distrib"

    $new_manager = (Get-ADOrganizationalUnit -Filter "distinguishedname -eq '$path'" -Properties managedby | `
    Select-Object managedby).managedby

    if ($new_manager) {
        <# Action to perform if the condition is true #>
        Set-ADUser -Identity $login -Manager $new_manager
    }
    else {
        Set-ADUser -Identity $login -clear Manager
    }

    Set-ADUser -Identity $login -Department $new.Replace("_","")
    Move-ADObject -Identity $user.DistinguishedName -TargetPath $path
    Remove-ADGroupMember -Identity $old -Members $user.samaccountname -Confirm:$false
    Remove-ADGroupMember -Identity $old_distribution_group -Members $user.samaccountname -Confirm:$false

    Add-ADGroupMember -Identity $new -Members $login
    Add-ADGroupMember -Identity $new_distribution_group -Members $login
}


# update name of promotion
function updatePromotion {
    param (
        [Parameter(Mandatory = $true)]
        [string] $old,
        [Parameter(Mandatory = $true)]
        [string] $new
    )

    Get-ADOrganizationalUnit -Filter "name -eq '$old'" `
    | Rename-ADObject -NewName $new

    #on modifier les groupe de distributio et de securité de la promo
    Get-ADObject -Filter "name -eq '$old'" | Rename-ADObject -NewName $new
    Get-ADObject -Filter "name -eq '$($old + "_distrib")'" | Rename-ADObject -NewName $($new + "_distrib")

    #on modifie le promotion de tous users de cette promtion
    Get-ADUser -Filter "department -eq '$($old.Replace("_",''))'" -Properties department `
    | Set-ADUser -Department $($new.Replace("_",''))
    
}


# function for update users from csv
function updateUsersFromCsv {
    param (
        [Parameter(Mandatory = $true)]
        [string] $pathToCsv
    )
    
    try {
        $users = loadData -pathToCsv $pathToCsv -delimiter ";"

        foreach ($user in $users) {

            $samaccountname = $user.login
            $promotion = $user.annee + "_" + $user.formation
            $display = $user.prenom +" "+ $user.nom

            if ($(isExistsUser -name $samaccountname)) {
                $adUser = Get-ADUser -Identity $samaccountname -Properties *

                if ($adUser.Surname -ne $user.nom -or $adUser.GivenName -ne $user.prenom ) {
                    <# Action to perform if the condition is true #>
                    $emailAndCommonName = formEmailAndCommonName -name $user.nom -surname $user.prenom
                    Set-ADUser  -Identity $samaccountname `
                                -Surname $user.nom `
                                -GivenName $user.prenom `
                                -DisplayName $display `
                                -EmailAddress $emailAndCommonName[0] `
                                -UserPrincipalName $emailAndCommonName[0] `
                    
                    $adUser | Rename-ADObject -NewName $emailAndCommonName[1] `
                }
        
                $department = $promotion.Replace("_",'')
        
                $year = $adUser.Department.Substring(0,4)
                $formation = $adUser.Department.Replace($year,'')
                $old_promotion = $year + '_' + $formation
        
                if ($adUser.Department -ne $department) {
                    <# Action to perform if the condition is true #>
                    updateUserPromotion -old $old_promotion -new $promotion -login $samaccountname
                }
        
            }
            else {
                Write-Warning "Login doesn't exists"
            }
        }
    }
    catch [System.SystemException] {
        "An error occurred that could not be resolved."
        Write-Host "$_.ScriptStackTrace"
    }
}


# function for update promotion from csv
function updatePromotionFromCsv {
    param (
            [Parameter(Mandatory = $true)]
            [string] $pathToCsv
    )

    try {
        $promotions = loadData -pathToCsv $pathToCsv -delimiter ";"

        foreach ($promotion in $promotions) {
            <# $promotion is the current item #>
            if ($(isExistsOU $promotion.ancien_nom)) {
                updatePromotion -old $promotion.ancien_nom -new $promotion.nouveau_nom
            }
            else {
                Write-Warning "Promotion OU to Update doesn't exists"
            }
        }
    }
    catch [System.SystemException] {
        "An error occurred that could not be resolved."
        Write-Host "$_"
    }
}


# function for assign manager
function electManager {
    param (
        [Parameter(Mandatory = $true)]
        [string] $pathToCsv
    )

    try {

        $managers = loadData -pathToCsv $pathToCsv -delimiter ";"

        foreach ($manager in $managers) {
            # all in check
            if ($(isExistsUser -name $manager.login)) {
                $department = $manager.annee + $manager.formation
                $promotion = $manager.annee + "_" + $manager.formation
                Get-ADOrganizationalUnit -Filter "name -eq '$promotion'" |
                        Set-ADOrganizationalUnit -ManagedBy $manager.login
        
                Get-ADUser -Filter "department -eq '$department'" -Properties manager | 
                        Set-ADUser -Manager $manager.login
            }
            else {
                Write-Warning "Login doesn't exists"
            }    
        }
    }
    catch [System.SystemException] {
        "An error occurred that could not be resolved."
        Write-Host "$_.ScriptStackTrace"
    }

}



# function for deactivate users
function deactivateUsers {
    param (
        [Parameter(Mandatory = $true)]
        [string] $pathToCsv
    )

    try {
        $users = loadData -pathToCsv $pathToCsv -delimiter ";"

        createOU -promotion_acronym "DELETED" -year $(Get-Date).ToString("yyyy")
    
        $path = "OU=" + $(Get-Date).ToString("yyyy") + "_DELETED," + $(convertDomainNameToPath -domain $domainName)
    
        foreach ($user in $users) {
            # all in check
            if ($(isExistsUser -name $user.login)) {
                $adUser = Get-ADUser -Identity $user.login -Properties *
                Move-ADObject -Identity $adUser.DistinguishedName -TargetPath $path
                $adUser.MemberOf | Remove-ADGroupMember -Members $user.login -Confirm:$false
                Disable-ADAccount -Identity $user.login
            }
            else {
                Write-Warning "Login doesn't exists"
            }
        }
    }
    catch [System.SystemException] {
        "An error occurred that could not be resolved."
        Write-Host $_.ScriptStackTrace
    }
}


#createUsersFromCsv -pathToCsv ./dataset/create_users.csv

#updateUserPromotion -old "2022_PSSI" -new "2022_CPS" -login "U20231452182132"

#updateUsersFromCsv -pathToCsv ./dataset/update_users.csv

#updatePromotionFromCsv -pathToCsv ./dataset/update_ou.csv

#formEmailAndCommonName -name "Sylla" -surname "Assi"

#updatePromotion -old "2022_CPS" -new "2022_CPN"

#deactivateUsers -pathToCsv ./dataset/deactivate.csv

#electManager -pathToCsv ./dataset/delegate.csv

function actionIsValid {
    param (
        [Parameter(Mandatory = $true)]
        [string] $name
    )

    $validActions = @("createUsers","updateUsers","deactivateUser","updateOu","electManager")

    if ($validActions.Contains($name)){
        return $true
    }

    return $false
}


function main {
    if ($action) {
        if ($(actionIsValid -name $action)) {
            if ($(Test-Path -Path $file -PathType Leaf)) {
                switch ($action) {
                    "createUsers" {  
                        "Creation of users"
                        createUsersFromCsv -pathToCsv $file
                    }
                    "updateUsers" {
                        "Update users"
                        updateUsersFromCsv -pathToCsv $file
                    }
                    "deactivateUser" {
                        "Deactivate User"
                        deactivateUsers -pathToCsv $file
                    }
                    "updateOu" {
                        "Update promotion"
                        updatePromotionFromCsv -pathToCsv $file
                    }
                    "electManager"{
                        "Election Manager"
                        electManager -pathToCsv $file
                    }
                }
            }
            else {
                Write-Warning "File not exists"
            }
        }
        else {
            Write-Warning "Nothing to do"
            displayHelp
        }
    }
    else {
        Write-Warning "Nothing to do"
        displayHelp
    }
}

main