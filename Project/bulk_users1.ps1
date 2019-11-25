# Import active directory module for running AD cmdlets
Import-Module activedirectory
  
#Store the data from ADUsers.csv in the $ADUsers variable
$ADUsers = Import-csv C:\bulk_users1.csv

#Loop through each row containing user details in the CSV file 
foreach ($User in $ADUsers)
{
	#Read user data from each field in each row and assign the data to a variable as below
		
	$Username 	= $User.username
	$Password 	= $User.password
	$Firstname 	= $User.firstname
	$Lastname 	= $User.lastname
	$OU 		= $User.ou #This field refers to the OU the user account is to be created in
    $email      = $User.email
    $streetaddress = $User.streetaddress
    $city       = $User.city
    #$zipcode    = $User.zipcode
    $state      = $User.state
    #$country    = $User.country
    $telephone  = $User.telephone
    $jobtitle   = $User.jobtitle
    $company    = $User.company
    $department = $User.department
    $Password = $User.Password


	#Check to see if the user already exists in AD
	if (Get-ADUser -F {SamAccountName -eq $Username})
	{
		 #If user does exist, give a warning
		 Write-Warning "A user account with username $Username already exist in Active Directory."
	}
	else
	{
		#User does not exist then proceed to create the new user account
		
        #Account will be created in the OU provided by the $OU variable read from the CSV file
		New-ADUser `
            -SamAccountName $Username `
            -UserPrincipalName "$Username@team2.local" `
            -Name "$Firstname $Lastname" `
            -GivenName $Firstname `
            -Surname $Lastname `
            -Enabled $True `
            -DisplayName "$Lastname, $Firstname" `
            -Path $OU `
            -City $city `
            -Company $company `
            -State $state `
            -StreetAddress $streetaddress `
            -OfficePhone $telephone `
            -EmailAddress $email `
            -Title $jobtitle `
            -Department $department `
            -AccountPassword (convertto-securestring $Password -AsPlainText -Force) -ChangePasswordAtLogon $True
    }
    #Create a new folder for users
    
    $sharepath = "C:\Parent-Directory\$Username\"

    $DirectoryToCreate = "C:\Parent-Directory\$Username\"

    if (-not (Test-Path -LiteralPath $DirectoryToCreate)) {

        try {
            New-Item -Path $DirectoryToCreate -ItemType Directory -ErrorAction Stop | Out-Null #-Force
        }
        catch {
            Write-Error -Message "Cannot create Directory '$DirectoryToCreate'. Error was $_" -ErrorAction Stop
        }
        "Created directory '$DirectoryToCreate'."

    }
    else {
        "Directory already existed for $Username!"
    }


    #Create a network share for each account

    
    $SMBToCreate = "C:\Parent-Directory\$Username\"

    if (-not (Test-Path -LiteralPath $SMBToCreate)) {

        try {
            New-SmbShare -Name $Username -Path $sharepath -FullAccess "team2\Administrator", "team2\domain admins" -Readaccess "team2\$Username"
        }
        catch {
            Write-Error -Message "Cannot create SMB Share '$SMBToCreate'. Error was: $_" -ErrorAction Stop
        }
        "Created Directory '$SMBToCreate'."
    }
    else {
        "SMB Share already exists!"
    }

}

#Output when user was created to the terminal screen
   Write-Host
Get-ADUser -Filter * -Properties created | Select-Object Name , created | Sort-Object created -Descending

#Create a new Mailbox for a postoffice/domain

if ((Get-PSSnapin -Name MailEnable.Provision.Command -ErrorAction SilentlyContinue) -eq $null )
{
    Add-PsSnapin MailEnable.Provision.Command
}
New-MailEnableMailbox -Mailbox "$username" -domain "team2.local" -Password "$password" -Right "USER"

$From = "team2.local"
$To = "$Username@activedirectorypro.com"
$Subject = "Sent from Client"
$Body = "Welcome!"
$SMTPServer = "dc.team2.local"
$SMTPPort = "587"
$Mailpass = ConvertTo-SecureString "password" -AsPlainText -Force
$Mailcreds = New-Object System.Management.Automation.PSCredential($From, $Mailpass)
Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -Credential 

# Get the list of MailBoxes in the post office
Get-MailEnableMailbox -Postoffice "team2.local"

