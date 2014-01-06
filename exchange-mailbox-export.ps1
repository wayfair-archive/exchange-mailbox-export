param(
	[Parameter(Mandatory = $true, Position = 0)]
	[ValidateScript({Test-Path -Path $_})]
	$UsersToArchiveTextFile,
	[Parameter(Mandatory = $true, Position = 1)]
	$NetworkMailboxExportDest = "\\Server\ServerPath",
	[Parameter(Mandatory = $true, Position = 2)]
	[string[]] $EmailTo = @("John Doe <jdoe@domain.com>"),
	[Parameter(Mandatory = $true, Position = 3)]
	[string[]] $EmailFrom = @("John Doe Admin <jdoeadmin@domain.com>"),
	[Parameter(Mandatory = $true, Position = 4)]
	$EmailSubject = "Mailbox Export Failures and Errors",
	[Parameter(Mandatory = $true, Position = 5)]
	$SMTPServer = "mail.com"
)

#Region Variable Declaration
$scriptPath = Split-Path $MyInvocation.MyCommand.Path

#Dot-Source Exchange Remoting capabilities and then try to connect
. "$env:ExchangeInstallPath\bin\RemoteExchange.ps1"; Connect-ExchangeServer -Auto -AllowClobber -ClearCache

#script Variables
$script:logPath = "$scriptPath\Logs"
$script:curDate = Get-Date -UFormat "%m_%d_%Y"
$script:logFile = "$logPath\specific_User_Log_$curDate.txt"
$script:newLine = "`r`n"
$script:emailStr = "`r`n"

#Check if the log directory exists, if not, make it
if(!(Test-Path -Path $logPath)){
	New-Item $logPath -ItemType Directory
}
#Check if the log file exists, if not, make it
if(!(Test-Path -Path $logFile)){
	New-Item $logFile -ItemType File
}
#EndRegion

#Region Functions
function createPST {
	param (
		#Doing some extra parameter validation. Ensuring the param isn't null / empty
		[ValidateNotNullOrEmpty()]
		$UserName
	)
	
	#Create a directory just for this user's exported mailbox
	if(!(Test-Path -Path "$NetworkMailboxExportDest\$UserName")){
		New-Item "$NetworkMailboxExportDest\$UserName" -ItemType Directory
	}
	
	#If the PST does not already exist, try to queue up an export
	if(!(Test-Path -Path "$NetworkMailboxExportDest\$UserName.pst")){
		
		#Set the current error variable equal to $null
		$curError = $null
		New-MailboxExportRequest -Mailbox $UserName -FilePath $NetworkMailboxExportDest\$UserName.pst -Priority:High -BadItemLimit:3 -Confirm:$false -Name:$UserName –Verbose -ErrorVariable curError
		
		#If an error was triggered trying to export, log it and add it to the email string
		if ($curError){
			$gd = Get-Date
			Write-Output "$gd - *WARNING* - Execution of the New-MailboxExportRequest cmdlet has failed for $UserName with error `r`n$curError" | Out-File $logFile -Append
			$script:emailStr = "$script:emailStr $newLine $gd - *WARNING* - Execution of the New-MailboxExportRequest cmdlet has failed for $UserName with error <br>$curError<br>"
		}
		
		#Otherwise, just log that it was successful
		else{
			$gd = Get-Date
			Write-Output "$gd - Successfully queued a mailbox export request for $UserName." | Out-File $logFile -Append
		}
	}
	
	#Else, log that it a file with that name already exists.
	else{
		$gd = Get-Date
		Write-Output "$gd - *NOTICE* - PST for $UserName at $NetworkMailboxExportDest\$UserName\$UserName.pst already exists.  Not queueing again." | Out-File $logFile -Append
		$script:emailStr = "$script:emailStr <br>$gd - *NOTICE* - PST for $UserName at $NetworkMailboxExportDest\$UserName\$UserName.pst already exists.  Not queueing again."
	}
}

Function Get-CompletedExportRequests { 
	
	#Get any completed Exports and log it
	$exportRequestsCompleted = Get-MailboxExportRequest | Where-Object { $_.Status -eq "Completed" } 
	$gd = Get-Date
	Write-Output "`r`n`r`n$gd - The following mailboxes are Completed with the export process.`
						$newLine`
						$exportRequestsCompleted`
						" | Out-File $logFile -Append
}

Function Get-InProgressExportRequests { 

	#Get any in progress Exports and log it
	$exportRequestsInProgress = Get-MailboxExportRequest | Where-Object { $_.Status -eq "InProgress" }
	$gd = Get-Date
	Write-Output "$gd - The following mailboxes are currently In Progress with the export process.`
						$newLine`
						$exportRequestsInProgress`
						" | Out-File $logFile -Append

}

#EndRegion


#Region Main

#Try to get the content of the file specified.  We can assume valid as it was set in the Param to test-path
$userList = Get-Content $UsersToArchiveTextFile

#Loop through all the users in the text file
foreach($user in $userList){
	
	#Make sure you didn't grab an empty string
	if(($user -ne "") -and ($user -ne $null)){
			createPST $user
	}
}

#Get the completed export requests and add them to the log file and email
Get-CompletedExportRequests

#Get the status of export requests that do not have a status of complete
Get-InProgressExportRequests

#Send an email with the logfile attached
$logFile | 
		Send-MailMessage `
			-SmtpServer $SMTPServer `
			-From $EmailFrom `
			-To $EmailTo `
			-Subject $EmailSubject `
			-BodyAsHtml `
			-Body $script:emailStr
			
#EndRegion