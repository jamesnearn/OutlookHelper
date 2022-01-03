function GetFolders {
    Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
    $Outlook = New-Object -ComObject Outlook.Application
    $namespace = $Outlook.GetNameSpace("MAPI")

    # $myAccount = $namespace.Folders | Where Name -eq "username@domain.com"
    $myAccount = $namespace.Folders.GetFirst()

    return $myAccount.Folders
}

function FindFolderAndCreateIfNecessary {
    param (
        $Folders,
        [string] $FolderName = ""
    )

    $targetFolder = $Folders | Where Name -eq $FolderName

    if ($targetFolder -eq $null) {
        $targetFolder = $Folders.Add($FolderName)
    }

    return $targetFolder
}



$inbox = GetFolders | Where Name -eq "Inbox"

Write-Host "moving Sent Items to Inbox"
$sentItems = GetFolders | Where Name -eq "Sent Items"
$sentItems.Items | % {
    $moveResults = $_.Move($inbox)
}



Write-Host "moving items"
$inboxAzureDevOps = FindFolderAndCreateIfNecessary -Folders $inbox.Folders -FolderName "azuredevops@microsoft.com"
$inboxNoReply = FindFolderAndCreateIfNecessary -Folders $inbox.Folders -FolderName "NoReply"
$inboxAccepted = FindFolderAndCreateIfNecessary -Folders $inbox.Folders -FolderName "Accepted"
$inboxExternal = FindFolderAndCreateIfNecessary -Folders $inbox.Folders -FolderName "External"
for ($i = $inbox.Items.Count - 1; $i -gt 0 ; $i--)
{
    $emailToCheck = $inbox.Items.Item($i)
    $sender = $emailToCheck.SenderEmailAddress

    if ($sender.Contains("azuredevops@microsoft.com")) {
	$moveResults = $emailToCheck.Move($inboxAzureDevOps)
    }
    elseif ($sender.Contains("noreply") -or $sender.Contains("no-reply")) {
	$moveResults = $emailToCheck.Move($inboxNoReply)
    }
    elseif ($emailToCheck.Subject.StartsWith("Accepted: ")) {
        $moveResults = $emailToCheck.Move($inboxAccepted)
    }
}
