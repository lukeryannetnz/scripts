<#

.SYNOPSIS
Saves all messages from the outlook inbox to the filesystem.

.DESCRIPTION
Reads messages out of the inbox, saves them to the filesystem in the path 
provided named with the time received, sender name and subject.

.EXAMPLE
./SaveOutlookEmailToFileSystem.ps1 -path "C:\emails"

.NOTES

.LINK

#>

param
(
      [Parameter(Mandatory = $True,valueFromPipeline=$true)][String] $path
)


function Make-StringFilenameSafe($inputstring)
{
	if($inputstring.Length -gt 25)
	{
		$inputstring = $inputstring.Substring(0,25);
	}

    foreach ($c in [System.IO.Path]::GetInvalidFileNameChars())
    {
        $inputstring = $inputstring.Replace("$c", "")
    }
    
    $inputstring;
}

Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
$olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]  
$outlook = new-object -comobject outlook.application 
$namespace = $outlook.GetNameSpace("MAPI") 
$folder = $namespace.getDefaultFolder($olFolders::olFolderInBox) 
 
foreach($item in $folder.items)
{
	$subject = $item.Subject.ToString();
	$subject = Make-StringFilenameSafe -inputstring $subject
	
	$sender = $item.SenderName.ToString();
	$sender = Make-StringFilenameSafe -inputstring $sender

	$filename = $item.ReceivedTime.tostring("dd-MM-yyyy-hh-mm-ss-fff") + "_" + $sender + "_" + $subject

	$filepath = "$path\$filename.msg";
	$filepath
 
	$item.SaveAs("$filepath")
}
