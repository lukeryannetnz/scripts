<#

.SYNOPSIS
Renames all .msg files in the provided folder according to its time received, sender name and subject.

.DESCRIPTION
Renames .msg files according to its time received, sender name and subject. NOTE this is destructive 
(the old file is deleted, unless the nondestructive parameter is supplied as $true)

.EXAMPLE
./RenameAllOutlookMsgs.ps1 -path "C:\emails"
./RenameAllOutlookMsgs.ps1 -path "C:\emails" -nondestructive $true

.NOTES

.LINK

#>

param
(
      [Parameter(Mandatory = $True,valueFromPipeline=$true)][String] $path,
      [Parameter(Mandatory = $False,valueFromPipeline=$true)][Bool] $nondestructive
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
$outlook = new-object -comobject outlook.application 
 
$msgFiles = get-childitem $path | where {$_.extension -eq ".msg"}

foreach($file in $msgFiles)
{
	$path = $file.FullName
	$item = $outlook.CreateItemFromTemplate($path)

	$subject = $item.Subject.ToString();
	$subject = Make-StringFilenameSafe -inputstring $subject

	$sender = $item.SenderName.ToString();
	$sender = Make-StringFilenameSafe -inputstring $sender

	$filename = $item.ReceivedTime.tostring("dd-MM-yyyy-hh-mm-ss-fff") + "_" + $sender + "_" + $subject

	$newpath = Split-Path -parent $path
	$newpath = "$newpath\$filename.msg";
	$newpath

	$item.SaveAs("$newpath")

	if(-not $nondestructive)
	{
		Remove-Item $path
	}
}
