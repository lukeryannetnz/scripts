<#

.SYNOPSIS
Renames .msg according to its time received, sender name and subject.

.DESCRIPTION
Renames .msg according to its time received, sender name and subject. NOTE this is destructive (the old
file is deleted, unless the nondestructive parameter is supplied as $true)

.EXAMPLE
./RenameOutlookMsg.ps1 -filepath "C:\emails\1.msg"
./RenameOutlookMsg.ps1 -filepath "C:\emails\1.msg" -nondestructive $true

.NOTES

.LINK

#>

param
(
      [Parameter(Mandatory = $True,valueFromPipeline=$true)][String] $filepath,
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
 
if ($filepath -notlike "*.msg") {
    Write-Verbose "Skipping $_ (not an .msg file)..."
    return
}

$item = $outlook.CreateItemFromTemplate($filepath)

$subject = $item.Subject.ToString();
$subject = Make-StringFilenameSafe -inputstring $subject

$filename = $item.ReceivedTime.tostring("dd-MM-yyyy-hh-mm-ss-fff") + "_" + $item.SenderName + "_" + $subject

$newpath = Split-Path -parent $filepath
$newpath = "$newpath\$filename.msg";
$newpath

$item.SaveAs("$newpath")

if(-not $nondestructive)
{
	Remove-Item $filepath
}