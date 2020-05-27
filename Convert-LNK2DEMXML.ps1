<#
	.DESCRIPTION
		This script converts shortcuts (LNK files)  from a selected folder (recursively) to
		VMware Dynamic Environment Manager XML configuration files that are saved in a
		selected folder.
	
	.EXAMPLE
		Convert-LNK2DEMXML.ps1
		
	.NOTES
		Version	: 0.4
		Author	: Ivan de Mes
		Website	: https://www.ivandemes.com
#>

Function Get-Folder($Description, $initialDirectory)

	<#
		.DESCRIPTION
			This function shows a folder browser dialog and outputs the selection to the console or a variable.
		
		.EXAMPLE
			Get-Folder "Your description goes here" "C:\Windows\System32"
	#>
	
	{
		[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
		
		$foldername = New-Object System.Windows.Forms.FolderBrowserDialog
		$foldername.Description = $Description
		$foldername.rootfolder = "MyComputer"
		$foldername.SelectedPath = $initialDirectory
		
		if($foldername.ShowDialog() -eq "OK")
			{
				$folder += $foldername.SelectedPath
			}
		
		return $folder
	}

Clear-Host

# Display important notice and ask for reaction
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
$varNoticeReaction=[System.Windows.Forms.MessageBox]::Show("The content of the VMware Dynamic Environment Manager (DEM) XML files created by this script is the same as the content of the DEM XML files created by DEM. However, DEM XML files created outside DEM are not officially supported by VMware.`n`nPress OK to acknowledge this notice, or press CANCEL to exit this script.","Important Notice",[System.Windows.Forms.MessageBoxButtons]::OKCancel,64)
Switch ($varNoticeReaction)
	{
		"OK"		{
					Write-Host "`nYou have acknowledged the important notice.`n"
				}
				
		"Cancel"	{
					Write-Host "`nYou have canceled the script.`n"
					Start-Sleep -Seconds 2
					Exit
				}
	}


$inPath = Get-Folder "Select the SOURCE folder that contains the shortcuts." "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs"
If ($inPath -eq $null)
	{
		Write-Host "`nYou did not specify a SOURCE folder. This script will now exit.`n"
		Start-Sleep -Seconds 2
		Exit
	}
	
$outPath = Get-Folder "Select the TARGET folder where the XML files will be saved." "C:\"
If ($outPath -eq $null)
	{
		Write-Host "`nYou did not specify a TARGET folder. This script will now exit.`n"
		Start-Sleep -Seconds 2
		Exit
	}

$startPath = Get-ChildItem $inPath -Recurse -Include *.lnk

ForEach ($Item in $startPath)
	{
		$Shell = New-Object -ComObject WScript.Shell
		$Properties = @{
					ShortcutName = $Item.Name
					Target = $Shell.CreateShortcut($Item).targetpath
					Arguments = $Shell.CreateShortcut($Item).arguments
					WorkingDirectory = $Shell.CreateShortcut($Item).workingdirectory
					WindowStyle = $Shell.CreateShortcut($Item).windowstyle
					IconLocation = $Shell.CreateShortcut($Item).iconlocation
					Description = $Shell.CreateShortcut($Item).description
				}
		
		$object = New-Object PSObject -Property $Properties
		
		If ($Item.DirectoryName -ne $inPath)
			{
				$newFolders = ($Item.DirectoryName).Replace("$inPath\","")
			}
		Else
			{
				$newFolders = ($Item.DirectoryName).Replace("$inPath","")
			}

		If ($newFolders -ne "")
			{
				New-Item -ItemType Directory -Path "$outPath\$newFolders" -Force | Out-Null
			}
		
		$xmlName = ($object.ShortcutName).Replace(".lnk",".xml")
		$shortcutName = ($object.ShortcutName).Replace(".lnk","")
		$targetPath = $object.Target
		$targetPathArguments = $object.arguments
		$workingDirectory = $object.workingdirectory
		$windowStyle = $object.windowstyle
		$iconLocationWithIndex = $object.IconLocation
		$iconIndexSplit = $iconLocationWithIndex.Split(",")
		$iconLocation = $iconIndexSplit[0]
		$iconIndex = $iconIndexSplit[1]
		$comment = $object.description
		
		Write-Host "Creating VMware Dynamic Environment Manager shortcut XML for: $shortcutName"
		
		$XmlWriter = [System.XML.XmlWriter]::Create("$outPath\$newFolders\$xmlName")
		$xmlWriter.WriteStartDocument()
		$xmlWriter.WriteStartElement("userEnvironmentSettings")
		$xmlWriter.WriteStartElement("setting")
		$xmlWriter.WriteAttributeString("type","shortcut")
		$xmlWriter.WriteAttributeString("lnk","$shortcutName")
		$xmlWriter.WriteAttributeString("path","$targetPath")
		
		If ($targetPathArguments -ne "")
			{
				$xmlWriter.WriteAttributeString("args","$targetPathArguments")
			}
		
		$xmlWriter.WriteAttributeString("showCmd","$windowStyle")
		
		If ($iconLocationWithIndex.Substring(0,1) -eq ",")
			{
				$xmlWriter.WriteAttributeString("iconPath","$targetPath")
			}
		Else
			{
				$xmlWriter.WriteAttributeString("iconPath","$iconLocation")
			}
		
		$xmlWriter.WriteAttributeString("iconIndex","$iconIndex")
		$xmlWriter.WriteAttributeString("programsMenu","")
		
		If ($workingDirectory -ne "")
			{
				$xmlWriter.WriteAttributeString("startIn","$workingDirectory")
			}
		
		$xmlWriter.WriteAttributeString("async","1")
		$xmlWriter.WriteAttributeString("comment","$comment")
		$xmlWriter.WriteEndElement()
		$xmlWriter.WriteEndElement()
		$xmlWriter.WriteEndDocument()
		$xmlWriter.Flush()
		$xmlWriter.Close()
		
		$newFolders = ""
	}

[Runtime.InteropServices.Marshal]::ReleaseComObject($Shell) | Out-Null

Write-Host ""
