<#
	.SYNOPSIS
		An example function to display how help should be written.
	
	.EXAMPLE
		Get-Help -Name Test-Help
	
	This shows the help for the example function.
#>

Function Get-Folder($Description, $initialDirectory)

	<#
		.SYNOPSIS
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

$inPath = Get-Folder "Select the SOURCE folder that contains the shortcuts." "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs"
$outPath = Get-Folder "Select the TARGET folder where the XML files will be saved." "C:\"

$StartMenu = Get-ChildItem $inPath -Recurse -Include *.lnk

ForEach ($Item in $StartMenu)
	{
		$Shell = New-Object -ComObject WScript.Shell
		$Properties = @{
					ShortcutName = $Item.Name
					Target = $Shell.CreateShortcut($Item).targetpath
					Arguments = $Shell.CreateShortcut($Item).arguments
					IconLocation = $Shell.CreateShortcut($Item).iconlocation
				}
		
		$object = New-Object PSObject -Property $Properties
		
		$xmlName = ($object.ShortcutName).Replace(".lnk",".xml")
		$shortcutName = ($object.ShortcutName).Replace(".lnk","")
		$targetPath = $object.Target
		$targetPathArguments = $object.arguments
		$iconLocationWithIndex = $object.IconLocation
		$iconIndexSplit = $iconLocationWithIndex.Split(",")
		$iconLocation = $iconIndexSplit[0]
		$iconIndex = $iconIndexSplit[1]
		
		Write-Host "Creating Dynamic Environment Manager shortcut XML for: $shortcutName"
		
		$XmlWriter = [System.XML.XmlWriter]::Create("$outPath\$xmlName")
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
		
		$xmlWriter.WriteAttributeString("showCmd","1")
		
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
		$xmlWriter.WriteAttributeString("async","1")
		$xmlWriter.WriteEndElement()
		$xmlWriter.WriteEndElement()
		$xmlWriter.WriteEndDocument()
		$xmlWriter.Flush()
		$xmlWriter.Close()
	}

[Runtime.InteropServices.Marshal]::ReleaseComObject($Shell) | Out-Null

Write-Host ""
