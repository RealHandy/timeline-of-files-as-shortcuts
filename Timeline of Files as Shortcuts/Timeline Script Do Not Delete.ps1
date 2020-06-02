# You can add/remove file extensions to be searched for here.
# Make sure you use * wildcards to indicate that you want ALL of
# that type of file to be included in the timeline.
# Comma-separate them, and make sure there's no comma after the last one.
$fileTypesToInclude = @(

"*.rvt",
"*.dwg",

# Add file types ABOVE this line, in quotes, using a *,
# and ending with a comma. Every time. Don't mess that up
# or the script won't run.

# If you want to stop including PDFs in the timeline,
# you can delete the next line, but make sure that whatever
# file type is listed last does NOT have a comma after it,
# or the script won't run.
"*.pdf"
)


# DO NOT CHANGE ANYTHING BELOW THIS LINE

#DO NOT CHANGE THIS LINE. DOING SO AND RUNNING THE SCRIPT
#COULD DELETE A LOT OF THINGS IF YOU PUT THE WRONG DIRECTORY HERE
$ShortcutsFolderName = "Timeline Results"
$ShortcutsPath = -join( ".\", $ShortcutsFolderName)


$ProjectPath = Resolve-Path("..")
Write-Host "Project path is " $ProjectPath

$ScriptDir = Resolve-Path(".")

If (!(Test-Path -path $ShortcutsPath)) {
	New-Item $ShortcutsPath -ItemType Directory
}
else {
	Write-Host "$($ShortcutsPath) Already Exists, deleting old contents"
	Remove-Item "$($ShortcutsPath)/*"
}

Write-Host "Updating" $ShortcutsFolderName

#Loop through all files in the directory tree for each file type

foreach ($fileType in $fileTypesToInclude ) {
  Write-Host "Adding " $fileType " to " $ShortcutsFolderName

  #$files = Get-ChildItem -Path $ProjectPath -File -Include *.png,*.pdf -Recurse
  # Use \\?\ and -LiteralPath to allow Get-ChildItem to return paths > 260,
  # but it makes -Include not work, so have to use -Filter,
  # which means looping separately for each file extension.
  $litPath = -join("\\?\", $ProjectPath.path)
  $files = Get-ChildItem -LiteralPath $litPath -File -Filter $fileType -Recurse

  $fileCount = 0
  foreach ($file in $files) {
    $fileCount = $fileCount + 1
    if ( $fileCount -gt 1 ) {
      $cursorPos = $host.UI.RawUI.CursorPosition
      [Console]::SetCursorPosition(0,$cursorPos.Y-1)
      Write-Host "                                                                                               "
      $cursorPos = $host.UI.RawUI.CursorPosition
      [Console]::SetCursorPosition(0,$cursorPos.Y-1)
    }
    Write-Host "Adding " $fileCount ": " $file.name


  	#Write-Host "File is " $file

  	$shortcutFullPath = (Join-Path $ShortcutsPath $file.name)
  	#Write-Host "shortcutFullPath is " $shortcutFullPath
  	$shortenedPath = $file.DirectoryName.Replace($litPath+"\","")
  	If ( $shortenedPath -eq $file.DirectoryName ) {
  		# This file is in the root of the project, so it has no '\' at the end
  		# of $ProjectPath, so say so.
  		$shortenedPath = -join( "the project directory (", 
  		                        $litPath.Substring($litPath.lastIndexOf('\') + 1), ")" )
  	}
  	$shortcutFullPath = -join( $shortcutFullPath, "   in ",
  	                           $shortenedPath.Replace("\"," - " ), ".lnk" )
  	#Write-Host "shortcutFullPath is " $shortcutFullPath

  	If ( ($ScriptDir.path.length + $shortcutFullPath.length) -ge 248 ) {
  		$firstDir = $shortenedPath.Substring(0, $shortenedPath.IndexOf('\'))
  		$lastDir = $shortenedPath.Substring($shortenedPath.lastIndexOf('\'))
  		$ellipsisedPath = -join( $firstDir, "\...", $lastDir)
  		#Write-Host "firstDir is " $firstDir
  		#Write-Host "lastDir is " $lastDir
  		#Write-Host "ellipsisedPath is " $ellipsisedPath
  		$shortcutFullPath = (Join-Path $ShortcutsPath $file.name)
  		$shortcutFullPath = -join( $shortcutFullPath, "   in ", 
  		                           $ellipsisedPath.Replace("\"," - " ), ".lnk" )

  		# If it's STILL too long a path, then just use the file name. Oh well.
  		#Write-Host "ellipsised shortcutFullPath length is " ($ScriptDir.path.length + $shortcutFullPath.length)
  		If ( ($ScriptDir.path.length + $shortcutFullPath.length) -ge 248 ) {
  			$shortcutFullPath = (Join-Path $ShortcutsPath $file.name)
  			$shortcutFullPath = -join($shortcutFullPath,".lnk")
  		}
  	}

  	If (!(Test-Path -path $shortcutFullPath)) {
      # Need WScript.Shell to create shortcuts.
      $WshShell = New-Object -comObject WScript.Shell

    	$Shortcut = $WshShell.CreateShortcut( $shortcutFullPath )

    	# file path starts with \\?\ at this point, so get rid of that.
    	$Shortcut.TargetPath = (Join-Path $file.DirectoryName.Substring(4) $file.name)
    	$Shortcut.Save()

    	#Write-Host "targetpath is " $Shortcut.TargetPath
  		#Write-Host "shortcutFullPath is " $shortcutFullPath
    	#$shortcutLitPath = -join($ProjectPath.path, "\Timeline of Files by Date\", $shortcutFullPath)
    	#Write-Host "$(Get-Member -InputObject $Shortcut)"
    	#Write-Host "shortcutLitPath is " $shortcutLitPath
    	#$savedShortcut = $(Get-ChildItem -LiteralPath $Shortcut.FullName() -File)
    	#Write-Host "savedShortcut is " $savedShortcut

    	$(Get-ChildItem -LiteralPath $shortcutFullPath -ErrorVariable err).lastwritetime = $file.lastwritetime
      if($err.Count -gt 0){
        Write-Host "The error occurred for file " $shortcutFullPath
        Write-Host " " # To keep the line above from being blanked by the cursor move
      }
  	}

  }

}