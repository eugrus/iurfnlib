# iurfnlib
# PowerShell functions to automate processes at law firms
# -- Evgeny Fishgalov, 2025

function KillHeadLessWord () {
	Get-Process -Name WINWORD | Where-Object {
			$_.MainWindowHandle -eq 0 -and $_.SessionId -eq $([System.Diagnostics.Process]::GetCurrentProcess().SessionId)
		} | Stop-Process -Force
}

function insertintoword { # Bausteine
	param ([string]$pathToInsertableTemplate)
	KillHeadLessWord
	$pathToInsertableTemplate = Resolve-Path -Path $pathToInsertableTemplate
	$msword = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Word.Application")
	$activeDocument = $msword.ActiveDocument
	$insertableTemplate = $msword.Documents.Open( # Positional arguments (COM interop) https://learn.microsoft.com/en-us/office/vba/api/word.documents.open
		$pathToInsertableTemplate,        # FileName
		[ref]$false,                      # ConfirmConversions = false (no convert dialog for RTF)
		[ref]$true,                       # ReadOnly
		[System.Type]::Missing,           # AddToRecentFiles
		[System.Type]::Missing,           # PasswordDocument
		[System.Type]::Missing,           # PasswordTemplate
		[System.Type]::Missing,           # Revert
		[System.Type]::Missing,           # WritePasswordDocument
		[System.Type]::Missing,           # WritePasswordTemplate
		[System.Type]::Missing,           # Format
		[System.Type]::Missing,           # Encoding
		[ref]$false                       # Visible = false (hides the template)
	)
	$templateRange = $insertableTemplate.Content
	$activeRange = $activeDocument.Content
	$activeRange.Collapse([ref]0)
	$activeRange.FormattedText = $templateRange.FormattedText
	$insertableTemplate.Close($false)
	$selection = $msword.Selection
	$selection.EndKey(6) # = wdStory https://learn.microsoft.com/en-us/office/vba/api/word.wdunits
}

function replaceinword { # Platzhalter
	param (
		[string]$findText,
		[string]$replacewithText
	)
	KillHeadLessWord
	$msword = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Word.Application")
	# https://learn.microsoft.com/en-us/office/vba/api/word.find.execute
	$MatchCase = $false
	$MatchWholeWord = $true
	$MatchWildcards = $false
	$MatchSoundsLike = $false
	$MatchAllWordForms = $false
	$Forward = $true
	$Wrap = 1 # https://learn.microsoft.com/en-us/office/vba/api/word.wdfindwrap
	$Format = $false
	$Replace = 2 # = wdReplaceAll https://learn.microsoft.com/en-us/office/vba/api/word.wdreplace
	$doc = $msword.ActiveDocument
	$doc.Content.Find.Execute($findText, $MatchCase, $MatchWholeWord, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap, $Format, $replacewithText, $Replace)
}

function FillSpaceholderInWord { # Bausteine mit Platzhaltern
	param (
	[string]$findText,
	[string]$pathToInsertableTemplate
	)
	KillHeadLessWord
	$pathToInsertableTemplate = Resolve-Path -Path $pathToInsertableTemplate
	$msword = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Word.Application")
	$activeDocument = $msword.ActiveDocument
	$insertableTemplate = $msword.Documents.Open( # Positional arguments (COM interop) https://learn.microsoft.com/en-us/office/vba/api/word.documents.open
		$pathToInsertableTemplate,        # FileName
		[ref]$false,                      # ConfirmConversions = false (no convert dialog for RTF)
		[ref]$true,                       # ReadOnly
		[System.Type]::Missing,           # AddToRecentFiles
		[System.Type]::Missing,           # PasswordDocument
		[System.Type]::Missing,           # PasswordTemplate
		[System.Type]::Missing,           # Revert
		[System.Type]::Missing,           # WritePasswordDocument
		[System.Type]::Missing,           # WritePasswordTemplate
		[System.Type]::Missing,           # Format
		[System.Type]::Missing,           # Encoding
		[ref]$false                       # Visible = false (hides the template)
	)

	# Get template content once
	$templateRange = $insertableTemplate.Content

	# Find and replace ALL occurrences of the placeholder
	$findRange = $activeDocument.Content.Duplicate
	$findRange.Find.ClearFormatting()
	$findRange.Find.Text = $findText
	$findRange.Find.MatchCase = $false
	$findRange.Find.MatchWholeWord = $true
	$findRange.Find.MatchWildcards = $false
	$findRange.Find.MatchSoundsLike = $false
	$findRange.Find.MatchAllWordForms = $false
	$findRange.Find.Forward = $true
	$findRange.Find.Wrap = 1 # wdFindContinue

	# Execute the find operation in a loop to replace ALL occurrences
	while ($findRange.Find.Execute()) {
		# Replace the found placeholder with the template content
		$findRange.FormattedText = $templateRange.FormattedText
		# Reset the range for the next search
		$findRange = $activeDocument.Content.Duplicate
		$findRange.Find.ClearFormatting()
		$findRange.Find.Text = $findText
		$findRange.Find.MatchCase = $false
		$findRange.Find.MatchWholeWord = $true
		$findRange.Find.MatchWildcards = $false
		$findRange.Find.MatchSoundsLike = $false
		$findRange.Find.MatchAllWordForms = $false
		$findRange.Find.Forward = $true
		$findRange.Find.Wrap = 1
	}

	$insertableTemplate.Close($false)
}


function Rubrumauslese () {
	KillHeadLessWord
	$msword = [Runtime.Interopservices.Marshal]::GetActiveObject('Word.Application')
	$doc = $msword.ActiveDocument
	$documentText = $doc.Content.Text -replace '\r',[System.Environment]::Newline

	$Az = [regex]::match($documentText, '\d{1,5}/\d{2}').Value

	$Mandant, $Gegner = ((Select-String -InputObject $documentText -Pattern '.*\./\..*').matches.value -split './.').trim()
	
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($msword) | Out-Null

	return [PSCustomObject]@{
		Mandant = $Mandant
		Gegner = $Gegner
		Az = $Az
	}
}

function FillFromJSON { # Platzhalter für kurze Eintragungen (bis 255 Zeichen) anhand JSON ersetzen

	param(
		[Parameter(Mandatory=$true)][string]$Selection,
		[Parameter(Mandatory=$true)][string]$pathToJSON
	)
	
	KillHeadLessWord
	
	$json = Get-Content -Path $pathToJSON -Raw -Encoding UTF8 | ConvertFrom-Json

	foreach ($entry in $json) {
		$pl = $entry.PH.Trim()
		$val = $entry.$Selection
		if ($val) {
			ReplaceInWord $pl $val
		}
	}
}

function FillFromFolder { # Längere Bausteine für einige Platzhalter aus einem Ordner holen
	
	param(
		[Parameter(Mandatory=$true)][string]$Selection,
		[Parameter(Mandatory=$true)][string]$Folder
	)
	
	KillHeadLessWord
	
	if (Test-Path "$Folder\$Selection") {
		$files = Get-ChildItem -Path "$Folder\$Selection" -Include "*.docx", "*.rtf" -Recurse
		
		foreach ($file in $files) {
			$fileName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
			$placeholder = "ph" + $fileName
			
			FillSpaceholderInWord $placeholder $file.FullName
		}
	}
}
