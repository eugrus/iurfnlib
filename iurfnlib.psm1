# iurfnlib
# PowerShell functions to automate procedures in legal firms
# -- Evgeny Fishgalov, 2025

function insertintoword { # Bausteine
	param ([string]$pathToInsertableTemplate)
	$pathToInsertableTemplate = Resolve-Path -Path $pathToInsertableTemplate
	$msword = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Word.Application")
	$activeDocument = $msword.ActiveDocument
	$insertableTemplate = $msword.Documents.Open($pathToInsertableTemplate)
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

function fillspaceholderinword { # Bausteine mit Platzhaltern
	param (
		[string]$findText,
		[string]$pathToInsertableTemplate
	)
	$pathToInsertableTemplate = Resolve-Path -Path $pathToInsertableTemplate
	$msword = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Word.Application")
	$activeDocument = $msword.ActiveDocument
	$insertableTemplate = $msword.Documents.Open($pathToInsertableTemplate)
	
	# Find the placeholder in the active document
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
	
	# Execute the find operation
	if ($findRange.Find.Execute()) {
		# Replace the found placeholder with the template content
		$templateRange = $insertableTemplate.Content
		$findRange.FormattedText = $templateRange.FormattedText
	}
	
	$insertableTemplate.Close($false)
}

function Rubrumauslese () {
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

function killheadlessword () {
	Get-Process -Name WINWORD | Where-Object {
			$_.MainWindowHandle -eq 0 -and $_.SessionId -eq $([System.Diagnostics.Process]::GetCurrentProcess().SessionId)
		} | Stop-Process -Force
}
