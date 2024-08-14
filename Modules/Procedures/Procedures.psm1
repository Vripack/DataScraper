$Public = Join-Path $PSScriptRoot 'Public'
$Private = Join-Path $PSScriptRoot 'Private'
$Functions = Get-ChildItem -Path $Public, $Private -Filter '*.ps1'

Foreach ($import in $Functions) {
	try {
	Write-Verbose "dot-sourcing file '$($import.fullname)"
	. $import.fullname
	}
	Catch {
		Write-Error -Message "Failed to import function $($import.name)"
	}
}