### Functions
function GetLocationOnFileServer
{
param($ProjectNumber,$Logfile) 
.{
    #Find project number on file server
    [System.String]$myString = $ProjectNumber
    $startPos = $myString.IndexOf("-")

    if ($startpos -gt 0) {
        $part = $ProjectNumber.split("-")

        try {
            [int]$MyTemp = $part[0]
            }
        catch {
            #Not a number. Abort
            $projectlocation = $null
            return $projectlocation
        }
    }
    else {
        try {
            [int]$MyTemp = $ProjectNumber
            }
        catch {
            #Not a number. Abort
            $projectlocation = $null
            return $projectlocation
        }
    }

    If ($MyTemp -lt 500)                         {$projectlocation = "\\FS01\projects\0000-0499\$ProjectNumber"}
    If ($MyTemp -gt 499 -And $MyTemp -lt 1000)   {$projectlocation = "\\FS01\projects\0500-0999\$ProjectNumber"}
    If ($MyTemp -gt 999 -And $MyTemp -lt 1500)   {$projectlocation = "\\FS01\projects\1000-1499\$ProjectNumber"}
    If ($MyTemp -gt 1499 -And $MyTemp -lt 2000)  {$projectlocation = "\\FS01\projects\1500-1999\$ProjectNumber"}
    If ($MyTemp -gt 1999 -And $MyTemp -lt 2500)  {$projectlocation = "\\FS01\projects\2000-2499\$ProjectNumber"}
    If ($MyTemp -gt 2499 -And $MyTemp -lt 3000)  {$projectlocation = "\\FS01\projects\2500-2999\$ProjectNumber"}
    If ($MyTemp -gt 2999 -And $MyTemp -lt 3500)  {$projectlocation = "\\FS01\projects\3000-3499\$ProjectNumber"}
    If ($MyTemp -gt 3499 -And $MyTemp -lt 4000)  {$projectlocation = "\\FS01\projects\3500-3999\$ProjectNumber"}
    If ($MyTemp -gt 3999 -And $MyTemp -lt 4500)  {$projectlocation = "\\FS01\projects\4000-4499\$ProjectNumber"}
    If ($MyTemp -gt 4499 -And $MyTemp -lt 5000)  {$projectlocation = "\\FS01\projects\4500-4999\$ProjectNumber"}
    If ($MyTemp -gt 4999 -And $MyTemp -lt 5500)  {$projectlocation = "\\FS01\projects\5000-5499\$ProjectNumber"}
    If ($MyTemp -gt 5499 -And $MyTemp -lt 6000)  {$projectlocation = "\\FS01\projects\5500-5999\$ProjectNumber"}
    If ($MyTemp -gt 5999 -And $MyTemp -lt 6500)  {$projectlocation = "\\FS01\projects\6000-6499\$ProjectNumber"}
    If ($MyTemp -gt 6499 -And $MyTemp -lt 7000)  {$projectlocation = "\\FS01\projects\6500-6999\$ProjectNumber"}
    If ($MyTemp -gt 6999 -And $MyTemp -lt 7500)  {$projectlocation = "\\FS01\projects\7000-7499\$ProjectNumber"}
    If ($MyTemp -gt 7499 -And $MyTemp -lt 8000)  {$projectlocation = "\\FS01\projects\7500-7999\$ProjectNumber"}
    If ($MyTemp -gt 7999 -And $MyTemp -lt 8500)  {$projectlocation = "\\FS01\projects\8000-8499\$ProjectNumber"}
    If ($MyTemp -gt 8499 -And $MyTemp -lt 9000)  {$projectlocation = "\\FS01\projects\8500-8999\$ProjectNumber"}
    If ($MyTemp -gt 8999 -And $MyTemp -lt 9500)  {$projectlocation = "\\FS01\projects\9000-9499\$ProjectNumber"}
    If ($MyTemp -gt 9499 -And $MyTemp -lt 10000) {$projectlocation = "\\FS01\projects\9500-9999\$ProjectNumber"}

    if (Test-Path $projectlocation) {
    }
    else {
        #Write-host "Folder Doesn't Exists. Search archive" -f Red
        $projectlocation = $projectlocation.Replace('FS01\projects','FS02\ProjectArchive')

        if (Test-Path $projectlocation) {
            #Write-host "Folder Exists at archive" -f Green
        }
        else {
            #Write-host "Folder Doesn't Exists at archive" -f Red
            $projectlocation = $null
        }
    }

    return
 }   | Out-Null

return $projectlocation
}