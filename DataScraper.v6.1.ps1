### This script will scrape project related information

### Logging Settings
$LogDate  = get-date -format yyyy-MM-dd
$Logpath  = "C:\temp\ScrapeLog"
$LogFile  = "$Logpath\ScrapeLog_$LogDate.txt"
If(!(test-path -PathType container $Logpath))
{
      New-Item -ItemType Directory -Path $Logpath
}

### Initiate functions
$modulePath = [Environment]::GetEnvironmentVariable('PSModulePath')
$modulePath += ";$PSScriptRoot\Modules"
[Environment]::SetEnvironmentVariable('PSModulePath', $modulePath)

Import-Module Procedures -Force

$Version             = "v1.0"

### Finding Variables
$SearchPathDetailold = "Project Management"
$SearchPathDetailnew = "9000 - Project Management"

# Initiate PDF
$PDFAppPath = "$PSScriptRoot\Modules\DLL\itextsharp.dll"
Add-Type -Path $PDFAppPath

write-Output "$(Get-Date): Starting Scrape procedure $Version" >> $logfile

#Get a list of all projects

foreach($i in 1000..8000){ 
    #get project location on file server
    $ProjectLocation = GetLocationOnFileServer -ProjectNumber $i -Logfile $logfile

    if ($ProjectLocation -eq $null) {
        write-Output "$(Get-Date): WARNING: project $i not found at file server" >> $logfile
        continue
    }
    else {
        write-Output "$(Get-Date): Location on server: $ProjectLocation" >> $logfile
        
        $SearchPath = $null
        $SearchPath = -join($ProjectLocation,"\",$SearchPathDetailnew)
        
        if (Test-Path -Path $SearchPath) {
            #Path is OK
        }
        else {
            $SearchPath = -join($ProjectLocation,"\",$SearchPathDetailold)
            if (Test-Path -Path $SearchPath) {
                #Path is OK
                }
            else {
                #There is no project management folder found. Search at whole project instead
                $SearchPath = $ProjectLocation 
            }
        }

        if ($SearchPath) {
            #######
            #Search for length over all

            $SearchUsingWord = $true
            $Length = SearchInPDFFiles -path $SearchPath -searchStrings @("Length over all","Length o.a.") -SearchType "decimal" -searchunit "m" -Logfile $logfile
            if ($Length -ne $null) {
                if($Length -gt 0) {
                    write-Output "$(Get-Date): Value found at PDF file: $Length" >> $logfile
                    $SearchUsingWord = $false
                }
            }
            else {
                #Search using Word docs
                if ($SearchUsingWord){
                    #Find total length
                    $Length = SearchInDOCXFiles -path $SearchPath -searchStrings @("Length over all","Length o.a.") -SearchType "decimal" -searchunit "m" -Logfile $logfile
                    if ($Length -ne $null) {
                        if($Length -gt 0) {
                            write-Output "$(Get-Date): Value found at Word file: $Length" >> $logfile
                        }
                    }
                    else {
                        #write-Output "$(Get-Date): NO Length found at Word files" >> $logfile
                    }
                }
            }
        }
    }
}#  

##########################################################################################################################################

write-Output "$(Get-Date): Finished complete scrape" >> $logfile
write-Output "" >> $logfile