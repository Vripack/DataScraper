### Functions
function SearchInPDFFiles
{
param($path,$searchStrings,$SearchType,$searchunit,$Logfile) 
.{  
    $files = Get-Childitem $path -Include *.pdf -Recurse -Exclude *memo*| Where-Object { !($_.psiscontainer) }

    # Loop through all files in the $path directory
    Foreach ($file In $files)
    {
        #write-Output "$(Get-Date): Start search at file $($file.FullName)" >> $logfile
        try {
            $pdf = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $($file.FullName)
        }
        catch {
            #Error opeing file
            write-Output "$(Get-Date): ERROR: unable to open file $($file.FullName)" >> $logfile
            return $null
        }
        
        $doctext = $null
        #copy content into clipboard and paste it into varialble. This strips content from unwanted stuff
        try {
            for ($page = 1; $page -le $pdf.NumberOfPages; $page++){
                $doctext += [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf,$page)
            }	    
        }
        catch {
            write-Output "$(Get-Date): ERROR: unable to get content of file $($file.FullName)" >> $logfile
            return $null
        }
        try {
            $pdf.Close()
            $pdf = $null
        }
        catch {
            write-Output "$(Get-Date): ERROR: unable to close file $($file.FullName)" >> $logfile
            return $null
        }

        if ($doctext) {
            $lines = $doctext -split "`n"

            for ($lineNumber = 0; $lineNumber -lt $lines.Length; $lineNumber++) {
                $line = $lines[$lineNumber]
                
                foreach ($searchString in $searchStrings) {
                    $positions = @()
                    $index = 0
            
                    while ($index -ne -1) {
                        $index = $line.IndexOf($searchString, $index, [System.StringComparison]::InvariantCultureIgnoreCase)
                        if ($index -ne -1) {
                            $positions += $index
                            $index += $searchString.Length  # Move past the current occurrence
                        }
                    }
            
                    if ($positions.Count -eq 0) {
                        #Write-Output "The string '$searchString' was not found in line $(($lineNumber + 1))."
                    } else {
                        foreach ($position in $positions) {
                            switch ($SearchType) {
                                "decimal" {
                                    #Search for decimal values
                                    #Extract the next 40 characters after the found position
                                    $startPos = $position + $searchString.Length
                                    if ($startPos + 40 -le $line.Length) {
                                        $nextChars = $line.Substring($startPos, 40)
                                    } else {
                                        $nextChars = $line.Substring($startPos)
                                        $nextChars = $nextChars.Replace("`t", " ") #remove tabs
                                        $nextChars = $nextChars.Replace(",", ".") #make comma a dot
                                    }
                                    switch ($SearchUnit) {
                                        "m" {
                                            #Search for length
                                            $FoundValue = GetMeterFromString $nextChars 

                                            if ($FoundValue -eq 0) {
                                                #Try finding feet
                                                $nextChars = $nextChars.replace("’","ft") 
                                                $nextChars = $nextChars.replace("$([char]0x2018)","ft")
                                                $nextChars = $nextChars.replace("$([char]0x2019)","ft") 
                                                $nextChars = $nextChars.replace("$([char]0x00B4)","ft") 
                                                $nextChars = $nextChars.replace("'","ft")
                            
                                                $FoundValue = GetFootFromString $nextChars 
                                                #write-Output "$(Get-Date): Value $FoundValue tested as feet" >> $logfile
                                                if ($FoundValue -gt 0) {
                                                    #Feet found. Make it meters
                                                    $FoundValue = $FoundValue * 0.3048
                                                }
                                                else {
                                                    #Still nothing. Can be unspecified.assume meters
                                                    $FoundValue = GetDecimalFromString $nextChars 
                                                    #write-Output "$(Get-Date): Value $FoundValue tested as feet" >> $logfile
                                                    if ($FoundValue -gt 0) {
                                                        #Value found. 
                                                    }
                                                    else {
                                                        #Still nothing.
                                                    }
                                                }
                                            }
                            
                                            if ($FoundValue) {
                                                if ($FoundValue -gt 0) {
                                                    #A value has been found. Return it
                                                    #write-Output "$(Get-Date): Value $FoundValue found at line '$line' of file $($file.FullName)" >> $logfile
                                                    #echo $doctext > "c:\temp\dump.txt"
                                                    return $FoundValue
                                                }
                                            }
                                            Break
                                        }
                                        "w" {
                                            #Search for weights
                                            $FoundValue = GetWeightFromString $nextChars 
                                            if ($FoundValue -gt 0) {
                                                #Value found. 
                                                return $FoundValue
                                            }
                                            else {
                                                #Still nothing.
                                            }
                                        }
                                    }
                                    break
                                }
                                "string" {
                                    #Search for string values

                                    return $null
                                    break
                                }
                            } 
                        }
                    }
                }
            }
        }
        else {
            #write-Output "$(Get-Date): File is empty" >> $logfile
        }
    
    }

    return $FoundValue
}   | Out-Null

return $FoundValue
}