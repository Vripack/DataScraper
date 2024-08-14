### Various Functions
function GetDecimalFromString {
    #(            Match the regular expression below and capture its match into backreference number 1
    #\d           Match a single digit 0..9
    #   +         Between one and unlimited times, as many times as possible, giving back as needed (greedy)
    #(?:          Match the regular expression below
    #   \.        Match the character “.” literally
    #   \d        Match a single digit 0..9
    #      +      Between one and unlimited times, as many times as possible, giving back as needed (greedy)
    #)?           Between zero and one times, as many times as possible, giving back as needed (greedy)
    #)

    # Function receives string containing decimal
    param([String]$myString)

    if ($myString -match '(\d+(?:\.\d+)?)') { [decimal]$matches[1] } else { [decimal]::Zero }
}

function GetMeterFromString {
    # Function receives string containing decimal
    param([String]$myString)

    if ($myString -match '(\d+(?:\.\d+)?)\s*(m|M)') { [decimal]$matches[1] } else { [decimal]::Zero }
}

function GetFootFromString {
    param([String]$myString)
    if ($myString -match '(\d+(?:\.\d+)?)\s*(ft)') { [decimal]$matches[1] } else { [decimal]::Zero }
}


function GetWeightFromString {
    param([String]$myString)
    if ($myString -match '(\d+(?:\.\d+)?)\s*(gt|mt)') { [decimal]$matches[1] } else { [decimal]::Zero }
}


function GetLengthFromString {
    # Function receives string containing decimal
    param([String]$myString)

    $nextChars = $myString
    $nextChars = $nextChars.Replace(",", ".") #make comma a dot
    $index1    = $nextChars.indexof("#",0)
    $index2    = $nextChars.indexof("'",0)

    if (($index1 -gt -1 -and $index2 -gt -1  -and $index2 -lt $index1) -or $index1 -eq -1){
        #There is a # sign, but it is after ', or there is no #
        $FoundValue = GetMeterFromString $nextChars 
        if ($FoundValue -eq 0) {
            #Try finding feet
            $nextChars = $nextChars.replace("’","ft") 
            $nextChars = $nextChars.replace("$([char]0x2018)","ft")
            $nextChars = $nextChars.replace("$([char]0x2019)","ft") 
            $nextChars = $nextChars.replace("$([char]0x00B4)","ft") 
            $nextChars = $nextChars.replace("'","ft")

            $FoundValue = GetFootFromString $nextChars 
            
            if ($FoundValue -gt 0) {
                #Feet found. Make it meters
                #echo "Length found in feet: $FoundValue"
                $FoundValue = $FoundValue * 0.3048
                #echo "$FoundValue meters"
            }
            else {
                #Still nothing. Can be unspecified.assume feet
                $FoundValue = GetDecimalFromString $nextChars 
                [string]$mytemp = $FoundValue 
                $dotindex = $mytemp.indexof(".",0) 

                if ($FoundValue -gt 0) {
                    if ($dotindex -gt -1) {
                        #a dot means meters
                        #echo "Length found in m with dot: $FoundValue"
                        #echo "$FoundValue meters"
                    }
                    elseif ($FoundValue -gt 300) {
                        #This case it is cm
                        #Value found. 
                        #echo "Length found in cm: $FoundValue"
                        $FoundValue = $FoundValue / 100
                        #echo "$FoundValue meters"
                    }
                    else {
                        #Value found. 
                        #echo "Length found in feet: $FoundValue"
                        $FoundValue = $FoundValue * 0.3048
                        #echo "$FoundValue meters"
                        $FoundValue = GetMeterFromString $nextChars 

                        if ($FoundValue -eq 0) {
                            $FoundValue = GetFootFromString $nextChars 
                        }
                    }

                }
                else {
                    #Still nothing.
                    #remove # and try again
                    $nextChars = $nextChars.Replace("#", "nr")
                    $FoundValue = GetMeterFromString $nextChars 
                    if ($FoundValue -eq 0) {
                        $FoundValue = GetFootFromString $nextChars
                    }
                }
            }
        }
        else {
            #echo "Length in meters: $FoundValue"
            }
    }
    else {
        $nextChars = $nextChars.Replace("#", "nr")

        $FoundValue = GetMeterFromString $nextChars 
        if ($FoundValue -eq 0) {
            $FoundValue = GetFootFromString $nextChars
        }
    }
    return $FoundValue
}

