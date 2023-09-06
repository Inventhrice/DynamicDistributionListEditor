#This contains a list of conditions that you can use in the script. Add more as you will, just make sure they are valid in this list (found in the link below)
# https://learn.microsoft.com/en-us/powershell/exchange/recipientfilter-properties?view=exchange-ps
$listOfConditions = "Department ", "DisplayName ", "Title ", "Manager ", "Office ", "RecipientType "

#This function validates the nuber inputs asked of from the user.
function getInput{
    param(
        $lowerLim,
        $upperLim
    )
    $inputValue = 0
    do {
        $inputValid = [int]::TryParse((Read-Host 'Enter a choice (number)'), [ref]$inputValue)
        if (-not $inputValid -and ($inputValue -ge $lowerLim) -and ($inputValue -le $upperLim)) {
            Write-Host "Input is not valid."
        }
    } while (-not $inputValid)
    return $inputValue
}

# A function to print a menu from an array
function printMenu{
    param(
        #the array of options to iterate through to print the menu
        $array
    )
    $i = 0
    foreach($el in $array){
        $i++
        Write-Host "$i) $el"
    }
}

# A function to get a filter from the user
function getFilter{
    param([int]$count)
    $filter = ""
    
    if($count -gt 0){
        Write-Host "1) -and`n2) -or"
        $inputOption = (getInput(1, 2))
        $filter = $(If($inputOption -eq 1){"-and"}Else{"-or"})
    }

    # if you can think of a better way to do this, be my guest.
    Write-Host "Type '1' if using -not, else type '2'"
    $notCondition = getInput(1,2)

    #asking the user for conditions
    printMenu($listOfConditions)
    $condtionInput = (getInput(1,($listOfConditions.Length+1)))-1 #adding a -1 so that it gives us the actual index in the array

    Write-Host "1) -eq`n2) -like"
    $comparison = getInput(1,2)

    #No input checking here because it can be whatever you want it to be. Messing it up is the user's problem, not ours lol.
    $filterTime = Read-Host -Prompt "Parameter for the filter"

    #this basically condenses everything to (-not(Title -eq 'NAME')) <- that is an example
    $filter += '(' + $(If($notCondition -eq 1){"-not("}Else{""}) + $listOfConditions[$condtionInput] + $(If($comparison -eq 1){"-eq '"}Else{"-like '"}) + $filterTime + $(If($notCondition -eq 1){"')"}Else{"'"}) + ")"
    
    return $filter
}

# Function that edits a DL, it returns a string that conforms to the -RecipientFilter argument in Set-DynamicDistributionGroup
function EditDL{
    $returnVal = ""
    $exitLoop = $False
    
    while(-not $exitLoop){
        #don't get rid of this. We add these options (and remove them later) so that they always appear as the last options
        $listOfFilters += "Add New"
        $listOfFilters += "Confirm changes to DL"
        
        # prints out the menu. Referring to the above comment, this is why we do not use the Add New option available to us in this function
        printMenu($listOfFilters)

        $i = ($listOfFilters.Count)
        $inputOption = (getInput(1, $i))-1
        
        $listOfFilters.Remove($listOfFilters[$i-1])
        $listOfFilters.Remove($listOfFilters[$i-2])
        
        if($inputOption -eq ($i-1)){
            $exitLoop = $True
        }

        elseif($inputOption -eq ($i-2)){
            $listOfFilters += getFilter($listOfFilters.Count)
        }
        else{
            #gets the filter from the list of options
            $filter = $listOfFilters[$inputOption]
            $filter = $filter.replace("-and","").replace("-or","")

            #gets all the recipient data from the filter
            if($filter.contains("-not") -eq $True){
                $details = Get-Recipient -RecipientPreviewFilter $($filter.substring(0,$filter.length-1).replace("-not(","")) | Select-Object DisplayName, WindowsLiveId, Title    
            }
            else{
                $details = Get-Recipient -RecipientPreviewFilter $filter | Select-Object DisplayName, WindowsLiveId, Title
            }

            #displays the filter, who's included in that filter, and the menu options.
            Write-Host $filter":`n" $details.DisplayName "`n1) Edit`n2) Remove`n3) Exit"

            $filterChoiceOption = getInput(1,3)
            
            if($filterChoiceOption -eq 1){
                #Edits the filter (by just asking to make a new one)
                $output = getFilter($listOfFilters.Count)
                #this HAS to be coded in this order as powershell makes $listOfFilters add a new entry if you just do $listOfFilters[$inputOption] = getFilter. It's very inconvinent and I hate it.
                $listOfFilters[$inputOption] = $output
            }
            elseif ($filterChoiceOption -eq 2) {
                #Removes the filter from the array. Easy!
                $listOfFilters.Remove($listOfFilters[$inputOption])
            }
            
        }
    }

    foreach($filter in $listOfFilters){
        $returnVal += $filter
    }

    #TODO: when hit complete, compile everything together and output the string
    return $returnVal
}

function getAndFormatDLRecipientFilter{
    [OutputType([System.Collections.ArrayList])]
    #code for getting all the recipients
    #This code takes all the recipient filters, replaces all the '-and'/'-or' with a comma so that it splits into a nice array
    [System.Collections.ArrayList] $group = (((Get-DynamicDistributionGroup -Identity $selectedGroup).RecipientFilter).replace(" -and ",',-and').replace(" -or ",',-or')) -split ","
    $i = 0
    for(;$i -lt $group.Count;$i++){
        $notFound = $True

        #This block of code takes all the filters and makes sure it is in the nice pretty format of ([condition] [-eq/-like] [value]). It also excludes all the extra stuff powershell tosses in.
        foreach ($condition in $listOfConditions){
            if($group[$i].contains($condition)){
                $group[$i] = ($group[$i]).replace('(','').replace(')','').replace("$condition", "($condition") + ')'
                
                if($group[$i].contains('-not')){
                    $group[$i] = ($group[$i].replace('-not','(-not')) + ')'
                }
        
                $group[$i] = $group[$i].replace(' )', ')')
                $notFound = $False
                break
            }
            
        }

        if($notFound -eq $True){    
            $group.Remove($group[$i])
            $i--
        }
    }
    return $group
}

#Install-Module -Name ExchangeOnlineManagement
#Import-Module ExchangeOnlineManagement 
#Connect-ExchangeOnline

#$pathToList = ".\ListofDL.txt"
#$pathToRecipient = ".\RecipientList.txt"
$selectedGroup = ""
$recipientFilter = ""


#Gets all the distribution groups and exports it into a txt file
#Get-DynamicDistributionGroup | Select-Object -prop DisplayName, Identity | Export-Csv $pathToList -NoTypeInformation

$allDL = Get-DynamicDistributionGroup | Select-Object -prop DisplayName, Identity

#Front end code for getting user input for DL choice
$i = 1
$inputValue = 0
$allDisplayNames = $allDL.DisplayName
$allDisplayNames += "Add New"
#prints out all the DLS
printMenu($allDisplayNames)
$i = $allDisplayNames.Count
$inputValue = (getInput(1, $i))-1

# logic for choosing Add New
if($inputValue -lt ($i-1)){
    $selectedGroup = $allDL[$inputValue].Identity
}


#logic for add new/existing DL
if($selectedGroup){
    [System.Collections.ArrayList] $group = getAndFormatDLRecipientFilter
    
    $recipientFilter = EditDL($group)
    Set-DynamicDistributionGroup -Identity $selectedGroup -RecipientFilter $recipientFilter
}
else{
    $selectedGroup = Read-Host "Enter name of DL"
    [System.Collections.ArrayList] $group = New-Object System.Collections.ArrayList
    $recipientFilter = EditDL($group)
    New-DynamicDistributionGroup -Name $selectedGroup -RecipientFilter $recipientFilter
}

