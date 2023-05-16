#Find Management File
$ManagementFileLocation = GCI "$env:userprofile"| ? {$_.name -notlike ".*"} | gci -recurse  | ? {$_.name -like "LinkManagement.csv"}

#Move to Management File location
SL $ManagementFileLocation.DirectoryName

#Move Data to Object
$ManagementFile = (Import-Csv $ManagementFileLocation.FullName -Header 'FileName','SourceLocation','FileLocation','OldLink','NewLink')[1..30]

function changeFileLinks{
    #Get links that need to be updated
    $OldLink = Read-host "What is the old link you wish to replace?"
    $OldPattern = [regex]::Escape($OldLink)  # Escape special characters in the old link

    $NewLink = Read-host "What is the new link you wish to replace the old one with?"

    #Work with each file.
    foreach ($Item in $ManagementFile) {
        # Make Path Variable
        $ItemPath = "$($Item.SourceLocation)\$($Item.FileLocation)\$($Item.Filename)"

        $Content = Get-Content -Path $ItemPath
        $ModifiedContent = foreach ($Line in $Content) {
            if ($Line -match $OldPattern) {
                $Line -replace $OldPattern, $NewLink
                $ReplacedLine = $Line
            } else {
                $Line
            }
        }
        # Update File
        $ModifiedContent | Set-Content -Path $ItemPath

        #Update Management File
        If($ReplacedLine -ne $null){#$host.EnterNestedPrompt()
            ($ManagementFile | ? {$_.filename -match "$($Newlink)" }).NewLink = "$NewLink"
            ($ManagementFile | ? {$_.filename -match "$($Newlink)" }).OldLink = "$($ReplacedLine.TrimStart())"      
        }
        #Remove Varible
        Remove-Variable -Name $ReplacedLine 
    }
    Get-Process | Where-Object { $_.Name -like "*EXCEL*" -and $_.MainWindowTitle -like "$($ManagementFileLocation.Name)*" } | %{$_.kill()}
    $ManagementFile | export-csv -NoTypeInformation -Force $ManagementFileLocation 
    & $ManagementFileLocation.FullName
}