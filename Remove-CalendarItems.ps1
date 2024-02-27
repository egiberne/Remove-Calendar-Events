<#
.SYNOPSIS
This script deletes calendar items in Microsoft Outlook based on the subject.

.DESCRIPTION
The script uses the Outlook COM object to interact with the user's Outlook application. It loops through each item in the user's calendar and deletes the ones whose subject matches the specified string.

.PARAMETER SubjectToDelete
The subject of the calendar items to delete.

.EXAMPLE
.\Remove-CalendarItem.ps1 -SubjectToDelete "Your Subject Here"

.NOTES
This script is intended to be run on the machine where Outlook is installed and the user is logged in.
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$SubjectToDelete
)

# Load the Outlook COM object
$Outlook = New-Object -ComObject Outlook.Application

# Get the MAPI namespace
$Namespace = $Outlook.GetNamespace("MAPI")

# Get the Calendar folder
$Calendar = $Namespace.GetDefaultFolder(9) # 9 corresponds to olFolderCalendar

# Loop through each item in the Calendar
foreach ($Item in $Calendar.Items) {
    
    Write-Host $title vs $SubjectToDelete -ForegroundColor Blue
    $isSubjectToDelete = $Item.Subject -eq $SubjectToDelete
    Write-Host "Test >   $isSubjectToDelete" -ForegroundColor Yellow
    # Check if the item's subject matches the one we want to delete
    if ($Item.Subject -eq $SubjectToDelete) {
        Write-Host "$Item.Subject found" -ForegroundColor Red
        # If it does, delete the item
        $Item.Delete()
        Write-Host "$Item.Subject deleted" -ForegroundColor Green
    }
}

# Clean up the COM object
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
