<!-- Back to top link -->
<a name="readme-top"></a>

<!-- NAME -->
# Remove Calendar Events
**Remove-CalendarEvents.ps1** 

<!-- ABSTRACT -->
## ABSTRACT 
This script deletes calendar items in Microsoft Outlook based on the subject.

<!-- ABOUT THE PROJECT -->
## DESCRIPTION
You may need to use this script if you have Outlook performance issues. This kind of problem occurs on macOS when the calendar has more than 5000 items/events.
The script uses the Outlook COM object to interact with the user's Outlook application. It loops through each item in the user's calendar and deletes the ones whose subject matches the specified string.

* What is it ?
    - Interactive script : 
    - Cmdlet script :
    
* Who is it for ?
    - Regular User
    
 * Why to use it ? 
    - Fix Outlook perf
    
 * When to use it ?
    - When you know the events you have to delete and do not have access to Exchange server or a Windows device.
    
 <p align="right">(<a href="#readme-top">back to top</a>)</p>
 
<!-- Getting Started -->
## QUICKSTART

### Prerequisites
Get information about
* Windows version ; Version must be 10 or alter
    * _Cmdlet_
    ```powershell
    Get-ComputerInfo
    ```
    * _Environment Class_
    ```powershell
    [Environment]::OSVersion
    ```
* Powershell version ; Version must be 5.1 or later
    * _Cmdlet_
    ```powershell
    Get-Host
    ```
    * _Automatic Variable_
    ```powershll
    $PSVersionTable
    ```
### Installation

1. Open a PowerShell prompt with eleveted permissions
2. Set the PowerShell execution policies for Windows computers to Unrestricted
3. Download the archive of the project
4. Extract the content of the archive


 <p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- ROADMAP -->
## ROADMAP

| Windows | Linux | MacOS|
| :----: | :---: | :--: |
| In progress | To be decided | To be decided |

- [ ] Windows
    - [x] Script
    - [ ] Cmdlet
   

<p align="right">(<a href="#readme-top">back to top</a>)</p>


<!-- LICENSE -->
## LICENSE

Distributed under the  Unlicense license. See `LICENSE` for more information.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- ACKNOWLEDGMENTS -->
## LINKS
* [Microsoft PowerShell Documentation](https://learn.microsoft.com/en-us/powershell/)
 
<p align="right">(<a href="#readme-top">back to top</a>)</p>
 

<!-- CONTACT -->
## CONTACT

:e-mail: 

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- CONTRIBUTOR -->
## CONTRIBUTOR
[![Contributor Covenant](https://img.shields.io/badge/Contributor%20Covenant-2.1-4baaaa.svg)](code_of_conduct.md)
