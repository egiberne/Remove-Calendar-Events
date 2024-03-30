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

* What is it ?  A script not a Cmdlet script.
    
* Who is it for ? Regular User.
    
 * Why to use it ? Fix Outlook perf.
    
 * When to use it ? When you know the events you have to delete and do not have access to Exchange server or a Windows device.
    
 <p align="right">(<a href="#readme-top">back to top</a>)</p>
 
<!-- Getting Started -->
## QUICKSTART

### Prerequisites
- Microsoft PowerShell
- Microsoft Outlook

### Installation

Regarding Microsoft PowerShell, go to the pages for :
- Windows 
[https://learn.microsoft.com/en-gb/powershell/scripting/install/installing-powershell-on-windows](https://learn.microsoft.com/en-gb/powershell/scripting/install/installing-powershell-on-windows)
- MacOS
[https://learn.microsoft.com/en-gb/powershell/scripting/install/installing-powershell-on-windows](https://learn.microsoft.com/en-gb/powershell/scripting/install/installing-powershell-on-macos)
- Linux 
[https://learn.microsoft.com/en-gb/powershell/scripting/install/installing-powershell-on-linux](https://learn.microsoft.com/en-gb/powershell/scripting/install/installing-powershell-on-linux)

Regarding Microsoft Outlook go to the pages :
- Individual subscirption
  https://account.microsoft.com/services/microsoft365
- Enterprise subscription
  https://portal.office.com/account

### Usage 
1. Download Remove-CalendarEvents.ps1 and save on your computer.
2. Start PowerShell and go to the folder where you saved the file.
3. Set the PowerShell execution policies for Windows computers to Unrestricted
   ```powershell
     Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
   ```
5. If you want to cancel the meeting calendar named,   ``` My Meeting   ``` , run the following command
    ```powershell
    .\Remove-CalendarEvents.ps1 -SubjectToDelete "My Meeting"
    ```


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

<!-- FEEDBACK -->
## FEEDBACK

If you have any feedback, please post on the [Issues section](https://github.com/egiberne/Remove-Calendar-Events/issues).

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- CONTRIBUTOR -->
## CONTRIBUTION
Feel free to contribute to its improve.


[![Contributor Covenant](https://img.shields.io/badge/Contributor%20Covenant-2.1-4baaaa.svg)](code_of_conduct.md)
