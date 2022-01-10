# SCCM-Patching-Report

Get your patching report directly in your mailbox instead of browser throught your SCCM Console.

## Requirements
Powershell version that support **DataVisualization**.  
Run on a server with **SCCM Module** (recommended to run from SCCM Server)  
Run the server the same day of the patching as script rely on **EnforcementDeadline**

## How it works ?:
1. Clone the repo `git clone https://github.com/SinJK/SCCM-Patching-Report.git`
2. edit the **Mail settings**:
```
#Define Mail settings
$sender = ""
$receiver = ""
$smtpserver = ""
######################
```
3. Run **Reporting.ps1**

For out-piechart, all credits goes to [Graham Beer](https://4sysops.com/archives/convert-powershell-csv-output-into-a-pie-chart/). 
I just tweaked it to have the right colors depending of the status of the machine
