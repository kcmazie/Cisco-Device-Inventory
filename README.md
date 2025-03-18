<!---
<head>
<meta name="google-site-verification" content="SiI2B_QvkFxrKW8YNvNf7w7gTIhzZsP9-yemxArYWwI" />
</head>
-->
[![Minimum Supported PowerShell Version][powershell-minimum]][powershell-github]&nbsp;&nbsp;
[![GPLv3 license](https://img.shields.io/badge/License-GPLv3-blue.svg)](http://perso.crans.org/besson/LICENSE.html)&nbsp;&nbsp;
[![made-with-VSCode](https://img.shields.io/badge/Made%20with-VSCode-1f425f.svg)](https://code.visualstudio.com/)&nbsp;&nbsp;
![GitHub watchers](https://img.shields.io/github/watchers/kcmazie/Cisco-Device-Inventory?style=plastic)

[powershell-minimum]: https://img.shields.io/badge/PowerShell-5.1+-blue.svg 
[powershell-github]:  https://github.com/PowerShell/PowerShell
<span style="background-color:black">
# $${\color{Cyan}Powershell \space "Cisco-Device-Inventory.ps1"}$$

#### $${\color{orange}Original \space Author \space : \space \color{white}Kenneth \space C. \space Mazie \space \color{lightblue}(kcmjr \space AT \space kcmjr.com)}$$

## $${\color{grey}Description:}$$ 
This is an extremely extensive Powershell script that I've been perfecting for years.  I'm still finalizing it but at the current stage it's quite functional.  You might still need to make tweaks for your own environment.

The script racks cisco switches, routers, wifi AP's, and voice gateway inventory in a multi-tab Excel spreadsheet.  This script has been under development for around 2 years.  I'm listing the first public release as version 1 even though it should be version 20.

The spreadsheet can be initially generated via a flat text list of IP addresses or pulled from a master copy.  Note that as long as the text file exists it will always be used so remove after import.

The first column contains the list of IP addresses OR the appropriate tab from a master inventory.  Each device type gets a dedicated worksheet labeled with the type.  Column A is the IP addresses.  The spreadsheet is color coded for readability according to end of support dates.  Various device milestone dates are in a lookup table to allow for color coding.  The colors are set according to a predefined priority number.  High priority (red) defaults to one year from "end of support" date, etc.  In most cases pre-existing data is left but options exist near the end of the script to force how data is written.  

A debug mode is available to display extra data to the screen for troubleshooting.  

Options like user and password are externalized in a companion XML file so that nothing sensitive is contained within the script.

## $${\color{grey}Notes:}$$ 
* Normal operation is with no command line options.
* Powershell 5.1 is the minimal version required.
* Plink.exe must be available in your path or the full path must be included in the commandline(s) below.  2 versions are used in case of version issues.  These are located in the same folder and renamed according to version (see around line 256).
* Excel must be available on the local PC.  SSH Keys must already be stored on the local PC through the use of PuTTY or connection will fail.  An option exists to add it below.

## $${\color{grey}Arguments:}$$ 
Command line options for testing: 
| Option | Description | Default Setting
| --------------------------- | ---------------------------------------------------------------------- | ----------------- |
| Console     | Set to true to enable local console result display. Defaults to false |
| Debug       | Generates extra console output for debugging.  Defaults to false |
| EnableExcel | Defaults to use Excel. | |
| SafeUpdate  | Forces a copy made with a date/Time stamp prior to editing the spreadsheet as a safety backup. | Defaults to false |
| StayCurrent | Will copy a new version of the spreadsheet from source if the date stamps don't match. | Defaults to false |

## $${\color{grey}Configuration:}$$ 
The script takes virtually all configuration from the companion XML file.  As previously noted the file must exist and if not found the script will abort.  A message will pop-up showing the basic settings should the file not be found.

The XML file broken down into multiple sections each of which falls under the section heading of "Settings".

* $${\color{darkcyan}"General"  Section:}$$ This section sets the run parameters such as username and password (or encrypted files to use), folder locations, email recipients, etc.
* $${\color{darkcyan}"Credentials"  Section:}$$ This section stores hard coded credentials or can be ignored by enabling a credential prompt.  If encrypted, pre-stored credentials are desired, use this: https://github.com/kcmazie/CredentialsWithKey
* $${\color{darkcyan}"Recipients"  Section:}$$ This section stores email addresses of potential status email recipients.
 
```xml
<?xml version="1.0" encoding="utf-8"?>
<Settings>
    <General>
        <SmtpServer>mailserver.company.org</SmtpServer>
        <SmtpPort>25</SmtpPort>
        <RecipientEmail>InformationTechnology@company.org</RecipientEmail>
        <SourcePath>C:\folder</SourcePath>
        <ExcelSourceFile>+NetworkedDevice-Master-Inventory.xlsx</ExcelSourceFile>
        <ExcelWorkingCopy>NetworkedDevice-Master-Inventory.xlsx</ExcelWorkingCopy>
        <Domain>company.org</Domain>
    </General>
    <Credentials>
        <PasswordFile>c:\AESPass.txt</PasswordFile>
        <KeyFile>c:\AESKey.txt</KeyFile>
        <WAPUser>admin</WAPUser>
        <WAPPass>wappass</WAPPass>
        <AltUser>user1</AltUser>
        <AltPass>userpass1</AltPass>
    </Credentials>    
    <Recipients>
        <Recipient>me@comapny.org</Recipient>
        <Recipient>you@comapny.org</Recipient>
        <Recipient>them@comapny.org</Recipient>
    </Recipients>
</Settings>
```
   
### $${\color{grey}Screenshots:}$$ 
* Coming soon.
   
<!-- ![Initial GUI](https://github.com/kcmazie/Site-Check/blob/main/Screenshot1.jpg "Initial GUI") -->
  
### $${\color{grey}Warnings:}$$ 
* Excel is set to be visible (can be changed) so don't mess with the spreadsheet while the script is running or the script will crash. 

### $${\color{grey}Enhancements:}$$ 
Some possible future enhancements are:
* I plan an eventual option to use SQLlite but that has not been added yet. 

### $${\color{grey}Legal:}$$ 
Public Domain. Modify and redistribute freely. No rights reserved. 
SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED.

That being said, please let me know if you find bugs, have improved the script, or would like to help. 

### $${\color{grey}Credits:}$$  
Code snippets and/or ideas came from many sources including but not limited to the following: 
* Code snippets and/or ideas came from too many sources to list...
  
### $${\color{grey}Version \\& Change History:}$$ 
* Last Update by  : Kenneth C. Mazie 
  * Initial Release :
    * v1.00 - 06-01-23 - Original release
  * Change History :
    * v1.90 - 00-00-00 - Numerous edits
    * v2.10 - 00-00-00 - Numerous edits
    * v3.00 - 10-16-24 - Added code for C9200 switch.  Added alternate SSH routine using Posh-SSH (still need to integrate).  Corrected error writing some data to Excel. Corrected numerous parsing errors.  Removed PS Gallery tags.
 </span>
