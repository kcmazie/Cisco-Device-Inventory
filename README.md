# Cisco-Device-Inventory



        File Name : Cisco-Device-Inventory.ps1
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr.com)
                   : 

                   : 
             Notes : Normal operation is with no command line options.  If pre-stored credentials 
                   : are desired use this: https://github.com/kcmazie/CredentialsWithKey
                   : Spreadsheet can be generated via a flat text list of IP addresses or pulled from a master
                   : copy.  The first column contains the list of IP addresses OR the appropriate tab from a 
                   : master inventory.  Each device type gets a dedicated worksheet labeled with the type.  Column A 
                   : is the IP addresses.  The spreadsheet is color coded for readability according to end of 
                   : support dates.  Various device milestone dates are in a lookup table to allow for color coding.
                   : The colors are set according to a predefined priority number.  High priority (red) defaults to 
                   : one year from "end of support" date, etc..  In most cases pre-existing data is left but options
                   : exist near the end of the script to force how data is written.  A debug mode is available to display
                   : extra data to the screen.  Options like user and password are externalized in an XML file 
                   : so that nothing sensitive is conbtained within the script.  Initial creation of the spreadsheet 
                   : can be done with a formatted flat text file.  If the text file exists it will always be used
                   : 
                   :
      Requirements : Plink.exe must be available in your path or the full path must be included in the 
                   : commandline(s) below.  2 versions are used in case of version issues.  These are located
                   : in the same folder and renamed according to version (see around line 256 below). Excel must be 
                   : available on the local PC.  SSH Keys must already be stored on the local PC through the
                   : use of PuTTY or connection will fail.  An option exists to add it below.
                   : 
   Option Switches : See descriptions above.
                   :
          Warnings : Excel is set to be visible (can be changed) so don't mess with it while the script is 
                   : running or it can crash.  Don't click in spreadsheet while running or the script will crash.
                   :   
             Legal : Public Domain. Modify and redistribute freely. No rights reserved.
                   : SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF 
                   : ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED.
                   : That being said, feel free to ask if you have questions...
                   :
           Credits : Code snippets and/or ideas came from too many sources to list...

<!---
<meta name="google-site-verification" content="SiI2B_QvkFxrKW8YNvNf7w7gTIhzZsP9-yemxArYWwI" />
-->
[![Minimum Supported PowerShell Version][powershell-minimum]][powershell-github]&nbsp;&nbsp;
[![GPLv3 license](https://img.shields.io/badge/License-GPLv3-blue.svg)](http://perso.crans.org/besson/LICENSE.html)&nbsp;&nbsp;
[![made-with-VSCode](https://img.shields.io/badge/Made%20with-VSCode-1f425f.svg)](https://code.visualstudio.com/)&nbsp;&nbsp;
![GitHub watchers](https://img.shields.io/github/watchers/kcmazie/Cisco-Device-Inventory?style=plastic)

[powershell-minimum]: https://img.shields.io/badge/PowerShell-5.1+-blue.svg 
[powershell-github]:  https://github.com/PowerShell/PowerShell

# $${\color{Cyan}Powershell \space "Cisco-Device-Inventory" \space Script}$$

#### $${\color{orange}Original \space Author \space : \space \color{white}Kenneth \space C. \space Mazie \space \color{lightblue}(kcmjr \space AT \space kcmjr.com)}$$

## $${\color{grey}Description:}$$ 
This is an extremely extensive Powershell script that I've been perfecting for years.  I'm still finalizing it but at the current stage it's quite functional.  You might still need to make tweaks for your own environment.

Tracks cisco switches, routers, wifi AP's, and voice gateway inventory in a multi-tab Excel spreadsheet.  Eventual option to use SQLlite has not been added yet.  This script has been under development for around 2 years.  I'm listing the first public release as version 1 even though it should be version 20.

My initial desire was to have a tool that when a WAN circuit outage alert was received all primary addresses at the site could be rapidly checked.  I wanted it simple, quick, and easy, as well as versitile enough that I could alter the lists with minimal effort, plus in order to share it could contain no proprietary data.

By itself the script is completely generic, there are no default settings.  Configuration is managed via a companion XML config file which must be located in the same folder and named the same as as the script (but obviously with a ".xml" extension).  If the config file is NOT detected a sample XML file will be created on first run that must be edited before use.  Data from the sample is shown in the screen shots below.  

The config file has fields to contain technical contact info as well as the location address for each site.  It also has the ability to list circuit IDs in the target list for reference.

The script dynamically adjusts it's size depending on the number of sites and/or targets in the config file up to a limit after which it will exceed the screen size.  

The script uses a simple ICMP echo (ping) to identify if a target is alive. Select one of the listed sites and the associated info from the config file will populate the GUI. Each time you select a different site the info will dynamically update. Once the site is loaded simply click the Execute button to run the ping check. Results are noted
along the right side of the screen.  After each run you can select a different site and execute again 
continuously until you click the cancel button. 

## $${\color{grey}Notes:}$$ 
* Normal operation is with no command line options.
* Powershell 5.1 is the minimal version required.

## $${\color{grey}Arguments:}$$ 
Command line options for testing: 
* "-console $true" will enable local console echo for troubleshooting
* "-debug $true" will only email to the first recipient on the list

## $${\color{grey}Configuration:}$$ 
The script takes virtually all configuration from the companion XML file.  As previously noted the file must exist and if not found  one will be created with basic settings.

The XML file broken down into 4 sections each of which falls under the section heading of "<Settings>".

#### $${\color{darkcyan}"General"  Section:}$$
   This section sets the visible script title and email settings for future use.
   ```xml
        <General>
            <TitleText>Enterprise Network Engineering "Quick" Site Check</TitleText>            
            <SmtpServer>Not Used</SmtpServer>
            <SmtpPort>25</SmtpPort>
            <EmailRecipient>Not Used</EmailRecipient>
            <EmailSender>Not Used</EmailSender>
        </General>
   ```
#### $${\color{darkcyan}"TargetTemplate"  Section:}$$
   This section formats names and order of the labels in the results on the right side of the GUI.  There is also a short section explenation.
   ```xml
        <TargetTemplate>LoopbackIP;GatewayIP;Circuit1;PrivateIP1;PublicIP1;Circuit2;PrivateIP2;PublicIP2</TargetTemplate> 
            <!--  NOTE: Target Template MUST match the order and number of targets in each site section. Below is an EXAMPLE. 
            <Target1>Loopback IP</Target1>
            <Target2>Gateway IP</Target2>
            <Target3>Circuit 1</Target3> 
            <Target4>Private IP 1</Target4>
            <Target5>Public IP 1</Target5>
            <Target6>Circuit 2</Target6>
            <Target7>Private IP 2</Target7>
            <Target8>Public IP 2</Target8>
            -->
   ```
 #### $${\color{darkcyan}"Sites"  Section:}$$
   This section is a list of sites.  Each site listed should have a dedicated section defining it's specifics.  The number after the comma identifies the section of the XML for that site.
   ```xml
        <Sites>
        	<!-- Site ID number must match site tag below, i.e "sitename,1" and "<Site1>" -->
            <site>New York,1</site>
            <site>San Francisco,2</site>
            <site>Los Angeles,3</site> 
            <site>Denver,4</site>
            <site>San Jose,5</site>
            <site>Seattle,6</site>
            <site>Atlanta,7</site>
            <site>Chicago,8</site>
            <site>London,9</site>
            <site>Paris,10</site> 
        </Sites>
   ```

#### $${\color{darkcyan}"SiteX"  Section:}$$
   Note that the number in the site header corresponds to the aforementioned site ID.  This section has site specific info and there should be one of these for each site listed in the "Sites" section.  The "TargetX" lines must match the order and number of items in the "TargetTemplate" section.  If the template lists a circuit ID, then that target should contain circuit info noit an IP.
   ```xml
        <Site1>
            <Designation>Site-01</Designation>
            <Name>New York Office</Name>
            <Address>123 Sesame Street, New York NY 12345</Address>
            <Contact>Joe Tech</Contact>
            <Email>techdude@123.com</Email>
            <CellPhone>123-456-7890</CellPhone>
            <DeskPhone>234-567-8901</DeskPhone>
            <Target1>10.0.1.1</Target1>
            <Target2>10.1.1.1</Target2>
            <Target3>Comcast: 45.YADA.12345..XYZL</Target3> 
            <Target4>10.2.2.7</Target4>
            <Target5>10.2.2.6</Target5>
            <Target6>ATT: 50ASDF666333PT</Target6>
            <Target7>10.3.3.7</Target7>
            <Target8>10.3.3.6</Target8>
        </Site1>
   ```
   
### $${\color{grey}Screenshots:}$$ 
   This is the initial GUI.
   
![Initial GUI](https://github.com/kcmazie/Site-Check/blob/main/Screenshot1.jpg "Initial GUI")

   This is the GUI after selecting a site.
   
![GUI With Site Selected](https://github.com/kcmazie/Site-Check/blob/main/Screenshot2.jpg "GUI With Site Selected")

   This is the GUI during a run (with some live IPs substituted).
   
![GUI While Test is Running](https://github.com/kcmazie/Site-Check/blob/main/Screenshot3.jpg "GUI While Test is Running")

  
### $${\color{grey}Warnings:}$$ 
* None 

### $${\color{grey}Enhancements:}$$ 
Some possible future enhancements are:
* Ability to email the results
* Still need to add some error checking when IP addresses are missing or misspelled.

### $${\color{grey}Legal:}$$ 
Public Domain. Modify and redistribute freely. No rights reserved. 
SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED. 

That being said, please let me know if you find bugs, have improved the script, or would like to help. 

### $${\color{grey}Credits:}$$  
Code snippets and/or ideas came from many sources including but not limited to the following: 
* n/a 
  
### $${\color{grey}Version \\& Change History:}$$ 
* Last Update by  : Kenneth C. Mazie 
  * Initial release : v1.00 - 09-06-24
  * Change History  : v2.00 - 00-00-00 
 
