# Cisco-Device-Inventory

This is an extremely extensive Powershell script that I've been perfecting for years.  I'm still finalizing it
but at the current stage it's quite functional.  You m ight still need to make tweaks for your own environment.

        File Name : Cisco-Device-Inventory.ps1
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr.com)
                   : 
       Description : Tracks cisco switches, routers, wifi AP's, and voice gateway inventory in a multi-tab
                   : Excel spreadsheet.  Eventual option to use SQLlite has not been added yet.  This script
                   : has been under development for around 2 years.  I'm listing the first public release as 
                   : version 1 even though it should be version 20.
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
