Param(
    [switch]$Console = $false,         #--[ Set to true to enable local console result display. Defaults to false ]--
    [switch]$Debug = $False,           #--[ Generates extra console output for debugging.  Defaults to false ]--
    [switch]$EnableExcel = $True,      #--[ Defaults to use Excel. ]--   
    [switch]$SafeUpdate = $False,      #--[ Forces a copy made with a date/Time stamp prior to editing the spreadsheet as a safety backup. ]--  
    [Switch]$StayCurrent = $False,     #--[ Will copy a new version of the spreadsheet from source if the date stamps don't match ]--
    $DeviceType = "Switch"             #--[ Default to running against switches ]--
    )


<#==============================================================================
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
                   : 
    Last Update by : Kenneth C. Mazie                                           
   Version History : v1.00 - 06-01-23 - Original release
    Change History : v1.90 - 00-00-00 - Numerous edits
                   : v2.10 - 00-00-00 - Numerous edits
                   : v3.00 - 10-16-24 - Added code for C9200 switch.  Added alternate SSH routine using Posh-SSH (still 
                   :                    need to integrate).  Corrected error writing some data to Excel. Corrected 
                   :                    numerous parsing errors.  Removed PS Gallery tags.
                   : v3.01 - 03-18-25 - Corrected minor typo              
		   :                  
==============================================================================#>
Clear-Host
#Requires -version 5

#--[ Variables ]---------------------------------------------------------------
$DateTime = Get-Date -Format MM-dd-yyyy_HHmmss 
$Today = Get-Date -Format MM-dd-yyyy 

#==[ RUNTIME TESTING OPTION VARIATIONS ]========================================
#$DeviceType = "Router"
#$DeviceType = "WAP"
#$DeviceType = "VG"
$DeviceType = "Switch"
$Console = $true
$EnableExcel = $true
$EnableSQLite = $false     #--[ No SQLlite yet ]--
$Script:Debug = $True
$SafeUpdate = $false
$StayCurrent = $false 
$CleanSheet = $False
#==============================================================================

if (!(Get-Module -Name posh-ssh*)) {    
    Try{  
        import-module -name posh-ssh
    }Catch{
        Write-host "-- Error loading Posh-SSH module." -ForegroundColor Red
        Write-host "Error: " $_.Error.Message  -ForegroundColor Red
        Write-host "Exception: " $_.Exception.Message  -ForegroundColor Red
    }
}

If($Script:Debug){
    $Console = $true
}
#--[ Upgrade Criticality Chart ]-----------------------------------------------
    [int]$PriorityCritical = 1   #-- Critical Priority = Less than 2 years from LDOS --
    [int]$PriorityHigh = 3       #-- High Priority = Less than 4 years from LDOS --
    [int]$PriorityMedium = 5     #-- Medium Priority = Less than 6 years from LDOS --
    #[int]$PriorityLow = 8       #-- Low Priority = greater than 6 years from LDOS --
    #
    #--[ Note that anything 10/100 gets high priority regardless of age. ]--
#------------------------------------------------------------------------------

#==============================================================================
#==[ Functions ]===============================================================
Function LoadDebug ($Config){
    $XmlDebug = New-Object -TypeName psobject 
}

Function LoadConfig ($Config){
    If ($Config -ne "failed"){
        $XmlOption = New-Object -TypeName psobject 
        $XmlOption | Add-Member -Force -MemberType NoteProperty -Name "Domain" -Value $Config.Settings.General.Domain
        $XmlOption | Add-Member -Force -MemberType NoteProperty -Name "SourcePath" -Value $Config.Settings.General.SourcePath 
        $XmlOption | Add-Member -Force -MemberType NoteProperty -Name "ExcelSourceFile" -Value $Config.Settings.General.ExcelSourceFile 
        $XmlOption | Add-Member -Force -MemberType NoteProperty -Name "ExcelWorkingCopy" -Value $Config.Settings.General.ExcelWorkingCopy
        $XmlOption | Add-Member -Force -MemberType NoteProperty -Name "PasswordFile" -Value $Config.Settings.Credentials.PasswordFile
        $XmlOption | Add-Member -Force -MemberType NoteProperty -Name "KeyFile" -Value $Config.Settings.Credentials.KeyFile
        $XmlOption | Add-Member -Force -MemberType NoteProperty -Name "WAPUser" -Value $Config.Settings.Credentials.WAPUser
        $XmlOption | Add-Member -Force -MemberType NoteProperty -Name "WAPPass" -Value $Config.Settings.Credentials.WAPPass
        $XmlOption | Add-Member -Force -MemberType NoteProperty -Name "AltUser" -Value $Config.Settings.Credentials.AltUser
        $XmlOption | Add-Member -Force -MemberType NoteProperty -Name "AltPass" -Value $Config.Settings.Credentials.AltPass
    }Else{
        StatusMsg "MISSING XML CONFIG FILE.  File is required.  Script aborted..." " Red" $True
        $Message = (
'--[ External XML config file example ]-----------------------------------
--[ To be named the same as the script and located in the same folder as the script ]--

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
        <Recipient>me@company.org</Recipient>
        <Recipient>you@company.org</Recipient>
        <Recipient>them@company.org</Recipient>
    </Recipients>
</Settings> ')
Write-host $Message -ForegroundColor Yellow
    }
    Return $XmlOption
}

Function Write2Excel ($WorkSheet,$Row,$Col,$NewData,$Format,$Debug){
    $Header = $WorkSheet.Cells.Item($Row,$Col).Text  
     
    If (($NewData -eq "") -or ($Null -eq $NewData)){
        If ($Debug){write-host "-- No Data, No Write -- (Col"$Col")" -ForegroundColor Yellow} 
        Return
    }

    If ($WorkSheet.Cells.Item($Row,3).Text -like "*ISSUES (Ping OK)*"){
        $WorkSheet.UsedRange.Rows.Item($Row).Interior.ColorIndex = 15  #--[ Background set to grey if ping OK but logon fails ]--
    }

    If ($Debug){
        write-host "`n- Workbook Name   :"$Excel.ActiveWorkbook.Name  -foregroundcolor Magenta  
        write-host "- Worksheet Name  :"$Worksheet.name -foregroundcolor Magenta  
        write-host "- Format          :"$Format
        write-host "- Col Header      :"$Header                                     #--[ To validate that data is going to the right column ]--
        write-host "- New Data        :"$NewData -ForegroundColor Green
    }

    If ($Script:NewSpreadsheet){                                                    #--[ Creating a new spreadsheet, set all cells to black ]--
        $Worksheet.Cells($Row, $Col).Font.Bold = $False
        $Worksheet.Cells($Row, $Col).Font.ColorIndex = 0                            #--[ Black ]--    
        $Existing = ""   
    }Else{                                                                          #--[ Using existing spreadsheet. ]--
        $Existing = $WorkSheet.Cells.Item($Row,$Col).Text                           #--[ Read existing spreadsheet cell data for comparison ]-- 
        If ($Debug){write-host "- Existing Data   :"$Existing -ForegroundColor Cyan}

        If (($NewData -eq "") -Or ($Null -eq $NewData) -and ($Format -eq "existing")){ 
            $NewData = $Existing        
        }

        If ($NewData -eq $Existing){                                    #--[ New data matches existing data ]--
            $Matched = $True
            If ($Debug){write-host "-- Data Matches..." -ForegroundColor yellow}    
            $Worksheet.Cells($Row, $Col).Font.Bold = $False    
            $Worksheet.Cells($Row, $Col).Font.ColorIndex = 0            #--[ Black to denote a new item or no change ]-- 
        }Else{
            $Matched = $False
            $Worksheet.Cells($Row, $Col).Font.Bold = $true
            $Worksheet.Cells($Row, $Col).Font.ColorIndex = 7            #--[ Violet to denote a change ]--
        }   
        $Worksheet.Cells($Row, $Col).NumberFormat = "@"                 #--[ Set cell format to TEXT by default ]--
        Switch ($Format){
            "red" {
                $Worksheet.Cells($Row, $Col).Font.Bold = $True
                $Worksheet.Cells($Row, $Col).Font.ColorIndex = 3
            }
            "green" {
                $Worksheet.Cells($Row, $Col).Font.Bold = $True
                $Worksheet.Cells($Row, $Col).Font.ColorIndex = 10
            }
            "number" {
                If ($Existing -gt $NewData){
                    $Worksheet.Cells($Row, $Col).Font.Bold = $false
                    $Worksheet.Cells($Row, $Col).Font.ColorIndex = 10       #--[ Green ]--
                }
                If ($NewData -gt $Existing){
                    $Worksheet.Cells($Row, $Col).Font.Bold = $true
                    $Worksheet.Cells($Row, $Col).Font.ColorIndex = 3        #--[ Red ]--
                }
            }
            "date" {
                $Worksheet.Cells($Row, $Col).NumberFormat = "mm/dd/yyyy"
                $Worksheet.Cells($Row, $Col).Font.Bold = $False
                $Worksheet.Cells($Row, $Col).Font.ColorIndex = 0            #--[ Black ]--
            }
            "mac" {                                                         #--[ Flag MAC conflicts Red but leave existing data ]--
                If ($NewData -ne $Existing){   
                    $Worksheet.Cells($Row, $Col).Font.Bold = $true
                    $Worksheet.Cells($Row, $Col).Font.ColorIndex = 3            #--[ Red ]--
                    $WorkSheet.cells.Item($Row, $Col) = $Existing
                    If ($Null -ne $Worksheet.Cells($Row, $Col).Comment()){      #--[ If a previous comment exists, remove it before adding new ]-
                        $Worksheet.Cells($Row, $Col).Comment.Delete()
                    }  
                [void]$WorkSheet.cells.Item($Row, $Col).AddComment("Detected MAC = "+$NewData)
                }
            }
        }
    }

    #--[ Perform the actual data write to Excel ]--  
    If ($Format -eq "url"){  
        $WorkSheet.Hyperlinks.Add($WorkSheet.Cells.Item($Row, $Col),$NewData.Split(";")[0],"","",$NewData.Split(";")[1]) | Out-Null
    }Else{
        $erroractionpreference = "silentlycontinue"
        If ($Matched){  #--[ New data equals old data, don't over write ]--
            If ($Debug){write-host "-- No Write --" -ForegroundColor green }
        }Else{
            If ($Debug){
                write-host "-- Writing        :"$NewData -ForegroundColor Red 
                write-host "-- Row            :"$Row -ForegroundColor Red 
                write-host "-- Col            :"$Col -ForegroundColor Red 
            }
            Try{
                $WorkSheet.cells.Item($Row,$Col) = $NewData
            }Catch{
                Write-host $_.Exception.Message
            }
        }
    } 
}

Function StatusMsg ($Msg, $Color, $Debug){
    $Debug = $true
    If ($Debug){Write-Host "-- Script Status: $Msg" -ForegroundColor $Color}
}

Function GetSSH ($TargetIP,$Command,$Credential){
    Get-SSHSession | Select-Object SessionId | Remove-SSHSession | Out-Null  #--[ Remove any existing sessions ]--
    New-SSHSession -ComputerName $TargetIP -AcceptKey -Credential $Credential | Out-Null
    $Session = Get-SSHSession -Index 0 
    $Stream = $Session.Session.CreateShellStream("dumb", 0, 0, 0, 0, 1000)
    $Stream.Write("terminal Length 0 `n")
    Start-Sleep -Milliseconds 60
    $Stream.Read() | Out-Null
    $Stream.Write("$Command`n")
    sleep -millisec 100
    $ResponseRaw = $Stream.Read()
    $Response = $ResponseRaw -split "`r`n" | ForEach-Object{$_.trim()}
    while (($Response[$Response.Count -1]) -notlike "*#") {
        Start-Sleep -Milliseconds 60
        $ResponseRaw = $Stream.Read()
        $Response = $ResponseRaw -split "`r`n" | ForEach-Object{$_.trim()}
    }
    Get-SSHSession | Select-Object SessionId | Remove-SSHSession | Out-Null  #--[ Remove the open session ]--
    Return $Response
}

Function CallPlink ($IP,$command,$Plver,$DeviceType){
    $ErrorActionPreference = "silentlycontinue"
    $SwitchPlink = $False
    $UN = $Env:USERNAME
    $DN = $Env:USERDOMAIN
    $UID = $DN+"\"+$UN

    If ($Env:Path -notlike "*putty*"){
        StatusMsg "Putty was not found in your system path.  This script will probably fail." "Red" $Debug 
        sleep -sec 10
    }
    If (!(Test-Path -Path "$PSScriptRoot\plink-v52.exe" -PathType leaf)){
        StatusMsg "Plink v52 not found, copying to script folder..." "Magenta" $Debug
        Copy-item -Path "c:\program files\putty\plink-v52.exe" -Destination $PSScriptRoot
    }
    If (!(Test-Path -Path "$PSScriptRoot\plink-v73.exe" -PathType leaf)){
        StatusMsg "Plink v73 not found, copying to script folder..." "Magenta" $Debug
        Copy-item -Path "c:\program files\putty\plink-v73.exe" -Destination $PSScriptRoot
    }
    If ($DeviceType -eq "WAP"){
        $Username = $ExtOption.WAPUser 
        $Password = $ExtOption.WAPPass 
    }Else{
        $PasswordFile = $ExtOption.PasswordFile 
        $KeyFile = $ExtOption.KeyFile 
        If (Test-Path -Path $PasswordFile){
            $Base64String = (Get-Content $KeyFile)
            $ByteArray = [System.Convert]::FromBase64String($Base64String)
            $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UID, (Get-Content $PasswordFile | ConvertTo-SecureString -Key $ByteArray)
        }Else{
            $Credential = $Script:ManualCreds
        }
        #------------------[ Decrypted Result ]-----------------------------------------
        $Password = $Credential.GetNetworkCredential().Password
        $Domain = $Credential.GetNetworkCredential().Domain
        $Username = $Domain+"\"+$Credential.GetNetworkCredential().UserName
    }

    If ($PlinkVControl -contains $IP){  #--[ Check lookup table for older plink version requirement ]--
        $SwitchPlink = $True
        StatusMsg "Switching Plink version" "Magenta" $Debug
    }

    #--[ Always store or Update SSH key in CurrentUser PuTTY registry key ]--
    StatusMsg "Automatically storing or updating SSH key..." "Magenta"
    If ($SwitchPlink){
        write-output "y" | plink-v52.exe -ssh $username@$IP 'exit' *>&1 | Out-Null
        Start-Sleep -Milliseconds 500
        $test = @(plink-v52.exe -ssh -no-antispoof -batch -pw $Password $username@$IP $command *>&1)   
    }Else{
        write-output "y" | plink-v73.exe -ssh $username@$IP 'exit' *>&1 | Out-Null
        Start-Sleep -Milliseconds 500
        $test = @(plink-v73.exe -ssh -no-antispoof -pw $Password $username@$IP $command *>&1)
    }
    #------------------------------------------------------------

    If ($test -like "*The server's host key is not cached in the registry*"){
        StatusMsg "Target SSH key not currently in system registry. --- This host will fail ---" "red"
    }
    If ($test -like "*Keyboard-interactive*"){
        StatusMsg "Target SSH key has been stored." "Magenta" $Debug
    }
    If ($test -like "*abandoned*"){
        StatusMsg "Switching Plink version" "Magenta" $Debug
        $SwitchPlink = $true
    }

    StatusMsg "Plink IP: $IP" "Magenta" $Debug
    If ($SwitchPlink){
        $Msg = 'Executing Plink v52 (Command = '+$Command+')'
        StatusMsg $Msg 'blue' $Debug
        $Result = @(plink-v52.exe -ssh -no-antispoof -batch -pw $Password $username@$IP $command *>&1) 
    }Else{
        $ErrorActionPreference = "silentlycontinue"
        $Msg = 'Executing Plink v73 (Command = '+$Command+')'
        StatusMsg $Msg 'magenta' $Debug
        $Result = @(plink-v73.exe -ssh -no-antispoof -batch -pw $Password $username@$IP $command *>&1)
    }

    ForEach ($Line in $Result){
        If ($Line -like "*denied*"){
            $Result = "ACCESS-DENIED"
            Break
        } 
    }
    StatusMsg "Data collected..." "Magenta" $Debug
    Return $Result
    Start-Sleep -millisec 500
} 

Function GetSource ($SourcePath,$ExcelSourceFile,$ExcelWorkingCopy){
    StatusMsg "Excel working copy was not found, copying from source..." "Magenta"
    If (Test-Path -Path $ExcelSourceFile -PathType Leaf){
        Try{
            Copy-Item -Path $ExcelSourceFile -Destination $ExcelWorkingCopy -force -ErrorAction:Stop
            Return $True
        }Catch{
            write-host $_.Exception.Message
            write-host $_.Error.Message
            Return $False   
            StatusMsg "Copy failed... " "red" 
        }
    }Else{   
        StatusMsg "Source file check failed... " "red"
        Return $False
    }
}

Function GetEOLDate ($ModelNum){    #--[ Formatted as EOL,EOS,LDOS.  Note that FULL model number is used in lookup ]--
    Switch ($ModelNum) {
        "ISR4331/K9" {Return "04/30/2019,10/29/2019,10/31/2024";Break}              # https://www.cisco.com/c/en/us/products/collateral/routers/4221-integrated-services-router-isr/eos-eol-notice-c51-742328.html
        "CISCO2911/K9" {Return "09/09/2016,12/09/2017,12/31/2022";Break}            # https://www.cisco.com/c/en/us/products/collateral/routers/2900-series-integrated-services-routers-isr/eos-eol-notice-c51-737831.html
        "CISCO43" {Return "11/09/2020,05/10/2021,05/21/2026";Break}                 # https://www.cisco.com/c/en/us/products/collateral/routers/4000-series-integrated-services-routers-isr/eos-eol-notice-c51-744454.html
        "C9300-48P" {Return "N/A,N/A,N/A";Break}                                    # https://www.cisco.com/c/en/us/support/switches/catalyst-9300-series-switches/series.html#~tab-documents
        "C9300-NW-A-24" {Return "N/A,N/A,N/A";Break}
        "C9300-NW-A-48" {Return "N/A,N/A,N/A";Break}
        "C9200CX-12P-2X2G" {Return "N/A,N/A,N/A";Break}
        "C6807-XL" {Return "05/02/2013,04/30/2022,04/30/2027";Break}                # https://www.cisco.com/c/en/us/products/collateral/switches/catalyst-6800-series-switches/eos-eol-notice-c51-2438805.html
        "WS-C3560C-8PC-S" {Return "10/31/2015,10/30/2016,10/31/2021";Break}         # https://www.cisco.com/c/en/us/products/collateral/switches/catalyst-3560-c-series-switches/eos-eol-notice-c51-736180.html
        "WS-C3560C-12PC-S" {Return "10/31/2015,10/30/2016,10/31/2021";Break}        # https://www.cisco.com/c/en/us/products/collateral/switches/catalyst-3560-c-series-switches/eos-eol-notice-c51-736180.html
        "WS-C3560G-24PS-S" {Return "01/31/2012,01/30/2013,01/31/2018";Break}        # https://www.cisco.com/c/en/us/products/collateral/switches/catalyst-3750-series-switches/eol_c51-696372.html
        "WS-C3560V2-24PS-S" {Return "11/14/2013,05/14/2016,05/31/2021";Break}       # https://www.cisco.com/c/en/us/products/collateral/switches/catalyst-3750-series-switches/eos-eol-notice-c51-730227.html
        "WS-C3560-12PC-S" {Return "10/31/2015,10/30/2016,10/31/2021";Break}         # https://www.cisco.com/c/en/us/products/collateral/switches/catalyst-3560-series-switches/eol_c51_519208.html
        "WS-C3560-24PS-S" {Return "1/4/2010,7/5/2010,7/31/2015";Break}              # https://www.cisco.com/c/en/us/products/collateral/switches/catalyst-3650-series-switches/eos-eol-notice-c51-744426.html
        "WS-C3560-48PS-A" {Return "1/4/2010,7/5/2010,7/31/2015";Break}              # https://www.cisco.com/c/en/us/products/collateral/switches/catalyst-3650-series-switches/eos-eol-notice-c51-744426.html
        "WS-C3560-48PS-S" {Return "1/4/2010,7/5/2010,7/31/2015";Break}              # https://www.cisco.com/c/en/us/products/collateral/switches/catalyst-3650-series-switches/eos-eol-notice-c51-744426.html
        "WS-C3560-8PC-S" {Return "10/31/2015,10/30/2016,10/31/2021";Break}          # https://www.cisco.com/c/en/us/products/collateral/switches/catalyst-3650-series-switches/eos-eol-notice-c51-744426.html
        "WS-C3560CX-8PC-S" {Return "05/01/2023,04/30/2024,04/30/2029";Break}        # https://www.cisco.com/c/en/us/products/collateral/switches/catalyst-3560-cx-series-switches/catalyst-3560-cx-serie-switche-eol.html
        "WS-C3560CX-12PC-S" {Return "05/01/2023,04/30/2024,04/30/2029";Break}       # https://www.cisco.com/c/en/us/products/collateral/switches/catalyst-3560-cx-series-switches/catalyst-3560-cx-serie-switche-eol.html 
        "WS-C3650-48PS" {Return "10/31/2020,10/31/2021,10/31/2026";Break}    
        "WS-C3750-V2-48PS" {Return "11/14/2013,5/14/2016,5/31/2021";Break}          # https://www.cisco.com/c/en/us/products/collateral/switches/catalyst-3750-series-switches/eol_c51-696372.html
        "WS-C3750-48PS-S" {Return "1/4/2010,7/5/2010,7/31/2015";Break}
        "WS-C3750V2-48PS-S" {Return "11/14/1013,5/14/2016,5/31/2021";Break}         # https://www.cisco.com/c/en/us/products/collateral/switches/catalyst-3750-series-switches/eos-eol-notice-c51-730227.html
        "WS-C3750X-24P-S" {Return "11/14/1013,5/14/2016,5/31/2021";Break}
        "WS-C3850-12XS" {Return "10/31/2019,10/31/2020,10/31/2025";Break}           # https://www.cisco.com/c/en/us/products/collateral/switches/catalyst-3850-series-switches/eos-eol-notice-c51-743072.html
        "WS-C3850-24P" {Return "10/31/2019,10/31/2020,10/31/2025";Break}
        "WS-C3850-48P" {Return "10/31/2019,10/31/2020,10/31/2025";Break}
        "WS-C6509-V-E" {Return "10/31/2019,10/30/2020,10/31/2025";Break}
        "WS-C6513" {Return "8/5/2011,8/4/2012,8/31/2017";Break}
        "WS-C4510R+E" {Return "10/31/2019,10/30/2020,10/31/2025";Break}             # https://www.cisco.com/c/en/us/products/collateral/switches/catalyst-4500-series-switches/eos-eol-notice-c51-743088.html
        "AIR-CAP3702I-A-K9" {Return "04/30/2018,04/30/2019,04/30/2024";Break}       # https://www.cisco.com/c/en/us/products/collateral/wireless/aironet-3700-series/eos-eol-notice-c51-740710.html
        "AIR-CAP3702I-B-K9" {Return "04/30/2018,04/30/2019,04/30/2024";Break}       # https://www.cisco.com/c/en/us/products/collateral/wireless/aironet-3700-series/eos-eol-notice-c51-740710.html
        "AIR-CAP3702I-N-K9" {Return "04/30/2018,04/30/2019,04/30/2024";Break}       # https://www.cisco.com/c/en/us/products/collateral/wireless/aironet-3700-series/eos-eol-notice-c51-740710.html
        "AIR-CAP3702E-A-K9"  {Return "04/30/2018,04/30/2019,04/30/2024";Break}       # https://www.cisco.com/c/en/us/products/collateral/wireless/aironet-3700-series/eos-eol-notice-c51-740710.html
        "AIR-AP2802I-B-K9" {Return "10/31/2021,05/21/2022,04/30/2027";Break}        # https://www.cisco.com/c/en/us/products/collateral/wireless/aironet-2800-series-access-points/aironet-2800-series-access-points-eol.html
        "C9120AXI-B" {Return "N/A,N/A,N/A";Break}                                   # https://www.cisco.com/c/en/us/products/collateral/wireless/aironet-2800-series-access-points/aironet-2800-series-access-points-eol.html
        "VG224" {Return "06/26/2003,06/26/2003,06/26/2008";Break}                   # https://www.cisco.com/c/en/us/obsolete/unified-communications/cisco-vg200-gateway.html
        Default {Return "N/A,N/A,N/A";Break}
    }
}

Function Open-Excel ($Excel,$ExcelWorkingCopy,$SheetName,$Console) {
    If (Test-Path -Path $ExcelWorkingCopy -PathType Leaf){
        $Script:NewSpreadsheet = $False
        $WorkBook = $Excel.Workbooks.Open($ExcelWorkingCopy)
        $WorkSheet = $Workbook.WorkSheets.Item($SheetName)
        $WorkSheet.activate()
    }Else{
        $Script:NewSpreadsheet = $True
        $Workbook = $Excel.Workbooks.Add()
        $Worksheet = $Workbook.Sheets.Item(1)
        $Worksheet.Activate()
        $WorkSheet.Name = $SheetName
        [int]$Col = 1
        [Int]$Row = 1
        $WorkSheet.cells.Item($Row,$Col++) = "LAN IP Address"  # A
        $WorkSheet.cells.Item($Row,$Col++) = "Hostname"        # B
        $WorkSheet.cells.Item($Row,$Col++) = "Connection"      # C
        $WorkSheet.cells.Item($Row,$Col++) = "Base MAC"        # D
        #--------------------------------------------------------
        $WorkSheet.cells.Item($Row,$Col++) = "Facility"        # E
        $WorkSheet.cells.Item($Row,$Col++) = "Address"         # F
        $WorkSheet.cells.Item($Row,$Col++) = "IDF"             # G
        $WorkSheet.cells.Item($Row,$Col++) = "Description"     # H
        $WorkSheet.cells.Item($Row,$Col++) = "Asset Tag"       # I
        #--------------------------------------------------------        
        $WorkSheet.cells.Item($Row,$Col++) = "Device Type"     # J
        $WorkSheet.cells.Item($Row,$Col++) = "Serial #"        # K
        $WorkSheet.cells.Item($Row,$Col++) = "Model #"         # L
        $WorkSheet.cells.Item($Row,$Col++) = "Mfg Date"        # M
        $WorkSheet.cells.Item($Row,$Col++) = "Age (Yrs)"       # N
        $WorkSheet.cells.Item($Row,$Col++) = "EOL Date"        # O
        $WorkSheet.cells.Item($Row,$Col++) = "EOS Date"        # P
        $WorkSheet.cells.Item($Row,$Col++) = "LDOS Date"       # Q
        $WorkSheet.cells.Item($Row,$Col++) = "Upgr Priority"   # R
        $WorkSheet.cells.Item($Row,$Col++) = "Processor"       # S
        $WorkSheet.cells.Item($Row,$Col++) = "RAM (MB)"        # T
        $WorkSheet.cells.Item($Row,$Col++) = "MB Serial #"     # U
        $WorkSheet.cells.Item($Row,$Col++) = "Firmware Ver"    # V
        $WorkSheet.cells.Item($Row,$Col++) = "Firmware Rel"    # W
        $WorkSheet.cells.Item($Row,$Col++) = "Firmware Family" # X
        $WorkSheet.cells.Item($Row,$Col++) = "Firmware Base"   # Y
        $WorkSheet.cells.Item($Row,$Col++) = "Last Reload"     # Z
        $WorkSheet.cells.Item($Row,$Col++) = "Days Up"         # AA
        $WorkSheet.cells.Item($Row,$Col++) = "Port Speed"      # AB
        Switch ($DeviceType){
            "Router" {
                $WorkSheet.cells.Item($Row,$Col++) = "WAN CID"     # AC
                $WorkSheet.cells.Item($Row,$Col++) = "WAN IP"      # AD
                $WorkSheet.cells.Item($Row,$Col++) = "PRI/ATM/FXO/FXS/VPN/SER"   # AE
                $WorkSheet.cells.Item($Row,$Col++) = "Date Inspected"      # AF
                $Worksheet.Name = "Routers"
                $Range = $WorkSheet.Range(("A1"),("AF1")) 
            }
            "WAP" {
                $WorkSheet.cells.Item($Row,$Col++) = "Host Switch"      # AC
                $WorkSheet.cells.Item($Row,$Col++) = "Host Sw Port"     # AD
                $WorkSheet.cells.Item($Row,$Col++) = "Date Inspected"   # AE
                $Worksheet.Name = "Wireless AP"
                $Range = $WorkSheet.Range(("A1"),("AE1")) 
            }
            "Switch" {
                $WorkSheet.cells.Item($Row,$Col++) = "Port Count"       # AC
                $WorkSheet.cells.Item($Row,$Col++) = "Stack Sw #"       # AD   
                $WorkSheet.cells.Item($Row,$Col++) = "Date Inspected"   # AE          
                $Worksheet.Name = "Switches"
                $Range = $WorkSheet.Range(("A1"),("AE1")) 
            }
            "VG" {
                $WorkSheet.cells.Item($Row,$Col++) = "Port Count"       # AC
                $WorkSheet.cells.Item($Row,$Col++) = "Stack Sw #"       # AD    
                $WorkSheet.cells.Item($Row,$Col++) = "Date Inspected"   # AE         
                $WorkSheet.Name = "Voice Gateways (VG)"
                $Range = $WorkSheet.Range(("A1"),("AE1")) 
            }
        }
        $Range.font.bold = $True
        $Range.HorizontalAlignment = -4108  #Alignment Middle
        $Range.Font.ColorIndex = 44
        $Range.Font.Size = 12
        $Range.Interior.ColorIndex = 56
        $Range.font.bold = $True
        1..4 | ForEach-Object {
            $Range.Borders.Item($_).LineStyle = 1
            $Range.Borders.Item($_).Weight = 4
        }
        $Resize = $WorkSheet.UsedRange
        [Void]$Resize.EntireColumn.AutoFit()
    }
    Return $WorkSheet
}

#=[ End of Functions ]====================================================

#=[ Lookup Tables ]=======================================================
$PlinkVControl = @()
$PlinkVControl += "10.0.40.2"
$PlinkVControl += "10.0.40.3"

$MfgDateCodes = @{
    "01" = "1997"; 
    "02" = "1998";
    "03" = "1999"; 
    "04" = "2000"; 
    "05" = "2001"; 
    "06" = "2002"; 
    "07" = "2003";
    "08" = "2004"; 
    "09" = "2005"; 
    "10" = "2006"; 
    "11" = "2007"; 
    "12" = "2008";
    "13" = "2009"; 
    "14" = "2010"; 
    "15" = "2011"; 
    "16" = "2012";
    "17" = "2013";
    "18" = "2014";
    "19" = "2015";
    "20" = "2016";
    "21" = "2017";
    "22" = "2018";
    "23" = "2019";
    "24" = "2020";
    "25" = "2021";
    "26" = "2022";
    "27" = "2023";
    "28" = "2024";
    "29" = "2025"
}

$PortSpeed = @{
    "C6807-XL" = "10/100/1000/TG";
    "WS-C6509-V-E" = "10/100/1000";
    "WS-C3560-8PC-S" = "10/100";
    "WS-C3560-48PS-S" = "10/100";
    "WS-C3560CX-12PC-S" = "10/100/1000";
    "WS-C3560CX-8PC-S" = "10/100/1000";
    "WS-C3650-48PS" = "10/100/1000";
    "WS-C3560-12PC-S" = "10/100";
    "WS-C3750-48PS-S" = "10/100/1000";
    "WS-C3750V2-48PS-S" = "10/100/1000";
    "WS-C3750X-24P-S" = "10/100/1000";
    "WS-C3850-12XS" = "10/100/1000";
    "WS-C3850-48P" = "10/100/1000";
    "WS-C4510R+E" = "10/100/1000";
    "C9200CX-12P-2X2G" = "10/100/1000";
    "C9300-48P" = "10/100/1000";
    "C9300-24P" = "10/100/1000";
    "AIR-CAP3702I-A-K9" = "10/100/1000";
    "AIR-CAP3702I-B-K9" = "10/100/1000";
    "AIR-CAP3702I-N-K9" = "10/100/1000";
    "AIR-CAP3702E-A-K9" = "10/100/1000";
    "AIR-AP2802I-B-K9" = "10/100/1000";
    "C9120AXI-B" = "10/100/1000"
}
#=[ End of Lookup Tables ]====================================================

#=[ Begin Processing ]========================================================

#--[ Load external XML options file ]------------------------------------------------
$ConfigFile = $PSScriptRoot+"\"+($MyInvocation.MyCommand.Name.Split("_")[0]).Split(".")[0]+".xml"
If (Test-Path $ConfigFile){                          #--[ Error out if configuration file doesn't exist ]--
    [xml]$Config = Get-Content $ConfigFile           #--[ Read & Load XML ]--  
    $ExtOption = LoadConfig $Config 
}Else{
    LoadConfig "failed"
    StatusMsg "MISSING XML CONFIG FILE.  File is required.  Script aborted..." " Red" $Debug
    break;break;break
}

If ($Debug){
#    $ExtOption            #--[ In debug mode write external config settings out to console ]--                        
}

If ($Null -eq $ExtOption.PasswordFile){
    $Script:ManualCreds = Get-Credential -Message 'Enter an appropriate Domain\User and Password to continue.'
}

#--[ Spreadsheet Processing ]-----------------------------------------------
$ExcelWorkingCopy = $PSScriptRoot+"\"+$ExtOption.ExcelWorkingCopy
$ExcelSourceFile = $ExtOption.SourcePath+"\"+$ExtOption.ExcelSourceFile

#--[ Kill any instances of Excel opened by PowerShell to avoid issues ]--
$ProcID = Get-CimInstance Win32_Process | where {$_.name -like "*excel*"}
ForEach ($ID in $ProcID){
    Foreach ($Proc in (get-process -id $id.ProcessId)){
        if (($ID.CommandLine -like "*/automation -Embedding") -Or ($proc.MainWindowTitle -like "$ExcelWorkingCopy*")){
            Stop-Process -ID $ID.ProcessId -Force
            StatusMsg "Killing any existing open instance of spreadsheet..." "Red" $Debug
            Start-Sleep -Milliseconds 100
        }
    }
}

Switch ($DeviceType){   #--[ Select which text file to use according to device type ]--
    "WAP" {
        StatusMsg "--- Processing Wireless Access Points ---" "Yellow" $Debug
        $ListFileName = "$PSScriptRoot\WAP-IPlist.txt"
        $SheetName = "Wireless AP"
        $IPList = @()
    }
    "VG" {
        StatusMsg "--- Processing Cisco Voice Gateways ---" "Yellow" $Debug
        $ListFileName = "$PSScriptRoot\VG-IPlist.txt"
        $SheetName = "Voice Gateways (VG)"
    }
    "Router" {
        StatusMsg "--- Processing Cisco Routers ---" "Yellow" $Debug
        $SheetName = "Routers"
        $ListFileName = "$PSScriptRoot\Router-IPlist.txt"
    }
    "Switch" {
        StatusMsg "--- Processing Cisco Switches ---" "Yellow" $Debug
        $SheetName = "Switches"
        $ListFileName = "$PSScriptRoot\Switch-IPlist.txt"
    }
    Default {
        StatusMsg "--- Processing Cisco Switches ---" "Yellow" $Debug
        $SheetName = "Switches"
        $ListFileName = "$PSScriptRoot\Switch-IPlist.txt"
    }
}
 
If ($EnableExcel){  
    $Excel = New-Object -ComObject Excel.Application                                #--[ Create new Excel COM object ]--
    StatusMsg "Preparing new Excel COM object..." "Magenta" $Debug
}

$IPList  = @()
If (Test-Path -Path $ListFileName){  #--[ Verify that a text file exists and pull IP's from it then create a new spreadsheet. ]--
    $LoadList = Get-Content $ListFileName  
    $Row = 2   
    StatusMsg "IP text list was found, loading IP list from it..." "green" $Debug
    If ($EnableExcel){  
        StatusMsg "Creating new Spreadsheet..." "green" $Debug
        $WorkSheet = Open-Excel $Excel "TempExcel" $SheetName $Console 
        $Row = 2
    }
    ForEach ($Item in $LoadList){              #--[ Clean up list prior to processing ]--
        #--[ Format is IP;facility;Address;IDF;description.  Semi-Colon delimited.  Lines starting with # are ignored. ]--
        If  ($Item.Split(";")[0] -ne "#"){
            If ($Item.Split(";").Count -gt 1){
                $IPList += $Item+";"+$Row      #--[ Processing a full list ]--
            }Else{
                $IPList += $Item+";;;;;"+$Row  #--[ Processing a list of only IP addresses ]--
            }
        }
        $Worksheet.Cells.Item($Row,1) = $Item
        $Row++
    }
    $Script:NewSpreadsheet = $True
    StatusMsg "Renaming IP text list file to .OLD..." "Magenta"
    Move-Item -Path $ListFileName -Destination ($ListFileName+".old") 
}Else{ #=======================================================================================================================
    $ErrorActionPreference = "stop"
    StatusMsg "IP text list not found, Attempting to process spreadsheet... " "cyan" $Console
    $Script:NewSpreadsheet = $False
    If (Test-Path -Path $ExcelWorkingCopy -PathType Leaf){
        StatusMsg "Spreadsheet working copy located." "Green" $Console
        If ($StayCurrent){
            StatusMsg "Renaming spreadsheet working copy to .BAK..." "Green" $Console
            #--[ Make a backup of the working copy, keep only the last 10 ]--
            $Backup = $PSScriptRoot+"\"+$DateTime+"_"+$ExtOption.ExcelWorkingCopy+".bak"
            Get-ChildItem -Path $PSScriptRoot | Where-Object {(-not $_.PsIsContainer) -and ($_.Name -like "*.bak")} | Sort-Object -Descending -Property LastTimeWrite | Select-Object -Skip 10 | Remove-Item
            Move-Item -Path $ExcelWorkingCopy -Destination $Backup 
            If (Test-Path -Path $ExcelSourceFile -PathType Leaf){                   #--[ Test for source file ]--
                StatusMsg "Master source file is available..." "Green" $Console
                Try{
                    StatusMsg "Attempting to copy new working file from master source..." "Magenta" $Console 
                    If ($Debug){                   
                        Copy-Item -Path $ExcelSourceFile -Destination $ExcelWorkingCopy -force -passthru -ErrorAction Stop | Out-Null
                    }Else{                 
                        Copy-Item -Path $ExcelSourceFile -Destination $ExcelWorkingCopy -force -ErrorAction Stop | Out-Null
                    }
                    If (Test-Path -Path $ExcelWorkingCopy -PathType Leaf){ 
                        StatusMsg "Verified new working copy of spreadsheet has been copied..." "Green" $Console
                        $WorkBook = $Excel.Workbooks.Open($ExcelWorkingCopy)
                    }
                    StatusMsg "Cleaning working copy... Removing all non-target worksheets..." "Green"
                    $Excel.displayalerts = $False
                    ForEach ($WorkSheet in $WorkBook.Worksheets){
                        If ($WorkSheet.Name -ne $SheetName){
                            $WorkSheet.Delete()
                        }
                    }
                }Catch{
                    StatusMsg "Unable to copy new local working copy from master..." "Red" $Console   
                    StatusMsg "Returning a copy of the recent backup... " "Green" $Console 
                    Copy-Item -Path $Backup -destination $ExcelWorkingCopy 
                }
                If (Test-Path -Path $ExcelWorkingCopy -PathType Leaf){ 
                    StatusMsg "Verified new working copy of spreadsheet has been copied..." "Green" $Console
                    $WorkBook = $Excel.Workbooks.Open($ExcelWorkingCopy)
                }Else{
                    StatusMsg "Working copy file was not found..." "Red" $Console
                    StatusMsg "Nothing to process.  Exiting... " "red" $Console 
                    Break;Break;Break
                }
            }Else{
                StatusMsg "Master source file not found..." "Yellow" $Console
                StatusMsg "Nothing to process.  Exiting... " "red" $Console 
                Break;Break;Break
            }
        }Else{
            StatusMsg "Using exisitng spreadsheet working copy." "Magenta" 
            $WorkBook = $Excel.Workbooks.Open($ExcelWorkingCopy)
        }    
    }
    $WorkSheet = $Workbook.WorkSheets.Item($SheetName)
    $WorkSheet.activate()  
    $Excel.DisplayAlerts = $false  
    $Excel.Visible = $True  
}

#--[ Generate IP list object from spreadsheet prior to processing ]--
$Row = 2   #--[ Row 1 is the header ]--
$IPList = @() 
$PreviousIP = ""
StatusMsg "Reading Spreadsheet..." "Magenta" $Debug
Do {
    $CurrentIP = $WorkSheet.Cells.Item($Row,1).Text   
    If ($CurrentIP -ne $PreviousIP){  #--[ Make sure IPs are added only once per switch stack ]--
        $Facility = $WorkSheet.Cells.Item($Row,5).Text
        $Address = $WorkSheet.Cells.Item($Row,6).Text
        $IDF = $WorkSheet.Cells.Item($Row,7).Text
        $Description = $WorkSheet.Cells.Item($Row,8).Text
        $IPList += ,@($CurrentIP+";"+$Facility+";"+$Address+";"+$IDF+";"+$Description+";"+$Row)
        $PreviousIP = $CurrentIP
    }
    $Selection = $WorkSheet.cells.Item($Row,1).EntireRow
    $Selection.Font.Bold = $False
    $Selection.Font.ColorIndex = 0
    $Row++
} Until (
    $WorkSheet.Cells.Item($Row,1).Text -eq ""   #--[ Condition that stops the loop if it returns true ]--
)
#$Excel.quit()  #--[ Close it.  Only do so if you are pulling the IP list and doing nothing else, otherwise bad things happen  ]--

#==[ Begin Processing of IP List ]=============================================
$Row = 2
$command = 'sh version'
[string]$Port = "23"
$ErrorActionPreference = "stop"

#--[ NOTE: If a line in the text file starts with "#," that line is ignored ]--
ForEach ($Line in $IPList ){ #| Where {($_.Split(",")[0].tostring()) -NotLike "#"}){
    $NoAccess = $False
    $IP = ($Line.Split(";")[0]) 
    $Facility = ($Line.Split(";")[1]) 
    $Address = ($Line.Split(";")[2]) 
    $IDF = ($Line.Split(";")[3]) 
    $Description = ($Line.Split(";")[4])

    #--[ The next line sets row color to pale blue to denote which row is being worked on ]--
    $Excel.ActiveSheet.UsedRange.Rows.Item($Row).Interior.ColorIndex = 20  

    $MfgDate = ""
    $HostName = ""
    $Result = ""
    $color = 1
    $SwitchNum = "Switch 01"
    $Connection = $False 
    $GigCount = 0
    $FeCount = 0
    $TenGCount = 0

    If ($Console){
        Write-host "`n--[ Current Device: "  -ForegroundColor yellow -NoNewline
        Write-Host $IP  -ForegroundColor cyan -NoNewline
        Write-Host " ]---------------------------------------------------" -ForegroundColor yellow 
    }
    
    StatusMsg "Current Line = $Line" "Magenta" $Debug
    StatusMsg "Spreadsheet row = $Row" "Magenta" $Debug
 
    #--[ Test and connect to target IP ]----------------------------------------------------------
    If ($IP -eq 1.1.1.1){  #}"10.10.10.1")){ #} -or ($IP -eq "10.10.40.6")){ # -or ($IP -eq "10.10.40.2")){    #--[ IP Exclusion List ]--        
        If ($Console){Write-Host "-- Script Status: Bypassing IP" -ForegroundColor Red}
    }Else{
        #--[ Test network connection.  Column 1 (A) ]----------------------------------------------
        If (Test-Connection -ComputerName $IP -count 1 -BufferSize 16 -Quiet){
            $Connection = $True
        }Else{
            Start-Sleep -Seconds 2
            If (Test-Connection -ComputerName $IP -count 1 -BufferSize 16 -Quiet){
                $Connection = $True
            }Else{
                StatusMsg "--- No Connection ---" "Red" $Debug
                If ($EnableExcel){
                    $WorkSheet.Cells.Item($Row, 1) = $IP 
                    $WorkSheet.Cells.Item($Row, 3) = "No Connection"         
                }
            }
        }

        If ($Connection){
            $Obj01 = New-Object -TypeName psobject   #--[ Collection for first device ]--
            $Obj = $Obj01        
            $Obj | Add-Member -MemberType NoteProperty -Name "Today" -Value $Today -force              #-------------[ Date of run ]-------------
            $Obj | Add-Member -MemberType NoteProperty -Name "Connection" -Value "OK" -force           #-------------[ Connection ]-------------
            $Obj | Add-Member -MemberType NoteProperty -Name "IPAddress" -Value $IP -force             #-------------[ IP Address ]-------------        
            $Obj | Add-Member -MemberType NoteProperty -Name "Facility" -Value $Facility -force        #-------------[ Facility ]---------------
            $Obj | Add-Member -MemberType NoteProperty -Name "Address" -Value $Address -force          #-------------[ Address ]----------------        
            $Obj | Add-Member -MemberType NoteProperty -Name "IDF" -Value $IDF -force                  #-------------[ IDF ]--------------------
            $Obj | Add-Member -MemberType NoteProperty -Name "Description" -Value $Description -force  #-------------[ Description ]------------
            $MfgDate = ""
            $Age = ""
            $BaseMac = ""
           
            If (Test-Path -Path 'C:\Program Files\PuTTY\'){ 
                Switch ($DeviceType){
                        "WAP" {
                            $Obj | Add-Member -MemberType NoteProperty -Name "DeviceType" -Value "Wireless AP" -force
                            #--[ 1st Command ]--
                            $command = 'sh version'
                            $Result = CallPlink $IP $command "" $DeviceType
                            Start-Sleep -Milliseconds 500
                            #--[ 2nd Command ]--
                            If ($Result -like "*can't get tty*"){
                                $command = 'sh run'
                                $Result2 = CallPlink $IP $command "" $DeviceType
                                Start-Sleep -Milliseconds 500
                            }
                            #--[ 3rd Command ]--
                            $command = 'sh cdp neighbors'
                            $CDPResult = CallPlink $IP $command "" $DeviceType
                            ForEach ($Line in $CDPResult){
                                If ($Line -like ("*"+$Domain+"*")){
                                    $Obj | Add-Member -MemberType NoteProperty -Name "HostSwitch" -Value ($Line.Split(".")[0]) -force
                                }
                                If ($Line -like ("*Gig*")){
                                    $Obj | Add-Member -MemberType NoteProperty -Name "HostSwPort" -Value ("Gi"+($Line.Split(" ")[-1])) -force
                                }
                            }
                        }
                        "VG" {
                            $Obj | Add-Member -MemberType NoteProperty -Name "DeviceType" -Value "Voice Gateway" -force
                            $Data = ""
                            #  $Result = @(Get-Telnet -RemoteHost $IP -Commands "user","en","password","sh ver","`l")   
                            #--[ Telnet option format ]-------
                            #$Commands = "user","en","password","sh ver"," "   
                            $socket = new-object Net.Sockets.TcpClient 
                            $socket.Connect($IP, $Port)
                            #$Socket = New-Object System.Net.Sockets.TcpClient($RemoteHost, $Port)
                            If ($Socket){
                                $Stream = $Socket.GetStream()
                                $Writer = New-Object System.IO.StreamWriter($Stream)
                                $Buffer = New-Object System.Byte[] 2048 #1024 
                                $Encoding = New-Object System.Text.AsciiEncoding    
                                #--[ Start issuing the commands ]--------------
                                ForEach ($Command in $Commands){
                                    $Writer.WriteLine($Command) 
                                    $Writer.Flush()
                                    Start-Sleep -Milliseconds 1000 # $WaitTime
                                }
                                Start-Sleep -Milliseconds 4000 # ($WaitTime * 4)
                                While($Stream.DataAvailable){
                                    $Read = $Stream.Read($Buffer, 0, 2048) 
                                    $Data += ($Encoding.GetString($Buffer, 0, $Read))
                                }
                                $Result = $Data.Split("`n")  #--[ Convert the bulk text dump to an array ]--
                            }Else{
                                If ($Console){Write-Host "Unable to connect to host: $($RemoteHost):$Port"}
                            }
                        }
                        Default {                             
                            $ErrorActionPreference = "stop"
                            If ($DeviceType -eq "Router"){
                                $Obj | Add-Member -MemberType NoteProperty -Name "DeviceType" -Value "Router" -force
                                #--[ Pull Router Version Data ]--------------------
                                $command = 'sh version'
                                $Result = CallPlink $IP $command "72" $DeviceType
                                Start-Sleep -sec 2
                                #--[ Pull Router Serial and Model ]----------------
                                $Command = "sh license udi"
                                $UID = CallPlink $IP $command
                                Start-Sleep -sec 2
                                $UID = ($UID -replace '\s+', ' ').Split(' ')[-1] 
                                #--[ Pull Router Interface Data ]------------------
                                $Command = "sh int | incl bia"
                                $MacList = CallPlink $IP $command "" $DeviceType
                                Start-Sleep -sec 2
                                #--[ Pull Router Runtime Data ]--------------------
                                $Command = "sh run"
                                $Runtime = CallPlink $IP $command "" $DeviceType
                                Start-Sleep -sec 2
                                $CID = ""
                                ForEach ($Line in $Runtime){  #--[ Extract Router WAN IP and Circuit ID ]--
                                    #--[ Note that description must be formatted properly ]--
                                    #--[ Example: description WAN CARRIER:ATT TYPE=ASEoD SPEED=100Mbps SITE=ABCD CID=AS/ABCD/1234//PT ]--
                                    If ($Line -like "*CARRIER*"){
                                        $CID = ($Line.Split("=")[-1]).Trim()
                                        $Obj | Add-Member -MemberType NoteProperty -Name "CID" -Value $CID -force  #-------------[ Router Circuit ID Number ]-----------------
                                    }
                                    
                                    #--[ Note that in order to parse the correct IP you must delimit it somehow or the very 1st IP is tagged ]--
                                    If ($Line -like "*address 10.2*"){
                                        $WanIP = ($Line.Split(" ")[3]).Trim()
                                        $Obj | Add-Member -MemberType NoteProperty -Name "WanIP" -Value $WanIP -force  #-------------[ Router WAN IP ]-----------------
                                        Break
                                    }
                                }
                                ForEach ($Line in $MacList){
                                    If (($Line -like "*ISR43*") -And ($BaseMAC -eq "")){    #--[ Filter MAC for ISR 4300 routers ]--
                                        $BaseMAC = ($Line.Split("(")[1].Split(" ")[1]).Replace(")", '')
                                        Break
                                    }ElseIf (($Line -like "*Hardware*") -And ($Line -notlike "*000.000*") -And ($BaseMAC -eq "")){
                                        $BaseMAC = ($Line.Split("(")[1].Split(" ")[1]).Replace(")", '')
                                        #$BaseMac = $BaseMac -replace '..(?!$)', '$&:'      #--[ Regex converts string to MAC format ]--
                                        Break
                                    }                          
                                }
                                $BaseMac = $BaseMAC.Replace('.', '')
                                $BaseMac = $BaseMAC -replace '..(?!$)', '$&:'
                                $Obj | Add-Member -MemberType NoteProperty -Name "BaseMAC" -Value $BaseMAC -force
                            }Else{ 
                                $Obj | Add-Member -MemberType NoteProperty -Name "DeviceType" -Value "Switch" -force
                                StatusMsg "Calling PLINK" "Magenta"
                                $command = 'sh version'
                                $Result = CallPlink $IP $command "" $DeviceType
                                #--[ New Posh-SSH call for future use ]--
                                #$Response = GetSSH $IP $Command $Credential
                                #----------------------------------------
                            }
                        }
                    } 
                }Else{
                    Write-Host "Cannot find PLINK.EXE.   Aborting..." -ForegroundColor Red
                    break;break
                }
            #}
            #----------------------------------------------------------------------------------------------

            If ($Result -eq 'ACCESS-DENIED'){  #--[ Bad user or password or some other access error ]--
                StatusMsg  "--- NO ACCESS ---   " "red"
                $NoAccess = $True
            }
            If ($DeviceType -eq "Router"){
                $SerialNum = $UID.split(":")[1]
                $Obj | Add-Member -MemberType NoteProperty -Name "SerialNum" -Value $SerialNum -force  #-------------[ Router Serial Number ]-----------------
                $ModelNum = $UID.Split(":")[0]
                $Obj | Add-Member -MemberType NoteProperty -Name "ModelNum" -Value $ModelNum -force    #-------------[ Router Model Number ]-----------------
            }
            $PRI = "0"
            $ATM = "0"
            $FXO = "0"
            $FXS = "0"
            $VPN = "0"
            $DeviceCount = 1       
            $Age = 0
            #$PortSpd = "No model reference found"
            $SwitchNum = "Switch 01"
            If ($Result -eq ""){
                $HostName = "NO DATA COLLECTED"
                $PortSpd = "Unknown"
                $MfgDate = "Unknown"
            }#>

            #==[ Parse Main Result Variables ]=============================================================
            ForEach ($Line in $Result){    
                Start-Sleep -Milliseconds 50  #--[ Slow things down a bit to avoid data skewing ]--      
                $Total = 0
                If (($Line -like "*uptime is*") -or ($Line -like "*switch uptime*")) {
                    If ($Line -notlike "*:*"){   
                        $Hostname = $Line.TrimStart().split(" ")[0].Trim()
                        $Obj | Add-Member -MemberType NoteProperty -Name "Hostname" -Value $Hostname -force  #-------------[ Hostname ]-----------------
                    }
                    $Line = $Line.Replace("uptime is","") 
                    $Line = $Line.Replace($Hostname,"") 
                    $Line = $Line.Replace("Switch uptime","")
                    $Line = $Line.Replace(":","") 
                    $Line = $Line.TrimStart()
                    $split = $Line.Split(",")
                    $Years = 0
                    $Weeks = 0
                    $Days = 0    
                    $Hours = 0
                    #$Minutes = 0  #--[ Not used ]--
                    ForEach ($item in $split){
                        $ErrorActionPreference = "silentlycontinue"
                        If ($item -like "*year*"){
                            [int]$Years = $item.Split(" ")[0]     
                            [int]$YearHours = $Years*8760
                            $Obj | Add-Member -MemberType NoteProperty -Name "YearsUp" -Value $Years -force  #-------------[ Years Up ]-----------------                
                            $Total = [int]$Total+$YearHours
                        }
                        If ($item -like "*week*"){  
                            If ($item -like "*:*"){ 
                                [int]$Weeks = ($item.trimstart().Split(":")[0]).Split(" ")[0]
                                [int]$WeekHours = $Weeks*168
                            }ElseIf ($Item -notlike ("*switch*")){
                                [int]$Weeks = ($item.TrimStart().Split(" ")[0])
                                $WeekHours = [int]$Weeks*168
                            }
                            $Obj | Add-Member -MemberType NoteProperty -Name "WeeksUp" -Value $Weeks -force  #-------------[ Weeks Up ]-----------------                
                            $total = $Total+$WeekHours
                        }
                        If ($item -like "*day*"){
                            If (($item.Split(" ")[-2]) -match "^\d+$"){
                                [int]$Days = $item.Split(" ")[-2]
                            }Else{
                                [int]$Days = $item.trimstart().Split(" ")[0]
                            }
                            $Obj | Add-Member -MemberType NoteProperty -Name "DaysUp" -Value $Days -force  #-------------[ Days Up ]-----------------                
                            $Total = [int]$Total+($Days*24)
                        }
                        If ($item -like "*hour*"){  #--[ Detect hour count ]--
                            [int]$Hours = ($item.TrimStart().Split(" ")[0])
                            $Total = [int]$Total+$Hours
                            $Obj | Add-Member -MemberType NoteProperty -Name "HoursUp" -Value $Hours -force  #-------------[ Hours Up ]-----------------           
                        }
                    }    
                    $TotalUpDays = [math]::Round([int]$Total/24,1)  #(([int]$Years)*365)+(([int]$Weeks)*7)+([int]$Days)
                    $ErrorActionPreference = "stop"
                    If ($DeviceType -ne "VG"){
                        If (!($Obj.Hostname)){
                            $Obj | Add-Member -MemberType NoteProperty -Name "Hostname" -Value $Hostname  #------------[ Hostname for non VG ]---------------
                        }
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "TotalDaysUp" -Value $TotalUpDays -force  #--------------[ Days Up ]---------------
                }

                If ($DeviceType -eq "VG"){
                    If (($Line -like "*#*") -and ($HostName -eq "")){            
                        $Hostname = $Line.Split("#")[0]
                        If (!($Obj.Hostname)){
                            $Obj | Add-Member -MemberType NoteProperty -Name "Hostname" -Value $Script:Hostname -force  #----------[ Hostname for VG ]---------------
                        }
                    }
                }

                If (($DeviceType -eq "Router") -or ($DeviceType -eq "VG")){
                    If ($Line -like "*/PRI*"){
                        $PRI = $Line.Split(" ")[0]
                    }
                    If ($Line -like "*ATM*"){
                        $ATM = $Line.Split(" ")[0]
                    }
                    If ($Line -like "*FXO*"){
                        $FXO = $Line.Split(" ")[0]
                    }
                    If ($Line -like "*FXS*"){
                        $FXS = $Line.Split(" ")[0]
                    }
                    If ($Line -like "*(VPN)*"){
                        $VPN = $Line.Split(" ")[0]
                    }
                    If ($Line -like "*Serial*"){
                        $Serial = $Line.Split(" ")[0]
                    }
                    If ($Line -like "*Gigabit Ethernet*"){
                        $PortSpd = "1000"
                    }
                    If ($Line -like "*FastEthernet*"){
                        $PortSpd = "100"
                    }
                    $RtrPorts = $PRI+"/"+$ATM+"/"+$FXO+"/"+$FXS+"/"+$VPN+"/"+$Serial
                    $Obj | Add-Member -MemberType NoteProperty -Name "RtrPorts" -Value $RtrPorts -Force  #-------------[ Router Ports ]-----------------
                }Else{                    
                    If ($Line -like "*cisco ws-*"){
                        $ModelNum = $Line.Split(" ")[1].Trim()                                            
                    }ElseIf ($Line -like "*Model number*") {
                        $ModelNum = $Line.Split(":")[1].Trim()
                        If ($DeviceType -eq "WAP"){
                            $PortCount = 2
                        }Else{
                            If (($Line.Split("-").Count) -gt 2){
                                $PortCount = $Line.Split("-")[2].Split("P")[0]
                            }Else{
                                $PortCount = $Line.Split("-")[1].Split("P")[0]                   
                            }
                        }
                        If ($DeviceType -eq "WAP"){
                            $Obj | Add-Member -MemberType NoteProperty -Name "PortCount" -Value 2 -force  #---------------[ Port Count ]-----------------             
                        }Else{
                            $Obj | Add-Member -MemberType NoteProperty -Name "PortCount" -Value $PortCount -force  #------[ Port Count ]-----------------             
                        }
                        $Obj | Add-Member -MemberType NoteProperty -Name "ModelNum" -Value $ModelNum -force  #-------------[ Model Number 2nd Possible ID ]--------------- 
                    }
                    if($PortSpeed.ContainsKey([string]$ModelNum)){
                        $PortSpd = $PortSpeed[[string]$ModelNum]
                    }Else{
                        $PortSpd = "No model reference found"
                    } 
                   
                    $Obj | Add-Member -MemberType NoteProperty -Name "PortSpeed" -Value $PortSpd -force  #-------------[ Port Speed ]-----------------  
                   
                }#>
                
                If (($Line -like "System serial*") -or ($Line -like "Top Assembly Serial Number*")){  
                    #--[ Column 6 (F) Variation 1 ]-------------------------------------------------------------------------------------------
                    $CodedDate = (($Line.Split(":")[1]).TrimStart()).SubString(3,2)   
                    $MfgDate = $MfgDateCodes[[string]$CodedDate]
                    $SerialNum = ($Line.Split(":")[1]).Split(" ")[1].Trim()
                    $Obj | Add-Member -MemberType NoteProperty -Name "SerialNum" -Value $SerialNum -force 
                }

                If ($Line -like "Processor board ID*"){  
                    #--[ Column 6 (F) Variation 2 ]-------------------------------------------------------------------------------------------       
                    $CodedDate = (($Line.Split(" ")[3]).TrimStart()).SubString(3,2) 
                    $MfgDate = $MfgDateCodes[[string]$CodedDate]
                    $SerialNum = ($Line.Split(" ")[3]).Trim()
                    $Obj | Add-Member -MemberType NoteProperty -Name "SerialNum" -Value $SerialNum -force 
                }

                If (($DeviceType -eq "Router") -or ($DeviceType -eq "VG")){
                    #--[ Possible router date formats ]--
                    # Copyright (c) 1986-2016 by Cisco Systems, Inc.
                    # Cisco IOS-XE software, Copyright (c) 2005-2016 by cisco Systems, Inc.
                    If ($Line -like "*t (c)*"){
                        $MfgDate = (($Line.Split(')')[1]).Split('-')[1]).Split(" ")[0]
                    }    
                }
                $Obj | Add-Member -MemberType NoteProperty -Name "MfgDate" -Value $MfgDate -force

                #--[ Hardware Details ]-------------------------------------------------------------------------------------------------------
                If ($Line -like "Motherboard Serial*"){
                    $MBSerial = $Line.Split(":")[1].Trim()
                    $Obj | Add-Member -MemberType NoteProperty -Name "MBSerialNum" -Value $MBSerial -force
                }  
                If ($Line -like "Processor Board ID*"){
                    $MBSerial = $Line.Split(" ")[3].Trim()
                    $Obj | Add-Member -MemberType NoteProperty -Name "MBSerialNum" -Value $MBSerial -force
                }
    
                If ($Line -like "Base ethernet MAC Address*"){  
                    $BaseMAC = ($Line.subString($Line.length-17)).ToUpper()
                    $Obj | Add-Member -MemberType NoteProperty -Name "BaseMAC" -Value $BaseMAC -force
                }

                #--[ Processor & Model ]-----------------------------------------------------
                If (($Line -like "*) processor*") -or ($Line -like "*CPU at*")){ 
                    #--[ Possible variations ]--
                    # cisco WS-C4510R+E       (P5040)    processor (revision 2)  with 4194304K         bytes of physical memory.
                    # cisco WS-C3850-48P      (MIPS)     processor               with 4194304K         bytes of physical memory.
                    # cisco WS-C3560CX-8PC-S  (APM86XXX) processor (revision L0) with 524288K          bytes of memory.
                    # cisco C9300-48P         (X86)      processor               with 1301503K/6147K   bytes of memory.
                    # cisco AIR-CAP3702I-A-K9 (PowerPC)  processor (revision A0) with 376814K/134656K  bytes of memory.
                    # cisco C9200CX-12P-2X2G  (ARM64)    processor               with 630809K/3071K    bytes of memory.
                    # cisco C6807-XL          (M8572)    processor (revision )   with 1785856K/262144K bytes of memory.
                    # PowerPC CPU at 800Mhz, revision number 0x2151
                    # Note: 2900 series routers dont report proc due to variations so all are "various"                                        
                    #$Processor = ($Line.Split(" ")[0]) -replace '[\()\\]+'

                    If ($Line -like "*CPU at*"){
                        $Processor = $Line.Split(" ")[0] 
                    }Else{
                        $Processor = ($Line.Split("(")[1]).Split(")")[0] 
                    }
                    If (($Line -like "*cisco*") -and ($Null -eq $Obj.ModelNum)){
                        $ModelNum = $Line.Split(" ")[1].Trim()   
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "Processor" -Value $Processor -force   #----------------------[ Processor ]---------------   
                    $Obj | Add-Member -MemberType NoteProperty -Name "ModelNum" -Value $ModelNum -force     #----------------------[ Model Number ]---------------   
                }

                #--[ Get Model Number ]--
                If (($Line -like "*model number*") -and ($Null -eq $Obj.ModelNum)){
                    $ModelNum = $Line.Split(" ")[1].Trim()   
                    $Obj | Add-Member -MemberType NoteProperty -Name "ModelNum" -Value $ModelNum -force     #----------------------[ Model Number ]---------------   
                }
            
                #--[ Calculate RAM ]--------------------------------              
                If ($Line -like "*k bytes of physical memory*"){                   
                    If ($Line -like "*C4510*"){
                        [string]$Ram1 = ($Line.Split(" ")[7]) -replace "k"  
                    }ElseIf (($Line.Split(" ")[0]) -eq "cisco"){
                        [string]$Ram1 = ($Line.Split(" ")[5]) -replace "k"   
                    }Else{
                        [string]$Ram1 = ($Line.Split(" ")[0]) -replace "k"   
                    }
                    [Int]$Ram = (([Int]($Ram1.split("/")[0]))+(([Int]($Ram1.split("/")[1])))) / 1024 #/ 1024
                    $Obj | Add-Member -MemberType NoteProperty -Name "RAM" -Value $Ram -force 
                }
                If (($Line -like "*k bytes of memory*") -and ($Line -notlike "*C9*") -and ($Line -notlike "*ISR4331*")){ 
                    If ($Line -like "*CISCO29*"){
                        # If ($Line -like "*AIR-CAP*"){
                        $Ram1 = ($Line.Split(" ")[5]) -replace "k"   
                        #--[ Compensate for no 2900 series CPU data ]----------------
                        $Processor = "Various"
                        $Obj | Add-Member -MemberType NoteProperty -Name "Processor" -Value $Processor -force                     
                    }Else{
                        $Ram1 = ($Line.Split(" ")[7]) -replace "k"   
                    }
                    [Int]$Ram = [math]::Ceiling(([Int]($Ram1.split("/")[0])+([Int]($Ram1.split("/")[1])))/1000)                
                    $Obj | Add-Member -MemberType NoteProperty -Name "RAM" -Value $Ram -force 
                }

                #--[ Last Reload ]-----------------------------------------------------------------        
                If (($Line -like "Last reset from*") -or ($Line -like "Last reload reason*")){
                    If ($Line -like "*from*"){    
                        $LastReload = $Line.Split(" ")[3].Trim()
                    }ElseIf (($Line.Split(":")[1]) -ne " "){
                        $LastReload = $Line.Split(":")[1].Trim()
                    }   
                    $Obj | Add-Member -MemberType NoteProperty -Name "LastReload" -Value $LastReload -force
                }

                #--[ OS Version ]---------------------------------------------------------------------
                If ($Line -like "Cisco IOS Software*"){
                    If ($Line -like "*Catalyst*"){
                        If ($DeviceType -eq "Switch"){
                            $FWVersion = ($Line.split(",")[2]).Split(" ")[2].Trim()
                            $FWFamily = $Line.split(",")[0] #+", "+$Line.split(",")[1]
                            $FWBase = $Line.split(",")[1].trim()
                        }Else{
                            $FWVersion = ($Line.split(",")[3]).Split(" ")[2].Trim()
                            $FWFamily = $Line.split(",")[0]+", "+$Line.split(",")[1]
                            $FWBase = $Line.split(",")[2].trim()
                        }                
                        $FWRelease = $Line.subString($Line.length-4).trimend(")")
                    }Else{
                        $FWFamily = $Line.split(",")[0]
                        $FWVersion = $Line.split(",")[2].Split(" ")[2].trim()
                        $FWBase = $Line.split(",")[1].trim()
                        $FWRelease = (($Line.split(",")[3]).substring(19)).trim().trimend(")")
                    }

                    $Obj | Add-Member -MemberType NoteProperty -Name "FWVersion" -Value $FWVersion -Force
                    $Obj | Add-Member -MemberType NoteProperty -Name "FWRelease" -Value $FWRelease -Force
                    $Obj | Add-Member -MemberType NoteProperty -Name "FWFamily" -Value $FWFamily -Force
                    $Obj | Add-Member -MemberType NoteProperty -Name "FWBase" -Value $FWBase -Force
                    $Color++
                } #>

                #--[ Hostname ]----------------------------------------------------------------------
                If (!($Obj.Hostname)){
                    $Obj | Add-Member -MemberType NoteProperty -Name "Hostname" -Value $Script:Hostname -force
                }

                #--[ EOS, EOL, LDOS Dates & Upgrade Criticality ]------------------------------------       
                If ($MfgDate -ne "unknown"){ 
                    [String]$MfgEOLDate = (GetEOLDate $ModelNum).Split(",")[0]
                    $Obj | Add-Member -MemberType NoteProperty -Name "EOLDate" -Value $MfgEOLDate -force
                    [String]$MfgEOSDate = (GetEOLDate $ModelNum).Split(",")[1]
                    $Obj | Add-Member -MemberType NoteProperty -Name "EOSDate" -Value $MfgEOSDate -force
                    [String]$MfgLDOSDate = (GetEOLDate $ModelNum).Split(",")[2]
                    $Obj | Add-Member -MemberType NoteProperty -Name "LDOSDate" -Value $MfgLDOSDate -force 
                    $ThisYear = (Get-Date).ToString('yyyy')
                    $Age = [int]$ThisYear-[int]$MfgDate
                    $LDOSYear = $MfgLDOSDate.Split("/")[2]
                    [string]$UpgradePriority = "Low" 
                    If(($MfgLDOSDate.Split("/")[1]) -eq "A"){
                        [string]$UpgradePriority = "Low" 
                        [String]$ldosyear = "N/A"
                        [String]$ldosdiff = "N/A"  
                    }Else{
                        [int]$LDOSDiff = $LDOSYear-$ThisYear
                        If($LDOSDiff -le $PriorityCritical){     
                            [string]$UpgradePriority = "Critical"
                        }elseIf(($LDOSDiff -le $PriorityHigh) -and ($LDOSDiff -gt $PriorityCritical)){
                            [string]$UpgradePriority = "High"        
                        }elseIf(($LDOSDiff -le $PriorityMedium) -and ($LDOSDiff -gt $PriorityHigh)){
                            [string]$UpgradePriority = "Medium"        
                        }else{
                            [string]$UpgradePriority = "Low"        
                        }
                    }
                }#>
                    
                #--[ 10/100 Switches should be considered obsolete ]----------------------------------
                If($PortSpd -eq "10/100"){  
                    $UpgradePriority = "Critical"
                }

                $Obj | Add-Member -MemberType NoteProperty -Name "Age" -Value $Age -force 
                $Obj | Add-Member -MemberType NoteProperty -Name "UpgradePriority" -Value $UpgradePriority -force
                If ($DeviceType -eq "WAP"){
                    $Obj | Add-Member -MemberType NoteProperty -Name "SwitchNum" -Value "01" -force
                }

                #--[ Port Counts ]----------------------------------------------------------
                If ($Line -like "*Gigabit Ethernet interfaces*"){
                    $GigCount = $GigCount+$Line.Split(" ")[0]
                }
                If ($Line -like "*FastEthernet interfaces*"){
                    $FeCount = $FeCount+$Line.Split(" ")[0]            
                }
                If ($Line -like "*Ten Gigabit Ethernet interfaces*"){
                    $TenGCount = $TenGCount+$Line.Split(" ")[0] 
                }

                $Obj | Add-Member -MemberType NoteProperty -Name "GigPorts" -Value $GigCount -force      #-------------[ Total Gig Port Count ]-----------------
                $Obj | Add-Member -MemberType NoteProperty -Name "FastEPorts" -Value $FeCount -force     #-------------[ Total FastE Port Count ]-----------------
                $Obj | Add-Member -MemberType NoteProperty -Name "TenGigPorts" -Value $TenGCount -force  #-------------[ Total Ten Gig Port Count ]-----------------
            
                #--[ Cycle through multiple Switches ]------------------------------------------------
                If ($DeviceType -eq "Switch"){   #--[ Only process if working on switches ]--
                    $Obj | Add-Member -MemberType NoteProperty -Name "SwitchNum" -Value "01" -force
                    $Obj | Add-Member -MemberType NoteProperty -Name "Age" -Value $Age -force
                    $Obj | Add-Member -MemberType NoteProperty -Name "MfgDate" -Value $MfgDate -force
                    $Obj | Add-Member -MemberType NoteProperty -Name "UpgradePriority" -Value $UpgradePriority -force
                    Switch ($Line) {
                        "Switch 02" {  
                            $DeviceCount++          
                            # $Key = $Line.Split(" ")[1]
                            $Color++
                            $SwitchNum = $Line
                            $Obj | Add-Member -MemberType NoteProperty -Name "SwitchNum" -Value ($SwitchNum.Split(" ")[1]) -force
                            #--[ Serialize and Deserialize data using BinaryFormatter ]--
                            $ms = New-Object System.IO.MemoryStream
                            $bf = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
                            $bf.Serialize($ms, $Obj)
                            $ms.Position = 0
                            #--[ Write deep copied data ]--
                            $Obj02 = $bf.Deserialize($ms)
                            $ms.Close()
                        }
                        "Switch 03" {
                            $DeviceCount++          
                            #$Key = $Line.Split(" ")[1]
                            $Color++
                            $SwitchNum = $Line
                            $Obj | Add-Member -MemberType NoteProperty -Name "SwitchNum" -Value ($SwitchNum.Split(" ")[1]) -force
                            #--[ Serialize and Deserialize data using BinaryFormatter ]--
                            $ms = New-Object System.IO.MemoryStream
                            $bf = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
                            $bf.Serialize($ms, $Obj)
                            $ms.Position = 0
                            #--[ Write deep copied data ]--
                            $Obj03 = $bf.Deserialize($ms)
                            $ms.Close()
                        }
                        "Switch 04" {
                            $DeviceCount++          
                            # $Key = $Line.Split(" ")[1]
                            $Color++
                            $SwitchNum = $Line
                            $Obj | Add-Member -MemberType NoteProperty -Name "SwitchNum" -Value ($SwitchNum.Split(" ")[1]) -force
                            #--[ Serialize and Deserialize data using BinaryFormatter ]--
                            $ms = New-Object System.IO.MemoryStream
                            $bf = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
                            $bf.Serialize($ms, $Obj)
                            $ms.Position = 0
                            #--[ Write deep copied data ]--
                            $Obj04 = $bf.Deserialize($ms)
                            $ms.Close()
                        }
                        "Switch 05" {
                            $DeviceCount++          
                            #$Key = $Line.Split(" ")[1]
                            $Color++
                            $SwitchNum = $Line
                            $Obj | Add-Member -MemberType NoteProperty -Name "SwitchNum" -Value ($SwitchNum.Split(" ")[1]) -force
                            #--[ Serialize and Deserialize data using BinaryFormatter ]--
                            $ms = New-Object System.IO.MemoryStream
                            $bf = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
                            $bf.Serialize($ms, $Obj)
                            $ms.Position = 0
                            #--[ Write deep copied data ]--
                            $Obj05 = $bf.Deserialize($ms)
                            $ms.Close()
                        }
                        "Switch 06" {
                            $DeviceCount++          
                            # $Key = $Line.Split(" ")[1]
                            $Color++
                            $SwitchNum = $Line
                            $Obj | Add-Member -MemberType NoteProperty -Name "SwitchNum" -Value ($SwitchNum.Split(" ")[1]) -force
                            #--[ Serialize and Deserialize data using BinaryFormatter ]--
                            $ms = New-Object System.IO.MemoryStream
                            $bf = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
                            $bf.Serialize($ms, $Obj)
                            $ms.Position = 0
                            #--[ Write deep copied data ]--
                            $Obj06 = $bf.Deserialize($ms)
                            $ms.Close()
                        }
                        "Switch 07" {
                            $DeviceCount++          
                            # $Key = $Line.Split(" ")[1]
                            $Color++
                            $SwitchNum = $Line
                            $Obj | Add-Member -MemberType NoteProperty -Name "SwitchNum" -Value ($SwitchNum.Split(" ")[1]) -force
                            #--[ Serialize and Deserialize data using BinaryFormatter ]--
                            $ms = New-Object System.IO.MemoryStream
                            $bf = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
                            $bf.Serialize($ms, $Obj)
                            $ms.Position = 0
                            #--[ Write deep copied data ]--
                            $Obj07 = $bf.Deserialize($ms)
                            $ms.Close()
                        }
                        "Switch 08" {
                            $DeviceCount++          
                            #$Key = $Line.Split(" ")[1]
                            $Color++
                            $SwitchNum = $Line
                            $Obj | Add-Member -MemberType NoteProperty -Name "SwitchNum" -Value ($SwitchNum.Split(" ")[1]) -force
                            #--[ Serialize and Deserialize data using BinaryFormatter ]--
                            $ms = New-Object System.IO.MemoryStream
                            $bf = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
                            $bf.Serialize($ms, $Obj)
                            $ms.Position = 0
                            #--[ Write deep copied data ]--
                            $Obj08 = $bf.Deserialize($ms)
                            $ms.Close()
                        }
                    }                
                }#>
                $LineCnt++
            } #--[ End of Result Variable line parsing ]--

            #--[ Make sure there are enough spreadsheet rows for data in case a new switch was added to a stack ]--
            If ($DeviceCount -gt 1){  
                StatusMsg "Verifying and adding spreadsheet rows if needed..." "Magenta" $Debug
                $DeviceRow = $Row
                $AddRows = 0
                $DeviceRowCnt = 0
                While ($DeviceRowCnt -lt $DeviceCount ){
                    $RowIP = $WorkSheet.Cells.Item($DeviceRow,1).Text 
                    If ($RowIP -eq $IP){
                        $DeviceRow++
                        $DeviceRowCnt++
                    }Else{
                        $AddRows = $DeviceCount-$DeviceRowCnt
                        break;break
                    }
                }  
                StatusMsg "Adding $AddRows" "Magenta" $Debug    
                $xlShiftDown = -4121
                While ($AddRows -gt 0){
                    $objRange = $WorkSheet.Range("A$testRow").EntireRow
                    [void]$objRange.Insert($xlShiftDown)
                    $AddRows--
                }
            }

            $TotalPorts = $GigCount+$FeCount+$TenGCount
            $Obj | Add-Member -MemberType NoteProperty -Name "TotalPorts" -Value $TotalPorts -force    #-------------[ Total Port Count ]-----------------
            If ($Obj.PortCount -eq ""){
                $Obj | Add-Member -MemberType NoteProperty -Name "PortCount" -Value $TotalPorts -force  #-------------[ Port Count ]-----------------      
            }

            #--[ Final tweaks prior to writing data to spreadsheet ]--
            If (($Obj.Connection -eq "OK") -and ($Null -eq $Obj.UpgradePriority)){
                $Obj | Add-Member -MemberType NoteProperty -Name "Connection" -Value "ISSUES (Ping OK)" -force
            }

            StatusMsg "Storing processed results" "Magenta" $Debug

            $Loop = 1
            While($Loop -le ($DeviceCount)){
                Switch ($Loop){
                    "1" {$Obj = $Obj01;$Color = 2}
                    "2" {$Obj = $Obj02;$Color = 3}
                    "3" {$Obj = $Obj03;$Color = 4}
                    "4" {$Obj = $Obj04;$Color = 5}
                    "5" {$Obj = $Obj05;$Color = 6}
                    "6" {$Obj = $Obj06;$Color = 7}
                    "7" {$Obj = $Obj07;$Color = 8}
                    "8" {$Obj = $Obj08;$Color = 9}
                }          

                If ($Console){
                    # Write-host "---------------------------------------------------------------------------" -ForegroundColor yellow 
                    Write-host "Collected Data Results      :" -ForegroundColor DarkCyan
                    Write-Host "Spreadsheet Row             :"$Row -ForegroundColor $color
                    Write-Host "Connection                  :"$Obj.Connection -ForegroundColor $color
                    Write-Host "LAN IP Address              :"$Obj.IPAddress -ForegroundColor $color
                    Write-Host "Hostname                    :"$Obj.Hostname -ForegroundColor $color
                    Write-Host "Base Ethernet MAC Address   :"$Obj.BaseMAC -ForegroundColor $color
                    Write-Host "Facility                    :"$Obj.Facility -ForegroundColor $color
                    Write-Host "Address                     :"$Obj.Address -ForegroundColor $color
                    Write-Host "IDF                         :"$Obj.IDF -ForegroundColor $color
                    Write-Host "Description                 :"$Obj.Description -ForegroundColor $color
                    Write-Host "Asset Tag                   :"$Obj.AssetTag -ForegroundColor $color
                    Write-Host "Device Type                 :"$Obj.DeviceType -ForegroundColor $color
                    Write-Host "Serial Number               :"$Obj.SerialNum -ForegroundColor $color
                    Write-Host "Model Number                :"$Obj.ModelNum -ForegroundColor $color            
                    Write-Host "Year of Manufacture         :"$Obj.MfgDate -ForegroundColor $color
                    Write-Host "Device age (Years)          :"$Obj.Age -ForegroundColor $color
                    Write-Host "End of Life Year (EOL)      :"$Obj.EOLDate -ForegroundColor $color
                    Write-Host "End of Sale Year (EOS)      :"$Obj.EOSDate -ForegroundColor $color            
                    Write-Host "End of Support Year (LDOS)  :"$Obj.LDOSDate -ForegroundColor $color
                    Write-Host "Upgrade Priority            :"$Obj.UpgradePriority -ForegroundColor $color
                    Write-Host "Processor                   :"$Obj.Processor -ForegroundColor $color
                    Write-Host "RAM (MB)                    :"$Obj.Ram -ForegroundColor $color  
                    Write-Host "Last reload reason          :"$Obj.LastReload -ForegroundColor $color
                    Write-Host "Motherboard Serial Number   :"$Obj.MBSerialNum -ForegroundColor $color
                    Write-Host "Firmware Version            :"$Obj.FWVersion -ForegroundColor $color 
                    Write-Host "Firmware Release            :"$Obj.FWRelease -ForegroundColor $color        
                    Write-Host "Firmware Family             :"$Obj.FWFamily -ForegroundColor $color
                    Write-Host "Firmware Base               :"$Obj.FWBase -ForegroundColor $color    
                    Write-Host "Uptime Years                :"$Years -ForegroundColor $color
                    Write-Host "Uptime Weeks                :"$Weeks -ForegroundColor $color
                    Write-Host "Uptime Days                 :"$Days -ForegroundColor $color
                    #Write-Host "Uptime Hours                :"$Hours -ForegroundColor $color
                    #Write-Host "Uptime Minutes              :"$Minutes -ForegroundColor $color
                    #Write-Host "Uptime Days Total           :"$TotalUpDays -ForegroundColor $color
                    Write-Host "Port Speed                  :"$Obj.PortSpeed -ForegroundColor $color
                    Write-Host "Total Gig Ports             :"$Obj.GigPorts -ForegroundColor $color
                    Write-Host "Total Fast Eth Ports        :"$Obj.FastEPorts -ForegroundColor $color
                    Write-Host "Total 10 Gig Ports          :"$Obj.TenGigPorts -ForegroundColor $color
                    Write-Host "Total Port Count            :"$Obj.TotalPorts -ForegroundColor $color
                    Switch ($DeviceType){
                        "Router" {
                            Write-Host "WAN Circuit ID              :"$Obj.CID -ForegroundColor $color
                            Write-Host "WAN IP                      :"$Obj.WanIP -ForegroundColor $color
                            Write-Host "PRI Port count              :"$PRI -ForegroundColor $color
                            Write-Host "ATM Interface count         :"$ATM -ForegroundColor $color
                            Write-Host "FXO Interface count         :"$FXO -ForegroundColor $color
                            Write-Host "FXS Interface count         :"$FXS -ForegroundColor $color
                            Write-Host "VPN Module count            :"$VPN -ForegroundColor $color
                        }
                        "WAP" {
                            Write-Host "Host Switch                 :"$Obj.HostSwitch -ForegroundColor $color
                            Write-Host "Host Switch Port #          :"$Obj.HostSwPort -ForegroundColor $color                
                        }
                        Default {
                            #Write-Host "Port count                  :"$Obj.PortCount -ForegroundColor $color
                            Write-Host "Switch Sequence in Stack    :"$Obj.SwitchNum -ForegroundColor $color              
                        }
                    }
                }

                If ($NoAccess){
                    Write2Excel $WorkSheet $Row 3 "-- ACCESS DENIED --"                   #--[ Column 03 (C) ]--  
                }Else{
                    If ($EnableExcel){   #--[ Calls Function Write2Excel ($WorkSheet,$Row,$Col,$Data,$Format,$Debug) ]--
                        $Index = $WorkSheet.Cells($Row,1).Interior.ColorIndex             #--[ Determine existing cell color index ]--
                        $Excel.ActiveSheet.UsedRange.Rows.Item($Row).Interior.ColorIndex = 20   #--[ Set row color to pale blur ]-- 
                        Write2Excel $WorkSheet $Row 1 $Obj.IPAddress "existing"           #--[ Column 01 (A) ]--
                        Write2Excel $WorkSheet $Row 2 $Obj.Hostname "existing"            #--[ Column 02 (B) ]-- 
                        Write2Excel $WorkSheet $Row 3 $Obj.Connection                     #--[ Column 03 (C) ]--  
                        Write2Excel $WorkSheet $Row 4 $Obj.BaseMAC "mac"                  #--[ Column 04 (D) ]--  
                        Write2Excel $WorkSheet $Row 5 $Obj.Facility                       #--[ Column 05 (E) ]--  
                        Write2Excel $WorkSheet $Row 6 $Obj.Address                        #--[ Column 06 (F) ]--   
                        Write2Excel $WorkSheet $Row 7 $Obj.IDF                            #--[ Column 07 (G) ]--   
                        Write2Excel $WorkSheet $Row 8 $Obj.Description                    #--[ Column 08 (H) ]--  
                        #--[ Jump over Asset Tag column ]--  
                        #Write2Excel $WorkSheet $Row 9 $Obj.AssetTag                      #--[ Column 09 (I) ]--
                        Write2Excel $WorkSheet $Row 10 $Obj.DeviceType                    #--[ Column 10 (J) ]--
                        Write2Excel $WorkSheet $Row 11 $Obj.SerialNum                     #--[ Column 11 (K) ]--            
                        Write2Excel $WorkSheet $Row 12 $Obj.ModelNum "existing"           #--[ Column 12 (L) ]-- 
                        Write2Excel $WorkSheet $Row 13 $Obj.MfgDate                       #--[ Column 13 (M) ]--
                        Write2Excel $WorkSheet $Row 14 $Obj.Age                           #--[ Column 14 (N) ]--
                        Write2Excel $WorkSheet $Row 15 $Obj.EOLDate "date"                #--[ Column 15 (O) ]--  
                        Write2Excel $WorkSheet $Row 16 $Obj.EOSDate "date"                #--[ Column 16 (P) ]-
                        Write2Excel $WorkSheet $Row 17 $Obj.LDOSDate "date"               #--[ Column 17 (Q) ]-- 
                        Write2Excel $WorkSheet $Row 18 $Obj.UpgradePriority               #--[ Column 18 (R) ]--  
                        Write2Excel $WorkSheet $Row 19 $Obj.Processor                     #--[ Column 19 (S) ]--
                        Write2Excel $WorkSheet $Row 20 $Obj.RAM                           #--[ Column 20 (T) ]--
                        Write2Excel $WorkSheet $Row 21 $Obj.MBSerialNum                   #--[ Column 21 (U) ]--   
                        Write2Excel $WorkSheet $Row 22 $Obj.FWVersion                     #--[ Column 22 (V) ]--
                        Write2Excel $WorkSheet $Row 23 $Obj.FWRelease                     #--[ Column 23 (W) ]--
                        Write2Excel $WorkSheet $Row 24 $Obj.FWFamily                      #--[ Column 24 (X) ]--
                        Write2Excel $WorkSheet $Row 25 $Obj.FWBase                        #--[ Column 25 (Y) ]--
                        Write2Excel $WorkSheet $Row 26 $Obj.LastReload                    #--[ Column 26 (Z) ]--
                        Write2Excel $WorkSheet $Row 27 $Obj.TotalDaysUp                   #--[ Column 27 (AA) ]--
                        Write2Excel $WorkSheet $Row 28 $Obj.PortSpeed                     #--[ Column 28 (AB) ]--
                        #--[ Unique Items]
                        Switch ($DeviceType){
                            "Router" { 
                                Write2Excel $WorkSheet $Row 29 $Obj.CID                       #--[ Column 29 (AC ]--
                                Write2Excel $WorkSheet $Row 30 $Obj.WanIP                     #--[ Column 30 (AD) ]--
                                Write2Excel $WorkSheet $Row 31 $Obj.RtrPorts                  #--[ Column 31 (AE) ]--
                                Write2Excel $WorkSheet $Row 31 $Obj.RtrPorts                  #--[ Column 31 (AF) ]--
                                $Range = $WorkSheet.Range(("A$Row"),("AF$Row")) 
                            }
                            "WAP" {
                                Write2Excel $WorkSheet $Row 29 $Obj.HostSwitch                #--[ Column 29 (AC) ]--
                                Write2Excel $WorkSheet $Row 30 $Obj.HostSwPort                #--[ Column 30 (AD) ]-- 
                                Write2Excel $WorkSheet $Row 30 $Obj.HostSwPort                #--[ Column 30 (AE) ]-- 
                                Write2Excel $WorkSheet $Row 31 $Obj.Today                     #--[ Column 30 (AF) ]--
                                $Range = $WorkSheet.Range(("A$Row"),("AF$Row")) 
                            }
                            Default {
                                Write2Excel $WorkSheet $Row 29 $Obj.TotalPorts                #--[ Column 28 (AC) ]--
                                Write2Excel $WorkSheet $Row 30 $Obj.SwitchNum                 #--[ Column 29 (AD) ]--
                                Write2Excel $WorkSheet $Row 31 $Obj.Today                     #--[ Column 30 (AE) ]--
                                $Range = $WorkSheet.Range(("A$Row"),("AE$Row")) 
                            }
                        }  
                    }

                    $Range.HorizontalAlignment = -4131
                    1..4 | ForEach-Object {
                        $Range.Borders.Item($_).LineStyle = 1
                        $Range.Borders.Item($_).Weight = 2
                    }
                    $Resize = $WorkSheet.UsedRange
                    [Void]$Resize.EntireColumn.AutoFit() 

                    #--[ Set row background according to upgrade priority ]--
                    Switch ($Obj.UpgradePriority){
                        "Low" {
                            $Color = 35  #--[ Light Green]--        
                        }
                        "Medium" {
                            $Color = 19  #--[ Light Yellow]--
                        }
                        "High" {
                            $Color = 40  #--[ Light Orange]--
                        } 
                        "Critical" {
                            $Color = 3  #--[ Red ]--
                        }
                        Default {
                            $Color = 15  #--[ Gray ]--
                            If (($NewData -eq "OK") -and ($WorkSheet.Cells.Item($Row,$Col).Text )){
                                $WorkSheet.UsedRange.Rows.Item($Row)
                            }
                        }  
                    }  

                    $Index = $WorkSheet.Cells($Row,1).Interior.ColorIndex  #--[ Determine existing cell color index ]--
                    If ($Index -ne $Color){   
                        $WorkSheet.UsedRange.Rows.Item($Row).Interior.ColorIndex = $Color
                    }
                    If ($Color -eq 3){
                        $WorkSheet.UsedRange.Rows.Item($Row).Font.ColorIndex = 6  #--[ Yellow text if background is red ]--
                    }
                }  

#----------[ SQLlite Experimental Area]-------------------            

            
                    <#          
                    If ($EnableSQLite){
                    Import-Module PSSQLite
                    $DataSource = "b:\tracker\tracker.sqlite"

                    SELECT
                    employeeid,
                    firstname,
                    lastname,
                    state,
                    city,
                    PostalCode
                FROM
                    employees
                WHERE
                    employeeid = 4;

                    # initial import
                    $query = "INSERT INTO SWITCH (MACAddress, IPAddress, HostName, SerialNum, ModelNum, Facility, Floor, Description, AssetTag, MBSerial, SwitchNum, MfgDate, DaysUp, Age, EOL, EOS, UpgradePriorityority, PortCount, PortSpeed, Processor, RAM, LastReload, FirmwareVer, FirmwareRel, FirmwareFamily, FirmwareBase) VALUES (@MACAddress, @IPAddress, @HostName, @SerialNum, @ModelNum, @Facility, @Floor, @Description, @AssetTag, @MBSerial, @SwitchNum, @MfgDate, @DaysUp, @Age, @EOL, @EOS, @UpgradePriorityority, @PortCount, @PortSpeed, @Processor, @RAM, @LastReload, @FirmwareVer, @FirmwareRel, @FirmwareFamily, @FirmwareBase)"

                    Invoke-SqliteQuery -DataSource $DataSource -Query $query -SqlParameters @{

                    #
                  #  $ImportObj = [pscustomobject]@{
                    MACAddress = $Obj.BaseMAC
                    IPAddress = $Obj.IPAddress
                    HostName = $Obj.Hostname
                    SerialNum = $Obj.SerialNum
                    ModelNum = $Obj.ModelNum 
                    Facility = $Obj.Facility
                    Floor = $Obj.Floor 
                    Description = $Obj.Description
                    AssetTag = ""
                    MBSerial = $Obj.MBSerialNum
                    SwitchNum = $Obj.SwitchNum 
                    MfgDate = $Obj.MfgDate 
                    DaysuP = $oBJ.DaysUp
                    Age = $Obj.Age
                    EOL = $Obj.Warranty 
                    EOS = $Obj.Warranty 
                    UpgradePriorityority = $Obj.UpgradePriority
                    PortCount = $Obj.PortCount 
                    PortSpeed = $Obj.PortSpeed
                    Processor = $Obj.Processor 
                    RAM = $Obj.RAM
                    LastReload = $Obj.LastReload  
                    FirmwareVer = $Obj.FWVersion 
                    FirmwareRel = $Obj.FWRelease 
                    FirmwareFamily = $Obj.FWFamily
                    FirmwareBase = $Obj.FWBase
                 } 
                  #>               
                #   Invoke-SQLiteBulkCopy -DataTable $ImportObj -DataSource $DataSource -Table SWITCH -NotifyAfter 1000 -verbose
                
                
                #       Invoke-SqliteQuery -DataSource $DataSource -Query "SELECT * FROM SWITCH"

                #}
#----------[ SQLlite Experimental Area]-------------------               
              
                
                $Loop++
                $Row++
            }
        }Else{
            StatusMsg "--- No Connection ---" "Red" $Debug
            If ($EnableExcel){
                $WorkSheet.Cells.Item($Row, 1) = $IP 
                $WorkSheet.Cells.Item($Row, 3) = "No Connection" 
                $WorkSheet.UsedRange.Rows.Item($Row).Interior.ColorIndex = 15  #--[ Background set to grey if connection fails ]--
                $Row++
            }
        }
        StatusMsg "Clearing run variables." "magenta" $Debug
        Remove-Variable Obj -ErrorAction "silentlycontinue"
        Remove-Variable Obj01 -ErrorAction "silentlycontinue"
        Remove-Variable Obj02 -ErrorAction "silentlycontinue"
        Remove-Variable Obj03 -ErrorAction "silentlycontinue"
        Remove-Variable Obj04 -ErrorAction "silentlycontinue"
        Remove-Variable Obj05 -ErrorAction "silentlycontinue"
        Remove-Variable Obj06 -ErrorAction "silentlycontinue"
        Remove-Variable Obj07 -ErrorAction "silentlycontinue"
        Remove-Variable Obj08 -ErrorAction "silentlycontinue"        
    }
    StatusMsg "End of Item $IP" "magenta" $Debug
}

#--[ Cleanup ]--
$Excel.DisplayAlerts = $False
Write-host ""
Try{ 
    If ($Script:NewSpreadsheet -And (Test-Path -Path $ExcelWorkingCopy)){
        StatusMsg '!!! Existing working spreadsheet is NOT being overwritten !!!' "Yellow" $Debug
        StatusMsg 'Saving as "NewSpreadsheet.xlsx" ...' "Green" $Debug
        $Excel.ActiveWorkbook.SaveAs($PSScriptRoot+"\NewSpreadsheet.xlsx")
    }ElseIf(!(Test-Path -Path $ExcelWorkingCopy)){
        StatusMsg "Saving as a new working spreadsheet... " "Green" $Debug
        $Excel.ActiveWorkbook.SaveAs($ExcelWorkingCopy)
    }Else{  
        StatusMsg "Saving working spreadsheet... " "Green" $Debug       
        $Excel.ActiveWorkbook.Save() 
    }
    $Excel.quit()                                               #--[ Quit Excel ]--
    [Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) #--[ Release the COM object ]--
}Catch{
    StatusMsg "Save Failed..." "Red" $Debug
    Write-Host "`n`n  NOTICE !!! ---  Spreadsheet has NOT been saved.  Please manually save at this time ---`n`n" -ForegroundColor Yellow
    Write-Host $_.Exception.Message -ForegroundColor Red
}

Write-Host "`n--- COMPLETED ---" -ForegroundColor red
 
