# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Custom :: Application Information
# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

<#

               Template Version: 1.0.0
            
            Application Version: 

                      File Name: 

           Custom Areas Updated: 

                       Synopsis: 

                    Description: 

                 Known Issue(s): 

             Additional Note(s): 

                   Reference(s): 

                Original Author: 

                 Author Contact: 

                    Author Site: https://

              Author Repository: https://

                     Change Log: V1.00.00 - 00/00/0000 - Initial Release

#>

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Required :: Variables (Customisation Required)
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    [Boolean] $LaunchReportOnCompletion = $True

    [String]  $ErrorActionPreference    = "SilentlyContinue"

    [Object]  $CurrentDate              = Get-Date
    
    [String]  $CurrentDate              = $CurrentDate.ToString('dd-MM-yyyy')
    
    [String]  $Author                   = "Author Name"

    [String]  $ConsoleTitle             = "PowerShell Console Window Title"

    [String]  $CompanyName              = "Company Name"
    
    [String]  $ReportVersion            = "1.0.0"
    
    [String]  $ReportTitle              = "HTML Report Title Name"

    [String]  $ReportLocation           = "$env:USERPROFILE\Desktop\"

    [String]  $ReportFileName           = "$ReportTitle - $CurrentDate.html"

    [String]  $CompanyLogoBorder        = "none" # Replace with "1px dashed black" to see the logo border
    
    [String]  $CompanyLogoWidth         = "127px"

    [String]  $CompanyLogoHeight        = "50px"

    [String]  $CompanyLogoTop           = "28px"

    [String]  $CompanyLogoRight         = "37px"

    [String]  $CompanyLogoBase64String  = "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"

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Required :: Variables, Calculations, Expressions & Initial Configuration
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Try { $Host.UI.RawUI.WindowTitle = $ConsoleTitle } Catch {} Finally { $Error.Clear(); Clear-Host; }

    Write-Progress -Activity 'Generating Report' -Status 'Loading Script Variables' -PercentComplete 0; Sleep -Milliseconds 750

    $ThisPC      = $env:COMPUTERNAME

    $CurrentUser = $env:USERNAME

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Custom :: Variables, Calculations, Expressions & Initial Configuration
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

# This custom area allows you to keep your own variables in one place (Globally) allowing them to be used from various locations below.
# The current variables in this section are here for the purpose of demonstrating and are used by the existing template queries below.

    $Custom_TimestampAtBoot    = Get-WmiObject Win32_PerfRawData_PerfOS_System | Select-Object -ExpandProperty systemuptime
    $Custom_CurrentTimestamp   = Get-WmiObject Win32_PerfRawData_PerfOS_System | Select-Object -ExpandProperty Timestamp_Object
    $Custom_Frequency          = Get-WmiObject Win32_PerfRawData_PerfOS_System | Select-Object -ExpandProperty Frequency_Object

    $Custom_UptimeInSec        = (($Custom_CurrentTimestamp - $Custom_TimestampAtBoot)/$Custom_Frequency)
    $Custom_Time               = (Get-Date) - (New-TimeSpan -seconds $Custom_UptimeInSec) 
    $Custom_Date               = (Get-Date) - (New-TimeSpan -Day 1)

    $Custom_LogicServiceColour = @{Stopped    = ' style="background-color: #F8D7DA; color: #721C24;">Stopped<';Running = ' style="background-color: #D4EDDA; color: #15572a;">Running<';}
    $Custom_LogicEventColour   = @{Error      = ' style="background-color: #F8D7DA; color: #721C24;">Error<'  ;Warning = ' style="background-color: #FFF3CD; color: #856404;">Warning<';}

    $Custom_FreeSpace          = @{Expression = {[int]($_.Freespace/1GB)};         Name = 'Free Space (GB)'}
    $Custom_Size               = @{Expression = {[int]($_.Size/1GB)};              Name = 'Size (GB)'      }
    $Custom_PercentFree        = @{Expression = {[int]($_.Freespace*100/$_.Size)}; Name = 'Free (%)'       }

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Required :: Functions
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Function Update-ProgressBar {

        # This function allows you to update the Progress Bar within the PowerShell Console.

        # Example Usage: Update-ProgressBar -Activity "Generating Report" -Status "System Tab: BIOS" -PercentComplete 20;

        Param(
            [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
            [String] $Activity,
            [String] $Status,
            [Int]    $PercentComplete

        )

        Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete;
        
        Sleep -Milliseconds 250;

    }
    
    Function Add-HTMLTableAttribute {

        # This function allows you to add a custom attribute such as ID or Class to your auto-generated table(s).
        # It will however auto-generate <col> and <colgroup> tags, therefore there is a script block within...
        # the $Scripts variable section that will iterate through the DOM and remove all instances of <col> and <colgroup>...
        # When the DOM is loading.
        # The original Source will contain the elements but the Developer Tools (Live DOM) will remove them

        # Example Usage: $WMI = [WMI Query] | ConvertTo-HTML -Fragment | Out-String | Add-HTMLTableAttribute -AttributeName 'id' -Value 'tab-system-cpu';

        Param(
            [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
            [string] $HTML,
            [string] $AttributeName,
            [string] $Value

        )

        $XML = [xml]$HTML ; $Attribute = $XML.CreateAttribute($AttributeName) ;
        
        $Attribute.Value = $Value ; $XML.Table.Attributes.Append($Attribute) | Out-Null ;
        
        Return ($XML.OuterXML | Out-String) ;
    
    }

    Function Add-PreContentMessage {

        # This function allows you to add Pre-Content HTML directly above the Table output by displaying a banner with text and an image.
        # It allows you to change the Image ,Colour and Text of the banner using one of the 3 switches (Below)
        
        # - Error       : Red Banner, Red Text, traditional Error Icon
        # - Information : Blue Banner, Blue Text, traditional Information Icon
        # - Warning     : Amber Banner, Amber Text, traditional Warning Icon

        # Example Usage: $WMI = [WMI Query] | ConvertTo-HTML -Fragment -PreContent (Add-PreContentMessage -Type "Information" -Message "Displaying data for drive C: only.");

        Param(
            [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
            [String] $Type,
            [String] $Message

        )

        Switch ($Type) 
        {
            "Error"       { $PreContent = "<section class='pre-content-alert-error'><div class='pre-content-alert-image-error'></div><span class='pre-content-alert-text-error'>$Message</span></section>"                   }
            "Information" { $PreContent = "<section class='pre-content-alert-information'><div class='pre-content-alert-image-information'></div><span class='pre-content-alert-text-information'>$Message</span></section>" }
            "Warning"     { $PreContent = "<section class='pre-content-alert-warning'><div class='pre-content-alert-image-warning'></div><span class='pre-content-alert-text-warning'>$Message</span></section>"             }
        } 

        Return $PreContent | Out-String;
    
    }

# -------------------------------------
# Your Custom Queries --->
# -------------------------------------

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# System Tab :: Queries and Logic :: BIOS
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Update-ProgressBar -Activity "Generating Report" -Status "System Tab: BIOS" -PercentComplete 20;
    
    $Query_System_BIOS = Get-WmiObject -Class Win32_BIOS -Computername $ThisPC | Select SMBIOSBIOSVersion,Manufacturer,Name,SerialNumber,Version | ConvertTo-HTML -Fragment;

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# System Tab :: Queries and Logic :: Battery
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Update-ProgressBar -Activity "Generating Report" -Status "System Tab: Battery" -PercentComplete 20;
    
    $Query_System_Battery = Get-WmiObject -Class Win32_Battery -Computername $ThisPC | Select Caption, EstimatedChargeRemaining,EstimatedRunTime,Status | ConvertTo-HTML -Fragment;

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# System Tab :: Queries and Logic :: Operating System
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Update-ProgressBar -Activity "Generating Report" -Status "System Tab: Operating System" -PercentComplete 20;
    
    $Query_System_OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ThisPC | Select-Object -Property Organization,RegisteredUser,CSName,Caption,BuildNumber,ServicePackMajorVersion,Version, @{n='LastBootTime';e={$_.ConvertToDateTime($_.LastBootUpTime)}} | ConvertTo-HTML -Fragment;

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# System Tab :: Queries and Logic :: CPU Information
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Update-ProgressBar -Activity "Generating Report" -Status "System Tab: CPU" -PercentComplete 20;
    
    $Query_System_CPU = get-wmiobject -class win32_processor -ComputerName $ThisPC | Select-Object -Property Manufacturer,Name,Caption,LoadPercentage,Status,NumberOfCores,NumberOfLogicalProcessors | ConvertTo-HTML -Fragment | Out-String | Add-HTMLTableAttribute  -AttributeName 'id' -Value 'tab-system-cpu';

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# System Tab :: Queries and Logic :: Disk Information
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Update-ProgressBar -Activity "Generating Report" -Status "System Tab: Disk" -PercentComplete 20;

    $Query_System_Disk = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $ThisPC | Where {$_.DeviceID -eq 'C:'} | Select-Object -Property DeviceID, VolumeName, $Custom_Size, $Custom_FreeSpace, $Custom_PercentFree | ConvertTo-HTML -Fragment -PreContent (Add-PreContentMessage -Type "Information" -Message "Displaying data for drive C: only.");

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# System Tab :: Queries and Logic :: Event Viewer :: Application
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Update-ProgressBar -Activity "Generating Report" -Status "System Tab: Event Viewer (Application)" -PercentComplete 20;
    
    $Query_System_AppEvent = Get-EventLog -ComputerName $ThisPC -LogName Application -EntryType "Error","Warning"-after $Custom_Time | Select-Object -property EventID, EntryType, Source, TimeGenerated, Message | Sort-Object EntryType, TimeGenerated | ConvertTo-HTML -Fragment -PreContent (Add-PreContentMessage -Type "Information" -Message "Displaying Error &amp; Warning events only.");

    $Custom_LogicEventColour.Keys | ForEach {
    
        $Query_System_AppEvent = $Query_System_AppEvent -replace ">$_<",($Custom_LogicEventColour.$_)
        
    }

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# System Tab :: Queries and Logic :: Event Viewer :: System
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Update-ProgressBar -Activity "Generating Report" -Status "System Tab: Event Viewer (System)" -PercentComplete 20;
    
    $Query_System_SysEvent = Get-EventLog -ComputerName $ThisPC -LogName System -EntryType "Error","Warning" -After $Custom_Time | Select-Object -property EventID, EntryType, Source, TimeGenerated, Message | Sort-Object EntryType, TimeGenerated |  ConvertTo-HTML -Fragment -PreContent (Add-PreContentMessage -Type "Information" -Message "Displaying Error &amp; Warning events only.");

    $Custom_LogicEventColour.Keys | ForEach {
    
        $Query_System_SysEvent = $Query_System_SysEvent -replace ">$_<",($Custom_LogicEventColour.$_)
        
    }

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# System Tab :: Queries and Logic :: Hot Fix Information
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Update-ProgressBar -Activity "Generating Report" -Status "System Tab: Hot Fix" -PercentComplete 20;

    $Query_System_HotFix = Get-WmiObject -Class Win32_QuickFixEngineering -ComputerName $ThisPC | Select-Object -Property @{Name='Hot Fix ID'; Expression={([String]($_.HotFixID))}}, Caption, Description, @{Name='Install Date'; Expression={([DateTime]($_.InstalledOn)).ToLocalTime()}} | Sort HotFixID, InstalledOn | ConvertTo-HTML -Fragment;

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# System Tab :: Queries and Logic :: Installed Applications
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Update-ProgressBar -Activity "Generating Report" -Status "System Tab: Installed Applications" -PercentComplete 20;

    $Query_System_InstalledApps = Get-WmiObject -Class Win32_Product -ComputerName $ThisPC | Select-Object Name, Version, InstallDate, Vendor, HelpLink | Sort-Object Name | ConvertTo-HTML -Fragment;

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# System Tab :: Queries and Logic :: Services
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Update-ProgressBar -Activity "Generating Report" -Status "System Tab: Services" -PercentComplete 20;

    $Query_System_Services = Get-WmiObject -Class win32_service -ComputerName $ThisPC | Select-Object DisplayName, Name, StartMode, State | sort StartMode, State, DisplayName | ConvertTo-HTML -Fragment ;

    $Custom_LogicServiceColour.Keys | ForEach {
    
        $Query_System_Services = $Query_System_Services -replace ">$_<",($Custom_LogicServiceColour.$_)
        
    }

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Applications Tab :: Queries and Logic :: [Table 1]
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Applications Tab :: Queries and Logic :: [Table 2]
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Applications Tab :: Queries and Logic :: [Table 3]
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Deployment Tab :: Queries and Logic :: [Table 1]
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Deployment Tab :: Queries and Logic :: [Table 2]
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Deployment Tab :: Queries and Logic :: [Table 3]
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Messaging Tab :: Queries and Logic :: [Table 1]
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Messaging Tab :: Queries and Logic :: [Table 2]
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Messaging Tab :: Queries and Logic :: [Table 3]
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Security Tab :: Queries and Logic :: [Table 1]
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Security Tab :: Queries and Logic :: [Table 2]
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Security Tab :: Queries and Logic :: [Table 3]
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Advanced Tab :: Queries and Logic :: User Profile Service
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Update-ProgressBar -Activity "Generating Report" -Status "Advanced Tab: User Profile Service" -PercentComplete 80;

    $Query_Advanced_UserProfileService = Get-Eventlog -Computername $ThisPC -LogName System -EntryType "Error", "Warning" -After $Custom_Time -InstanceId "1500", "1511", "1530", "1533", "1534", "1542" | Select EventID, EntryType, Source, TimeGenerated, Message | Sort-Object EntryType, TimeGenerated | ConvertTo-HTML -Fragment -PreContent (Add-PreContentMessage -Type "Information" -Message " Displaying the following Event ID's only: 1500, 1511, 1530, 1533, 1534, 1542.");

    $Custom_LogicEventColour.Keys | ForEach {
    
        $Query_Advanced_UserProfileService = $Query_Advanced_UserProfileService -replace ">$_<",($Custom_LogicEventColour.$_)
        
    }

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Required :: Calculate Errors in Session for Error Dump (Advanced Tab)
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Update-ProgressBar -Activity "Generating Report" -Status "Advanced Tab: Verbose Error Dump" -PercentComplete 90;

    $ErrorArray = @();

    For($i = $Error.Count; $i -gt 0; $i--){

        [String]$ErrorItem = ($Error[$i -1].Exception.ToString(),$Error[$i -1].InvocationInfo.Line.ToString())
  
        $ErrorArray += $ErrorItem 

    }

    $ErrorDump = $($ErrorArray | ForEach-Object {[PSCustomObject]@{'Error' = $_}}) | ConvertTo-HTML -Fragment

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Required :: HTML :: DOM Elements :: Head Tag :: Includes Title, Styles
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Update-ProgressBar -Activity "Generating Report" -Status "Generating HTML Head" -PercentComplete 90;

    $Head = "
    <title>$ReportTitle</title>
    <!--
        Author  : $Author
        Date    : $CurrentDate
        Computer: $ThisPC
        Company : $CompanyName
        Version : $ReportVersion
    -->
    <style>

        body {
            background-color: #fff;
            padding: 30px;
            font-family: Arial;
            color: #4F4C4C;
        }

        .modal {
            display: none;
            position: fixed;
            z-index: 1;
            padding-top: 100px;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            overflow: auto;
            background-color: rgba(0,0,0,0.4);
        }

        .modal-content {
            position: relative;
            background-color: #fff;
            margin: auto;
            padding: 0;
            border: 1px solid #aaa;
            width: 50%;
            box-shadow: 0 4px 8px 0 rgba(0,0,0,0.2),0 6px 20px 0 rgba(0,0,0,0.19);
            -webkit-animation-name: animatetop;
            -webkit-animation-duration: 0.4s;
            animation-name: animatetop;
            animation-duration: 0.4s;
        }

        @-webkit-keyframes animatetop {
            from {top:-300px; opacity:0} 
            to {top:0; opacity:1}
        }

        @keyframes animatetop {
            from {top:-300px; opacity:0}
            to {top:0; opacity:1}
        }

        .modal-close {
            color: #aaa;
            float: right;
            padding-top: 4px;
            font-size: 28px;
            font-weight: bold;
        }

        .modal-close:hover,
        .modal-close:focus {
            color: #aaa;
            text-decoration: none;
            cursor: pointer;
        }

        .modal-header {
            padding: 2px 16px;
            background-color: #ededed;
            border-bottom: 1px solid #aaa;
            background-image: linear-gradient(to bottom,#fff 0%,#e0e0e0 100%);
            text-shadow: 0 1px 0 #fff;
        }

        .modal-body {
            padding: 2px 16px;
        }

        .modal-footer {
            padding: 2px 16px;
            background-color: #ededed;
            border-top: 1px solid #aaa;
            font-size: 12px;
            font-weight: normal;
            text-align: right;
        }

        #company-logo {
            position: absolute;
            background-repeat: no-repeat;
            width: $CompanyLogoWidth;
            height: $CompanyLogoHeight;
            top: $CompanyLogoTop;
            right: $CompanyLogoRight;
            border: $CompanyLogoBorder;
            background-image: url('data:image/png;base64,$CompanyLogoBase64String');
        }

        .button-clicked {
            background-image: linear-gradient(to bottom,#e0e0e0 0%,#f7f7f7 100%);
        }

        button {
            background-color: #4CAF50;
            background-image: linear-gradient(to bottom,#fff 0%,#e0e0e0 100%);
            border: 1px solid #aaa;
            border-radius: 4px;
            padding: 5px 10px;
            text-align: center;
            text-decoration: none;
            display: inline-block;
            margin-right: 5px;
            cursor: pointer;
            box-shadow: 0px 0px 2px 1px rgba(221,221,221,0.79);
            text-shadow: 0 1px 0 #fff;
        }

        button:focus {
            outline: #aaa;
        }

        button:hover {
            /*box-shadow: 0px 0px 2px 1px rgba(163, 204, 240.79);*/
        }

        details {
            margin-top: 5px;
            margin-bottom: 5px;
            border: 1px solid #aaa;
            border-radius: 4px;
            padding: .5em .5em 0;
            box-shadow: 0px 0px 6px 1px rgba(221,221,221,0.79);
        }

        details:hover {
            cursor: pointer;
            box-shadow: 0px 0px 6px 1px rgba(163, 204, 240.79);
        }

        details:focus {
            outline: #aaa;
        }

        details[open] {
             padding: .5em;
        }

        details[open] summary {
            border-bottom: 1px solid #aaa;
            margin-bottom: .5em;
         }

        summary {
            margin: -.5em -.5em 0;
            padding: .5em;
        }

        summary:focus {
            outline: #aaa;
        }
     
        .pre-content-alert-information { /* Blue */
            padding-top: 3px;
            padding-bottom: 5px;
            padding-left: 5px;
            padding-right: 5px;
            margin-bottom: 10px;
            border: 1px solid #39739d;
            /*background: linear-gradient(180deg, rgba(255,255,255,1) 12%, rgba(209,239,255,1) 100%);*/
            /*background-color: rgb(209,239,255);*/
            background-color: #e1ecf4;
        }

        .pre-content-alert-image-information {
            background-repeat: no-repeat;
            width: 16px;
            height: 16px;
            display: inline-block;
            margin-right: 6px;
            vertical-align:middle;
            background-image: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQBAMAAADt3eJSAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAAMFBMVEUAAADC0NxgkLcgZJoATozw8/ZQhbAAUZEAY7IAcMkAeNcwb6EAW6MAddL///////8+qcZNAAAAAXRSTlMAQObYZgAAAAFiS0dEDxi6ANkAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAHdElNRQflAxQMCDD1nNlJAAAAV0lEQVQI12NgYBAycVZkAIKwilkr21MZGFj33Fq1au3pAAa2U6vWvVq1JoFB/BaIsbaQQWMVGDQxWK4CiayazOAFYSxBMOBSGhBGE0I73EC4FXBLYc4AAJDbQkcK4XJeAAAAJXRFWHRkYXRlOmNyZWF0ZQAyMDIxLTAzLTIwVDEyOjA4OjQ3KzAwOjAww7cw4QAAACV0RVh0ZGF0ZTptb2RpZnkAMjAyMS0wMy0yMFQxMjowODo0NyswMDowMLLqiF0AAAAASUVORK5CYII=');
        }

        .pre-content-alert-text-information {
            vertical-align:middle;
            font-weight: normal;
            font-size: 14px;
            color: #39739d;
        }

        .pre-content-alert-error {
            padding-top: 3px;
            padding-bottom: 5px;
            padding-left: 5px;
            padding-right: 5px;
            margin-bottom: 10px;
            border: 1px solid #721C24;
            /*background: linear-gradient(180deg, rgba(255,255,255,1) 12%, rgba(209,239,255,1) 100%);*/
            /*background-color: rgb(209,239,255);*/
            background-color: #F8D7DA;
        }

        .pre-content-alert-image-error { /* Red */
            background-repeat: no-repeat;
            width: 16px;
            height: 16px;
            display: inline-block;
            margin-right: 6px;
            vertical-align:middle;
            background-image: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAMAAAAoLQ9TAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAAM1BMVEUAAADjvLfGd2uwQTClJhPBal2qJxPLMBXiNhbwOhe1Tj+8LBTrORfzX0P////5taj///9ex7kGAAAAAXRSTlMAQObYZgAAAAFiS0dEEJWyDSwAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAAHdElNRQflBRsJJyVbtSLSAAAAYUlEQVQY02VPyREAIQhDRBQF6b/b9frskAdDCEcAWEiYiTImeChcm0irXC7vQ+VARz/6ELsFk7F6EqtN39ynKSfAulN/oSLkdkQ/XFoGktt+x4RiIYyEpeFsMBatx+f+738QbAelah2KdAAAACV0RVh0ZGF0ZTpjcmVhdGUAMjAyMS0wNS0yN1QwOTozOTozNiswMDowMAPGksAAAAAldEVYdGRhdGU6bW9kaWZ5ADIwMjEtMDUtMjdUMDk6Mzk6MzYrMDA6MDBymyp8AAAAAElFTkSuQmCC');
        }

        .pre-content-alert-text-error {
            vertical-align:middle;
            font-weight: normal;
            font-size: 14px;
            color: #721C24;
        }

        .pre-content-alert-warning { /* Amber */
            padding-top: 3px;
            padding-bottom: 5px;
            padding-left: 5px;
            padding-right: 5px;
            margin-bottom: 10px;
            border: 1px solid #856404;
            /*background: linear-gradient(180deg, rgba(255,255,255,1) 12%, rgba(209,239,255,1) 100%);*/
            /*background-color: rgb(209,239,255);*/
            background-color: #FFF3CD;
        }

        .pre-content-alert-image-warning {
            background-repeat: no-repeat;
            width: 16px;
            height: 16px;
            display: inline-block;
            margin-right: 6px;
            vertical-align:middle;
            background-image: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQBAMAAADt3eJSAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAAIVBMVEUAAAD/7sH/+/D+zQD/uQD/wSD84QD83wAAAAD91gD///+RdSDdAAAAAXRSTlMAQObYZgAAAAFiS0dECmjQ9FYAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAHdElNRQflBRsJIihYc6oqAAAAYUlEQVQI12NgYGBgFGCAAGUjKEPEEUIzBZsqgBmMbiUQRapmyUEQJRntjhAlGW1gRYxuGW1gRapmGW1gRSJpGW3pjiAlaUAAVMToBhRJAypSNQMxgIpMZoKBM0OICxi4AgC5GRmzSdeRMAAAACV0RVh0ZGF0ZTpjcmVhdGUAMjAyMS0wNS0yN1QwOTozNDozOCswMDowMKYHYi0AAAAldEVYdGRhdGU6bW9kaWZ5ADIwMjEtMDUtMjdUMDk6MzQ6MzgrMDA6MDDXWtqRAAAAAElFTkSuQmCC');
        }

        .pre-content-alert-text-warning {
            vertical-align:middle;
            font-weight: normal;
            font-size: 14px;
            color: #856404;
        }

        table {
            border: 0px;
            border-collapse: collapse;
        }
    
        th { 
            padding: 5px;
            border: 1px solid #aaa;
            /*text-shadow: 0 1px 0 #fff;*/
            /*background-image: linear-gradient(to bottom,#fff 0%,#e0e0e0 100%);*/
            background-color: #e0e0e0;
        }

        tr:nth-child(odd) {
            background-color:#ededed;
        } 
    
        tr:nth-child(even) {
            background-color: #FFFFFF;
        }
       
        td {
            border: 1px solid #aaa;
            padding: 5px;
        }

        .table-cell-green {
            background-color: #D4EDDA; color: #15572a;
        }

        .table-cell-amber {
            background-color: #FFF3CD; color: #856404;
        }

        .table-cell-red {
            background-color: #F8D7DA; color: #721C24;
        }

    </style>
    "

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Required :: HTML :: DOM Elements :: Company Logo, Buttons & Tabs, Sections, Details, Summary
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Update-ProgressBar -Activity "Generating Report" -Status "Generating HTML Body" -PercentComplete 90;

    $Modal = "
    <aside id='modal' class='modal'>
        <div class='modal-content'>
            <div class='modal-header'>
                <span class='modal-close'>&times;</span>
                <h2>$CompanyName - $ReportTitle</h2>
            </div>
            <div class='modal-body'>
                <p>Computer Name: $ThisPC</p>
            </div>
            <div class='modal-footer'>
                <p>Report Version: $ReportVersion</p>
            </div>
        </div>
    </aside>
    "

    $CompanyLogo = "
    <aside id='company-logo'></aside>
    "

    $ButtonElements = "
    <nav>
        <button class='button-clicked' onclick=opentab('tab-system');navBtnClick(this);>System</button>
        <button onclick=opentab('tab-applications');navBtnClick(this);>Applications</button>
        <button onclick=opentab('tab-deployment');navBtnClick(this);>Deployment</button>
        <button onclick=opentab('tab-messaging');navBtnClick(this);>Messaging</button>
        <button onclick=opentab('tab-security');navBtnClick(this);>Security</button>
        <button onclick=opentab('tab-advanced');navBtnClick(this);>Advanced</button>
        <button onclick=openModal()>Report Information</button>
    </nav> 
    "

    $TabSystem = "
    <section id='tab-system' class='tab'>
        <h2>Health Check: System</h2>
        <details>
            <summary>BIOS</summary>
            $Query_System_BIOS
        </details>
        <details>
            <summary>Battery</summary>
            $Query_System_Battery
        </details>
        <details>
            <summary>Operating System</summary>
            $Query_System_OS
        </details>
        <details>
            <summary>Central Processing Unit (CPU)</summary>
            $Query_System_CPU
        </details>
        <details>
            <summary>Disk</summary>
            $Query_System_Disk
        </details>
        <details>
            <summary>Event Log: Application</summary>
            $Query_System_AppEvent
        </details>
        <details>
            <summary>Event Log: System</summary>
            $Query_System_SysEvent
        </details>
        <details>
            <summary>Hot Fix</summary>
            $Query_System_Hotfix
        </details>
        <details>
            <summary>Installed Applications</summary>
            $Query_System_InstalledApps
        </details>
        <details>
            <summary>Services</summary>
            $Query_System_Services
        </details>
    </section>

    <section id='tab-applications' class='tab' style='display:none'>
        <h2>Health Check: Applications</h2> 
    </section>

    <section id='tab-deployment' class='tab' style='display:none'>
        <h2>Health Check: Deployment</h2>
    </section>

    <section id='tab-messaging' class='tab' style='display:none'>
        <h2>Health Check: Messaging</h2>
    </section>

    <section id='tab-security' class='tab' style='display:none'>
        <h2>Health Check: Security</h2>
    </section>

    <section id='tab-advanced' class='tab' style='display:none'>
        <h2>Health Check: Advanced</h2>       
        <details>
            <summary>User Profile Service</summary>
            $Query_Advanced_UserProfileService
        </details>
        <details>
            <summary>Error Dump (Verbose Script Debugging Information)</summary>
            $ErrorDump
        </details>
    </section>
    "

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Required :: HTML :: DOM Elements :: Script Tag
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Update-ProgressBar -Activity "Generating Report" -Status "Generating HTML Scripts" -PercentComplete 90;

    $Scripts = "
    <script>

        // -------------------------------------------------------------------------------------------------------------------------------------
        // Optional Script --> HTML Cleanup : Remove Collections and Colection Groups <col> <colgroup> from DOM
        // -------------------------------------------------------------------------------------------------------------------------------------

            var col = document.getElementsByTagName('col'), index;
            for (index = col.length - 1; index >= 0; index--) {
                col[index].parentNode.removeChild(col[index]);
            }
            var colgroup = document.getElementsByTagName('colgroup'), index;
            for (index = colgroup.length - 1; index >= 0; index--) {
                colgroup[index].parentNode.removeChild(colgroup[index]);
            }

        // -------------------------------------------------------------------------------------------------------------------------------------
        // Required Script --> Tab Button Function System : Displays and Hides Tab Sections on Nav Button Click
        // -------------------------------------------------------------------------------------------------------------------------------------

            function opentab(tabName) {
                var i;
                var x = document.getElementsByClassName('tab');
                for (i = 0; i < x.length; i++) {
                x[i].style.display = 'none';  
                }
                document.getElementById(tabName).style.display = 'block'; 
            }

        // -------------------------------------------------------------------------------------------------------------------------------------
        // Required Script --> Nav Button Click Function : Toggles Pressed / De-Pressed Visual State of button by adding or removing css class
        // -------------------------------------------------------------------------------------------------------------------------------------
       
            function navBtnClick(elem) {
                var btn = document.getElementsByTagName('button');
                for (i = 0; i < btn.length; i++) {
                    btn[i].classList.remove('button-clicked')
                }
                elem.classList.add('button-clicked');
            }

        // -------------------------------------------------------------------------------------------------------------------------------------
        // Required Script --> Modal (Report Information Button)
        // -------------------------------------------------------------------------------------------------------------------------------------

            var modal = document.getElementById('modal');
            function openModal(){
                modal.style.display = 'block';
            }
            var modalClose = document.getElementsByClassName('modal-close')[0];
            modalClose.onclick = function() {
              modal.style.display = 'none';
            }
            window.onclick = function(event) {
                if (event.target == modal) {
                    modal.style.display = 'none';
                }
            }
            modal.style.display = 'block';

        // ------------------------------------------------------------------------------------------------------------------------------------
        // Custom Scripts: --> Disk Space Table Update : Set cell colour based on % Free Space
        // ------------------------------------------------------------------------------------------------------------------------------------

            var rows  = document.getElementsByTagName('table')[4].rows; // 5th Table from top, Index [4]
            var last  = rows[rows.length - 1];
            var cell  = last.cells[4]; // 5th cell from left, Index [4]
            var first = last.cells[0]; // Cell 1, Index [0]
        
            if (first.innerHTML === 'C:'){
                var value = cell.innerHTML;
                if (value > 30)                 { cell.classList.add('table-cell-green');};
                if (value > 10 && value <= 30)  { cell.classList.add('table-cell-amber');};
                if (value <= 10)                { cell.classList.add('table-cell-red')  ;};
            } else {
                alert ('The Disk Table has changed position and will not be colour coded. Please contact the author of this report to resolve the problem.');
            }

    </script>
    "

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Required :: Build :: Build HTML Page and Launch
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Update-ProgressBar -Activity "Generating Report" -Status "Compiling HTML Report" -PercentComplete 95;

    ConvertTo-HTML -Head $Head -Body "$Modal $CompanyLogo $ButtonElements $TabSystem $Scripts" -Title $ReportTitle | Out-File $ReportLocation$ReportFileName

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Custom :: Clean Up :: Clear Variables
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Update-ProgressBar -Activity 'Generating Report' -Status "Script Clean up" -PercentComplete 99;
    
        # $Error.Clear()  <# Custom #>

        # Clear-Variable  <# Custom #>

        # Remove-Variable <# Custom #>

# -----------------------------------------------------------------------------------------------------------------------------------------------------------
# Custom :: Launch Report
# -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Update-ProgressBar -Activity "Generating Report" -Status "Complete" -PercentComplete 100;

    If ($LaunchReportOnCompletion -eq $True) {

        Try     { Start-Process $ReportLocation$ReportFileName }
        Catch   {}
        Finally {}

    }