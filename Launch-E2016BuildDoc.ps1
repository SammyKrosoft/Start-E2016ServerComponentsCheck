<#
.SYNOPSIS
    Meet the Exchange 2016 Build Document Generator !

    This program launches a Graphical Interface that lets you update a Word template
    with a choice of custom values that you populated in an Excel spreadsheet.

.DESCRIPTION
    This is a program that takes values from an Excel Workbook, list all tabs, lets
    you choose from which Tab you want the values from, and update an MS Word template
    with the values from the selected Excel Tab.

    See the quick user guide for more information.

.PARAMETER NoNeedToCheckMSWord
    This switch bypasses the MS Word check to gain time if you are SURE and CERTAIN to 
    have MS Word 2013 or later version.

.PARAMETER CheckVersion
	This parameter dumps the current script version - the script stops processing after displaying the
	version if this parameter is specified, no matter what other parameter is also specified.

.INPUTS
    Chosen from the GUI.

.OUTPUTS
    MS Word report, appended with customer / department name, on an OUTPUT directory located
    under the script's directory.

.EXAMPLE
.\Launch-E2016BuildDoc.ps1

.EXAMPLE
.\Launch-E2016BuildDoc.ps1 -NoNeedToCheckMSWord

.NOTES
None

.LINK
    https://github.com/SammyKrosoft
#>
[CmdletBinding(DefaultParameterSetName="NormalRun")]
Param(
	[Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")] [switch]$NoNeedToCheckMSWord,
	[Parameter(Mandatory = $False, Position = 2, ParameterSetName = "checkversion")] [Switch] $CheckVersion
)

<# ------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) ------- #>
# Using Write-Debug and playing with $DebugPreference -> "Continue" will output whatever you put on Write-Debug "Your text/values"
# and "SilentlyContinue" will output nothing on Write-Debug "Your text/values"
$DebugPreference = "Continue"
# Set Error Action to your needs
$ErrorActionPreference = "SilentlyContinue"
#Script Version
$ScriptVersion = "1"
<# Version History
v1.0 -> v2
Added export of Outlook Anywhere with External Hostname (E2010, E2013, E2016) and Internal Hostname (not existing in E2010)
Fixed output file name
Added -DoNoExport switch, to not export to a file...
#> 
If ($CheckVersion) {Write-Host "Script Version v$ScriptVersion";exit}
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>

Function Get-PowerShellVersion {$PSVersion = $PSVersionTable.PSVersion.Major;Return $PSVersion}

Function Get-MSWordVersion {
    LogMag "Please wait while checking your MS Word version..."
    $MSWord = New-Object -ComObject Word.Application
    $MSWordversion = $MSWord.Version
    #Quitting Word gracefully, freeing the COM object and cleaning the variable
    $MSWord.Quit()
    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$MSword)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Remove-Variable MSword
    LogGreen "MSWord version installed is : $MSWordVersion"
    If ($MSWordVersion -ge 15){
        LogGreen "MSWord version is greater than 2013, we're good to go !"
    } Else {
        LogYellow "Alas, your MSWord version is older than 2013 ... exiting"
        exit
    }    
}



Function Get-MSWordVersionWPF {
    # Load a WPF GUI from a XAML file build with Visual Studio
    Add-Type -AssemblyName presentationframework, presentationcore
    $wpf = @{ }
    # NOTE: Either load from a XAML file or paste the XAML file content in a "Here String"
    #$inputXML = Get-Content -Path ".\WPFGUIinTenLines\MainWindow.xaml"
    $inputXML = @"
<Window x:Name="WPFProgress" x:Class="Just_a_progress_bar.MainWindow"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:local="clr-namespace:Just_a_progress_bar"
    mc:Ignorable="d"
    Title="Checking MS Word version" Height="104.545" Width="616.162">
<Grid>
    <ProgressBar x:Name="ProgressBar01" HorizontalAlignment="Left" Height="35" Margin="10,32,0,0" VerticalAlignment="Top" Width="591" Foreground="#FF06A1B0" Background="#FFBDB5B5"/>
    <TextBlock x:Name="txtStatus" HorizontalAlignment="Left" Margin="10,10,0,0" TextWrapping="Wrap" Text="Please wait..." VerticalAlignment="Top" Width="591"/>
</Grid>
</Window>
"@

    $inputXMLClean = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''
    [xml]$xaml = $inputXMLClean
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $tempform = [Windows.Markup.XamlReader]::Load($reader)
    $namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
    $namedNodes | ForEach-Object {$wpf.Add($_.Name, $tempform.FindName($_.Name))}

    #Get the form name to be used as parameter in functions external to form...
    $FormName = $NamedNodes[0].Name

    LogMag "Please wait while checking your MS Word version..."
    $MSWord = New-Object -ComObject Word.Application
    $MSWordversion = $MSWord.Version
    #Quitting Word gracefully, freeing the COM object and cleaning the variable
    $MSWord.Quit()
    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$MSword)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Remove-Variable MSword
    LogGreen "MSWord version installed is : $MSWordVersion"
    If ($MSWordVersion -ge 15){
        LogGreen "MSWord version is greater than 2013, we're good to go !"
        Return $True
    } Else {
        LogYellow "Alas, your MSWord version is older than 2013 ... exiting"
        Return $False
        exit
    }    

    #Define events functions
    #region Load, Draw (render) and closing form events
    #Things to load when the WPF form is loaded aka in memory
    $wpf.$FormName.Add_Loaded({
        #Update-Cmd
    })
    #Things to load when the WPF form is rendered aka drawn on screen
    $wpf.$FormName.Add_ContentRendered({
        #Update-Cmd
    })
    $wpf.$FormName.add_Closing({
        $msg = "Now launching the real thing..."
        write-host $msg
    })
    # End of load, draw and closing form events
    #endregion Load, Draw and closing form events

    #HINT: to update progress bar and/or label during WPF Form treatment, add the following:
    # ... to re-draw the form and then show updated controls in realtime ...
    $wpf.$FormName.Dispatcher.Invoke("Render",[action][scriptblock]{})


    # Load the form:
    # Older way >>>>> $wpf.MyFormName.ShowDialog() | Out-Null >>>>> generates crash if run multiple times
    # Newer way >>>>> avoiding crashes after a couple of launches in PowerShell...
    # USing method from https://gist.github.com/altrive/6227237 to avoid crashing Powershell after we re-run the script after some inactivity time or if we run it several times consecutively...
    $async = $wpf.$FormName.Dispatcher.InvokeAsync({
        $wpf.$FormName.ShowDialog() | Out-Null
    })
    $async.Wait() | Out-Null
}



Function Get-PowerShellVersion {
    $PSVer = $PSVersionTable.PSVersion.Major
    If ($PSVer -ge 3){
        LogGreen "PowerShell version is greater than 3, that's good ! Continuing..."
    } Else {
        LogMag "Alas your PowerShell version is lower than 3... exiting"
        exit
    }
}


Function Get-ExcelWorkSheetsNamesWPFGUIUpdate {
    <#
    .SYNOPSIS
        Get Excel worksheets names

    .DESCRIPTION
        Get Excel worksheets names and return a collection of names.
        We will use it to populate a WPF listbox where the user will be able
        to select the customer/department to update.
    #>
    [CmdLetBinding(DefaultParameterSetName = "NormalRun")]
    Param(
        [Parameter(Mandatory = $false, Position = 1, ParameterSetName = "NormalRun")][string]$ExcelInput,
        [Parameter(Mandatory = $false, Position = 2, ParameterSetName = "CheckOnly")][switch]$CheckVersion
    )

    <# ------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) ------- #>
    #Initializing a $Stopwatch variable to use to measure script execution
    $stopwatch2 = [system.diagnostics.stopwatch]::StartNew()
    #Using Write-Debug and playing with $DebugPreference -> "Continue" will output whatever you put on Write-Debug "Your text/values"
    # and "SilentlyContinue" will output nothing on Write-Debug "Your text/values"
    $DebugPreference = "Continue"
    # Set Error Action to your needs
    $ErrorActionPreference = "SilentlyContinue"
    #Script Version
    $ScriptVersion = "0.1"
    <# Version changes
    v0.1 : first script version
    v0.1 -> v0.5 : 
    #>
    $ScriptName = $MyInvocation.MyCommand.Name
    If ($CheckVersion) {Write-Host "SCRIPT NAME     : $ScriptName `nSCRIPT VERSION  : $ScriptVersion";exit}
    # Log or report file definition
    # NOTE: use $PSScriptRoot in Powershell 3.0 and later or use $LocalScriptExecPath = split-path -parent $MyInvocation.MyCommand.Definition in Powershell 2.0
    <# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
    
    $msg = "Checking file input..."
    $Percent = 0
    Update-WPFProgressBarAndStatus $msg $percent | out-null

    If (-Not $ExcelInput){
        $ExcelInput = ".\E2016Test.xlsx"
        LogMag "No Excel input file specified ... using default:" -b Yellow
        Write-Host $ExcelInput
    } Else {
        "Excel input file specified : $ExcelInput. Continuing ..." | Out-Host
    }

    $msg = "Checking if Excel file exists"
    $Percent = 0
    Update-WPFProgressBarAndStatus $msg $percent | out-null

    $FullXLFilePath = $PSScriptRoot + "\" + $ExcelInput
    $FileExists = Test-Path $FullXLFilePath

    LogBlue $FullXLFilePath
    If ($FileExists) {
        LogGreen "Excel file exists, continuing..."
    } Else {
        $msg = "Excel file $ExcelInput does not exist in current directory, exiting..."
        LogMag $msg
        [System.Windows.MessageBox]::Show($msg,"File not found","Ok","Error") | out-null
        Update-WPFProgressBarAndStatus "" 0 #Reset GUI progress bar
        Return $null #Return to GUI...
    }

    $msg = "Creating a Temporary file to work with..."
    $Percent = 5
    Update-WPFProgressBarAndStatus $msg $percent | out-null

    LogMag "Copying file to temp file to allow opening even if it's already opened"
    $TempXLFile =  $PSScriptRoot + "\" + "TempE2016BuildInputs.xlsx"
    Copy-item $FullXLFilePath -Destination $TempXLFile -Force
    LogBlue "Working copy becomes the Temporary file ..."
    $FullXLFilePath = $TempXLFile

    $msg = "Opening a new Excel instance..."
    $Percent = 10
    Update-WPFProgressBarAndStatus $msg $percent | out-null
    LogGreen $msg
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false

    $msg = "Opening Excel workbook..."
    $Percent = 20
    Update-WPFProgressBarAndStatus $msg $percent | out-null
    LogGreen $msg
    $Workbook = $Excel.Workbooks.Open($FullXLFilePath)
    $WorkSheetsObjectsList = $Workbook.Worksheets

    $msg = "Getting all worksheets (aka ""Tabs"") names..."
    $Percent = 40
    Update-WPFProgressBarAndStatus $msg $percent | out-null
    LogBlue $msg
    $WorkSheetsList = @()
    Foreach ($Worksheet in $WorkSheetsObjectsList){
        LogMag $($Worksheet.name)
        $WorkSheetsList += $($Worksheet.Name)
    }

    $msg = "Closing everything..."
    $Percent = 90
    Update-WPFProgressBarAndStatus $msg $percent | out-null
    Write-Host "Closing workbook..." -ForegroundColor Green
    $Workbook.Close()
    Write-Host "Releasing Workbook Com Object..." -ForegroundColor Green
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    Write-Host "Closing Excel..." -ForegroundColor Green
    $Excel.Quit()
    Write-Host "Releasing Excel Com Object..." -ForegroundColor Green
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "Cleaning Excel variable..." -ForegroundColor Green
    Remove-Variable excel
    Write-Host "Garbage Collection..." -ForegroundColor Green
    [System.GC]::Collect()
    Write-Host "WaitForPendingFinalizers..." -ForegroundColor Green
    [System.GC]::WaitForPendingFinalizers()

    <# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
    #Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
    $stopwatch2.Stop()
    $msg = "`n`nThe script took $([math]::round($($StopWatch2.Elapsed.TotalSeconds),2)) seconds to execute..."
    Write-Host $msg
    $msg = $null
    $StopWatch2 = $null
    <# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>

    $msg = "All closed !"
    $Percent = 100
    Update-WPFProgressBarAndStatus $msg $percent | out-null

    LogMag "Removing temporary file"
    Remove-Item $FullXLFilePath -Force

    Return $WorkSheetsList
}

Function Get-ExcelWorkSheetsNames {
    <#
    .SYNOPSIS
        Get Excel worksheets names

    .DESCRIPTION
        Get Excel worksheets names to populate list, dropdown or just validate...
    #>
    [CmdLetBinding(DefaultParameterSetName = "NormalRun")]
    Param(
        [Parameter(Mandatory = $false, Position = 1, ParameterSetName = "NormalRun")][string]$ExcelInput,
        [Parameter(Mandatory = $false, Position = 2, ParameterSetName = "CheckOnly")][switch]$CheckVersion
    )

    <# ------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) ------- #>
    #Initializing a $Stopwatch variable to use to measure script execution
    $stopwatch2 = [system.diagnostics.stopwatch]::StartNew()
    #Using Write-Debug and playing with $DebugPreference -> "Continue" will output whatever you put on Write-Debug "Your text/values"
    # and "SilentlyContinue" will output nothing on Write-Debug "Your text/values"
    $DebugPreference = "Continue"
    # Set Error Action to your needs
    $ErrorActionPreference = "SilentlyContinue"
    #Script Version
    $ScriptVersion = "0.1"
    <# Version changes
    v0.1 : first script version
    v0.1 -> v0.5 : 
    #>
    $ScriptName = $MyInvocation.MyCommand.Name
    If ($CheckVersion) {Write-Host "SCRIPT NAME     : $ScriptName `nSCRIPT VERSION  : $ScriptVersion";exit}
    # Log or report file definition
    # NOTE: use $PSScriptRoot in Powershell 3.0 and later or use $LocalScriptExecPath = split-path -parent $MyInvocation.MyCommand.Definition in Powershell 2.0
    <# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
    
    If (-Not $ExcelInput){
        $ExcelInput = ".\E2016Test.xlsx"
        LogMag "No Excel input file specified ... using default:" -b Yellow
        Write-Host $ExcelInput
    } Else {
        "Excel input file specified : $ExcelInput. Continuing ..." | Out-Host
    }

    $FullXLFilePath = $PSScriptRoot + "\" + $ExcelInput
    $FileExists = Test-Path $FullXLFilePath

    LogBlue $FullXLFilePath
    If ($FileExists) {
        LogGreen "Excel file exists, continuing..."
    } Else {
        $msg = "Excel file $ExcelInput does not exist in current directory, exiting..."
        LogMag $msg
        [System.Windows.MessageBox]::Show($msg,"File not found","Ok","Error") | out-null
        Return $null #Trying to return to GUI...
    }

    
    LogMag "Copying file to temp file to allow opening even if it's already opened"
    $TempXLFile =  $PSScriptRoot + "\" + "TempE2016BuildInputs.xlsx"
    Copy-item $FullXLFilePath -Destination $TempXLFile -Force
    LogBlue "Working copy becomes the Temporary file ..."
    $FullXLFilePath = $TempXLFile

    LogGreen "Opening a new Excel instance..."
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false

    LogGreen "Opening Excel workbook..."
    $Workbook = $Excel.Workbooks.Open($FullXLFilePath)
    $WorkSheetsObjectsList = $Workbook.Worksheets

    LogBlue "Getting all worksheets (aka ""Tabs"") names..."
    $WorkSheetsList = @()
    Foreach ($Worksheet in $WorkSheetsObjectsList){
        LogMag $($Worksheet.name)
        $WorkSheetsList += $($Worksheet.Name)
    }

    Write-Host "Closing workbook..." -ForegroundColor Green
    $Workbook.Close()
    Write-Host "Releasing Workbook Com Object..." -ForegroundColor Green
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    Write-Host "Closing Excel..." -ForegroundColor Green
    $Excel.Quit()
    Write-Host "Releasing Excel Com Object..." -ForegroundColor Green
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "Cleaning Excel variable..." -ForegroundColor Green
    Remove-Variable excel
    Write-Host "Garbage Collection..." -ForegroundColor Green
    [System.GC]::Collect()
    Write-Host "WaitForPendingFinalizers..." -ForegroundColor Green
    [System.GC]::WaitForPendingFinalizers()

    <# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
    #Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
    $stopwatch2.Stop()
    $msg = "`n`nThe script took $([math]::round($($StopWatch2.Elapsed.TotalSeconds),2)) seconds to execute..."
    Write-Host $msg
    $msg = $null
    $StopWatch2 = $null
    <# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>

    LogMag "Removing temporary file"
    Remove-Item $FullXLFilePath -Force

    Return $WorkSheetsList
}

Function LogMag {
    param(
        [Parameter(Mandatory = $false, Position = 1)][string]$message,
        [Parameter(Mandatory = $false)][string]$b = "black"
    )
    Write-Host $message -F Magenta -b $b
}

Function LogGreen {
    param(
        [Parameter(Mandatory = $false, Position = 1)][string]$message,
        [Parameter(Mandatory = $false)][string]$b = "black"
    )
    Write-Host $message -F Green -b $b
}

Function LogYellow {
    param(
        [Parameter(Mandatory = $false, Position = 1)][string]$message,
        [Parameter(Mandatory = $false)][string]$b = "black"
    )

    Write-Host $message -F Yellow -b $b
}

Function LogBlue {
    param(
        [Parameter(Mandatory = $false, Position = 1)][string]$message,
        [Parameter(Mandatory = $false)][string]$b = "black"
    )
    Write-Host $message -F Blue -b $b
}

Function Title1 ($title, $TotalLength = 100, $Back = "Yellow", $Fore = "Black") {
    $TitleLength = $Title.Length
    [string]$StarsBeforeAndAfter = ""
    $RemainingLength = $TotalLength - $TitleLength
    If ($($RemainingLength % 2) -ne 0) {
        $Title = $Title + " "
    }
    $Counter = 0
    For ($i=1;$i -le $(($RemainingLength)/2);$i++) {
        $StarsBeforeAndAfter += "*"
        $counter++
    }
    
    $Title = $StarsBeforeAndAfter + $Title + $StarsBeforeAndAfter
    Write-host
    Write-Host $Title -BackgroundColor $Back -foregroundcolor $Fore
    Write-Host
    
}

Function Update-WPFProgressBarAndStatus {
    Param(  [parameter(Position = 1)][string]$msg="Message",
            [parameter(Position=2)][int]$p=50,
            [parameter(Position = 3)][string]$status="Working...",
            [parameter(position = 4)][string]$color = "#FFC310BB",
            [parameter(position = 5)][string]$ProgressBarName = "ProgressBar")
    $wpf.$ProgressBarName.Color = $Color
    $wpf.$ProgressBarName.Value = $p
    $wpf.$ProgressBarName.Foreground
    Title1 $msg; StatusLabel $msg
    If ($p -eq 100){
        $status = "Done!"
    }
    Write-progress -Activity $msg -Status $status -PercentComplete $p
}


Function HereStringToArray ($HereString) {
    Return $HereString -split "`n" | %{$_.trim()}
}


#Function to compare fields from 2 arrays
Function Compare2Arrays {
    [CmdLetBinding()]
    Param(
        [Parameter(Position = 1)][array]$ReqFields,
        [Parameter(Position = 2)][array]$FieldsToCompareToReqFields
    )
    #comparing each formfield in the doc with the form field names defined to check if no one is missing
    LogYellow "There are $($ReqFields.count) text fields to check, the document contains $($FieldsToCompareToReqFields.count) Text Formfields" -b blue

    If ($($ReqFields.count) -ne $($FieldsToCompareToReqFields.count)){
    LogBlue "There is mismatch in the number of fields" -B yellow
    }

    $NbMatches = 0
    $MissingField = @()
    $found = $false
    $AtLeastOneMissing = $false
    Foreach ($chkitem in $ReqFields) {
        LogBlue "Checking $chkitem"
        Foreach ($docitem in $FieldsToCompareToReqFields){
            If ($chkItem -eq $($DocItem.'FormField Name')){
                LogGreen "This Field is in the Doc !"
                $found = $true
                $NbMatches += 1
            }
        }
        If (-not $found){
            LogMag "$chkitem not found in the document ..."
            $MissingField += $chkitem
            $AtLeastOneMissing = $True
        }
        $found = $false
    }

    If ($AtLeastOneMissing){
        Write-Host "At least one field is missing in the Doc" -ForegroundColor red -BackgroundColor yellow
        Write-Host "There are $($MissingField.count) fields missing in the doc"
        $MissingField | Out-Host
        Return $false
    } Else {
        Write-Host "All fields there !"
        Return $True
    }
}
    
#Function from Sam to get values from Exchange 2016
Function Get-E2016ReportValues {
    <#
    .SYNOPSIS
        Get Excel table data to be updated in the Word document

    .DESCRIPTION
        Get Excel table data to be updated in the Word document
    #>
    [CmdLetBinding(DefaultParameterSetName = "NormalRun1")]
    Param(
        [Parameter(Mandatory = $false, Position = 1, ParameterSetName = "NormalRun1")][string]$ExcelInput,
        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = "NormalRun1")][string]$Department,
        [Parameter(Mandatory = $false, Position = 3, ParameterSetName = "CheckOnly1")][switch]$CheckVersion
    )

    <# ------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) ------- #>
    #Initializing a $Stopwatch variable to use to measure script execution
    $stopwatch2 = [system.diagnostics.stopwatch]::StartNew()
    #Using Write-Debug and playing with $DebugPreference -> "Continue" will output whatever you put on Write-Debug "Your text/values"
    # and "SilentlyContinue" will output nothing on Write-Debug "Your text/values"
    $DebugPreference = "Continue"
    # Set Error Action to your needs
    $ErrorActionPreference = "SilentlyContinue"
    #Script Version
    $ScriptVersion = "0.1"
    <# Version changes
    v0.1 : first script version
    v0.1 -> v0.5 : 
    #>
    $ScriptName = $MyInvocation.MyCommand.Name
    If ($CheckVersion) {Write-Host "SCRIPT NAME     : $ScriptName `nSCRIPT VERSION  : $ScriptVersion";exit}
    # Log or report file definition
    # NOTE: use $PSScriptRoot in Powershell 3.0 and later or use $LocalScriptExecPath = split-path -parent $MyInvocation.MyCommand.Definition in Powershell 2.0
    <# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
    
    If (-Not $ExcelInput){
        $ExcelInput = "C:\Users\SammyKrosoft\OneDrive\_Boulot\How-To Procedures\Exchange 2016 docs\E2016Test.xlsx"
        Write-Host "No Excel input file specified ... using default:" -BackgroundColor Yellow
        Write-Host $ExcelInput
    } Else {
        "Excel input file specified : $ExcelInput. Continuing ..." | Out-Host
    }

    $FullXLFilePath = $PSScriptRoot + "\" + $ExcelInput
    $FileExists = Test-Path $FullXLFilePath

    LogBlue $FullXLFilePath
    If ($FileExists) {
        LogGreen "Excel file exists, continuing..."
    } Else {
        $msg = "Excel file does not exist, exiting..."
        LogMag $msg
        [System.Windows.MessageBox]::Show($msg)
        $wpf.$FormName.IsEnabled = $true
        StatusLabel "Ready !"
        Return "FileNotFound"#Return to GUI...
    }

    LogMag "Copying file to temp file to allow opening even if it's already opened"
    $TempXLFile =  $PSScriptRoot + "\" + "TempE2016BuildInputs.xlsx"
    Copy-item $FullXLFilePath -Destination $TempXLFile -Force
    LogBlue "Working copy becomes the Temporary file ..."
    $FullXLFilePath = $TempXLFile

    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false

    $Workbook = $Excel.Workbooks.Open($FullXLFilePath)
    $WorkSheet = $Workbook.Worksheets.item($Department)
    if ($WorkSheet) {
        LogMag "Department exists !" -b yellow 
    } Else {
        LogMag "No Excel Tab named $Department ..." -b blue
        Write-Host "Closing workbook..." -ForegroundColor Green
        $Workbook.Close()
        Write-Host "Releasing Workbook Com Object..." -ForegroundColor Green
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
        Write-Host "Closing Excel..." -ForegroundColor Green
        $Excel.Quit()
        Write-Host "Releasing Excel Com Object..." -ForegroundColor Green
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
        Write-Host "Cleaning Excel variable..." -ForegroundColor Green
        Remove-Variable excel
        Write-Host "Garbage Collection..." -ForegroundColor Green
        [System.GC]::Collect()
        Write-Host "WaitForPendingFinalizers..." -ForegroundColor Green
        [System.GC]::WaitForPendingFinalizers()
        Return "NoTab"
    }

    $Worksheet.Activate()
    $WSTable = $Worksheet.ListObjects.Item(1)

    $WSTableRows = $WSTable.ListRows

    #$WSTableRows.Count | Out-host
    #$Row = $WSTableRows[1]
    #$RowVals = $Row.Range
    write-host "IN THE EXCEL Function !"
    $WholeInputCollection = @()
    ForEach ($Row in $WSTableRows){
        $ValTrio = @() #Init & Re-init variable as we just want to store the values from each Row
        # there will be 3 columns that is 3 values for each Row
        Foreach ($Val in $($Row.Range)){
            #Write-Host $($Val.Text)
            $ValTrio += $Val.Text
        }
        #Write-Host "Trio is : $($ValTrio[0]),$($ValTrio[1]),$($ValTrio[2]) "
        $CustomObj = [PSCustomObject]@{
            Description = $($ValTrio[0])
            Value = $($ValTrio[1])
            BookMark = $($ValTrio[2])
        }
        $WholeInputCollection += $CustomObj
    }

    LogBlue "Tags loaded from Excel :" -b yellow
    $WholeInputCollection | Out-Host

    Write-Host "Closing workbook..." -ForegroundColor Green
    $Workbook.Close()
    Write-Host "Releasing Workbook Com Object..." -ForegroundColor Green
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
    Write-Host "Closing Excel..." -ForegroundColor Green
    $Excel.Quit()
    Write-Host "Releasing Excel Com Object..." -ForegroundColor Green
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    Write-Host "Cleaning Excel variable..." -ForegroundColor Green
    Remove-Variable excel
    Write-Host "Garbage Collection..." -ForegroundColor Green
    [System.GC]::Collect()
    Write-Host "WaitForPendingFinalizers..." -ForegroundColor Green
    [System.GC]::WaitForPendingFinalizers()

    <# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
    #Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
    $stopwatch2.Stop()
    $msg = "`n`nThe script took $([math]::round($($StopWatch2.Elapsed.TotalSeconds),2)) seconds to execute..."
    Write-Host $msg
    $msg = $null
    $StopWatch2 = $null
    <# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>

    LogMag "Removing temporary file"
    Remove-Item $FullXLFilePath -Force

    Return $WholeInputCollection
}
    
Function Update-Cmd {
    $WordDocument = $wpf.txtDocFileName.Text
    $ExcelInputs = $wpf.txtExcelFileName.Text
    $DepartmentAcronym = $wpf.txtCustomerAcronym.Text

    If ((-not $WordDocument) -or (-not $ExcelInputs) -or (-not $DepartmentAcronym)){
        $wpf.btnRun.IsEnabled = $false
        $Missing = @()
        If (-not $WordDocument){$Missing += "Word Document"}
        If (-not $ExcelInputs) {$Missing += "Excel Document"}
        If (-not $DepartmentAcronym) {$Missing += "Department name"}
        If ($Missing.count -gt 1){
            $Missing = $Missing -join ", "
        }
        $CmdLine = "Missing info: $Missing ... all 3 fields must be filled !" 
    } Else {
        $wpf.btnRun.IsEnabled = $true
        $CmdLine = "Update-E2016BuildDoc -DocFile ""$WordDocument"" -ExcelInputFile ""$ExcelInputs"" -Department ""$DepartmentAcronym"""
        If ($wpf.chkMonitoring.IsChecked){
            $CmdLine = $CmdLine + " -MonitorProcess"
        }
    }
    $wpf.txtCmdLine.Text = $CmdLine
}

Function OldStatusLabel ($Msg) {
    # Trick to enable a Label to update during work :
    # Follow with "Dispatcher.Invoke("Render",[action][scriptblobk]{})" or [action][scriptblock]::create({})
    $wpf.lblStatus.Content = $Msg
    $wpf.$FormName.Dispatcher.Invoke("Render",[action][scriptblock]{})
}

Function StatusLabel {
    [CmdletBinding()]
    Param(  [parameter(Position = 1)][string]$msg,
            [parameter(Position = 2)][string]$LabelObjectName = "lblStatus"
    )
    # Trick to enable a Label to update during work :
    # Follow with "Dispatcher.Invoke("Render",[action][scriptblobk]{})" or [action][scriptblock]::create({})
    $wpf.$LabelObjectName.Content = $Msg
    $wpf.$FormName.Dispatcher.Invoke("Render",[action][scriptblock]{})
}


Function Update-E2016BuildDoc {
    <#
    .SYNOPSIS
        Special script to read parameters in Excel to update Exchange 2016 Build document
        that is (c) Bernard Chouinard and Sam Drey

    .DESCRIPTION
        Special script to read parameters in Excel to update Exchange 2016 Build document
        that is (c) Bernard Chouinard and Sam Drey

    .PARAMETER DocFile
        Specifies the Word Exchange 2016 template file
        Note that the bookmarks will be checked by the script. If the wrong document is passed
        on this parameter, the script will inform you and stop.

    .PARAMETER ExcelInputFile
        Specifies the Excel Input file for the document to be updated with.

    .INPUTS
        None. You cannot pipe objects to that script.

    .OUTPUTS
        None for now

    .EXAMPLE
    .\Do-Something.ps1
    This will launch the script and do someting

    .EXAMPLE
    .\Do-Something.ps1 -CheckVersion
    This will dump the script name and current version like :
    SCRIPT NAME : Do-Something.ps1
    VERSION : v1.0

    .NOTES
    None

    .LINK
        https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

    .LINK
        https://github.com/SammyKrosoft
    #>
    [CmdLetBinding(DefaultParameterSetName = "NormalRun")]
    Param(
        [Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")][string]$DocFile,
        [Parameter(Mandatory = $False, Position = 2, ParameterSetName = "NormalRun")][string]$ExcelInputFile,
        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = "NormalRun")][string]$Department = "Dummy",
        [Parameter(Mandatory = $false, Position = 4, ParameterSetName = "NormalRun")][switch]$MonitorProcess,
        [Parameter(Mandatory = $false, Position = 4, ParameterSetName = "CheckOnly")][switch]$CheckVersion
    )

    <# ------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) ------- #>
    #Initializing a $Stopwatch variable to use to measure script execution
    $stopwatch = [system.diagnostics.stopwatch]::StartNew()
    #Using Write-Debug and playing with $DebugPreference -> "Continue" will output whatever you put on Write-Debug "Your text/values"
    # and "SilentlyContinue" will output nothing on Write-Debug "Your text/values"
    $DebugPreference = "Continue"
    # Set Error Action to your needs
    $ErrorActionPreference = "SilentlyContinue"
    #Script Version
    $ScriptVersion = "0.1"
    <# Version changes
    v0.1 : first script version
    v0.1 -> v0.5 : 
    #>
    $ScriptName = $MyInvocation.MyCommand.Name
    If ($CheckVersion) {Write-Host "SCRIPT NAME     : $ScriptName `nSCRIPT VERSION  : $ScriptVersion";exit}
    # Log or report file definition
    # NOTE: use $PSScriptRoot in Powershell 3.0 and later or use $LocalScriptExecPath = split-path -parent $MyInvocation.MyCommand.Definition in Powershell 2.0
    $LocalScriptExecPath = $PSScriptRoot
    $OutputReport = "$LocalScriptExecPath\$($ScriptName)_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
    # Other Option for Log or report file definition (use one of these)
    $ScriptLog = "$LocalScriptExecPath\$($ScriptName)-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
    <# ---------------------------- /SCRIPT_HEADER ---------------------------- #>

    #region Execution
    #region Prerequisites
    If ($MonitorProcess){
        invoke-expression 'cmd /c start powershell.exe -noprofile -Command { [console]::WindowWidth=90; [console]::WindowHeight=7; [console]::BufferWidth=[console]::WindowWidth;while($True){$WordProcess = get-process *Word*;cls;if ($WordProcess -ne $null){$WordProcess | out-host}Else{"No Wordprocess running..." | out-host};sleep 1}}'
        invoke-expression 'cmd /c start powershell.exe -noprofile -Command { [console]::WindowWidth=90; [console]::WindowHeight=7; [console]::BufferWidth=[console]::WindowWidth;while($True){$WordProcess = get-process *Excel*;cls;if ($WordProcess -ne $null){$WordProcess | out-host}Else{"No Excelprocess running..." | out-host};sleep 1}}'
    }

$TextFormFieldsList = @"
Partner_Nickname
Partner_FullName
EXCH_Source_Dir
E2016Extras_DIR
EXCH_INST_DIR
SMTP_Pri_Dom_1
Client_Endpoint
Internal_Url
External_Url
Autodiscover
E2016_Org_Unit
NIC_MAPI_NAme
NIC_MAPI_HW_NAme
NIC_REP_Name
NIC_REP_HW_Model
DEFAULT_GATEWAY
DNS1
DNS2
DOMAIN_NAME
EXTRAS_CD
FQDN_DOMAIN
CASARRAY
IP_ADDRESS
FIRST_SERVER_NAME
SUBNET_MASK
IPADDRESS1
SECOND_SERVER_NAME
SUBNET_MASK1
DB_First_Server
Dag_Name
Witness_Server
Witness_DIR
PageFile
Prod_Key
DB_Prefix
Cert_Name_3
Cert_Name_4
Cert_Name_5
Cert_Name_6
Cert_Name_7
"@
    
    #34 fields
    
    $FieldsArray = HereStringToArray $TextFormFieldsList
    
    #Loading Presentation Framework assembly for inputbox
    Add-Type -AssemblyName PresentationFramework

    #endregion
    #region Execution
    
    #region Parameters validation before continuing ...

    Title1 "Validating user input parameters"

    Write-Host "Path of the current script where we expect to find the documents (Word, Excel):" -f Yellow
    Write-Host $LocalScriptExecPath
    Write-Host "Word document passed in parameter:      " -n -f Yellow
    Write-host $WordDocument
    Write-Host "Excel document passed in parameters :   " -n -f yellow
    Write-Host $ExcelInputs
    Write-Host "The path  :                             " -n -f Yellow
    Write-host $LocalScriptExecPath\

    $msg = "Checking if you specified a -DocFile on the script"
    $Percent = 0
    Update-WPFProgressBarAndStatus $msg $percent

    if (-not $DocFile) {
        $DocName = "E2016BuildTest.docx"
        $DocPath = ".\"
        $Docfile = $DocPath + $DocName
        LogMag "No DocFile specified, trying to use $DocFile ..." -B yellow
    } Else {
        $DocName = Split-Path -Leaf -Path "$DocFile"
        $DocFile = $LocalScriptExecPath + "\" + $DocName
        LogGreen "Docfile specified: $DocName" -b blue
    }

    #endregion Parameters validation

    $msg = "Check for file existence"
    $Percent = 5
    Update-WPFProgressBarAndStatus $msg $percent

    $FileExists = Test-Path $DocFile

    If ($FileExists) {
        LogGreen "Word doc file exists, continuing..."
        LogBlue "Copying the master template as temp doc..."
        $TempDocFile = $LocalScriptExecPath + "\" + "TempE2016BuildDoc.docx"
        Copy-item $DocFile -Destination $TempDocFile
        LogBlue "Working copy becomes the Temporary file ..."
        $DocFile = $TempDocFile
    } Else {
        $msg = "The Word Doc file does not exist ... specify another one."
        LogMag $msg
        [System.Windows.MessageBox]::Show($msg)
        $wpf.$FormName.IsEnabled = $true
        StatusLabel "Ready !"
        Return #Trying to return to GUI...
    }

    $msg = "Creating new Word COM object"
    $Percent = 10
    Update-WPFProgressBarAndStatus $msg $percent

    $MSWord = New-Object -ComObject Word.Application

    <# INFO : ROUTINE TO END WORD PROCESS AND CLEAN THE COM OBJ AND THE VARIABLE
    $MSWord.Quit()
    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$MSword)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Remove-Variable MSword
    Exit
    #>

    $MSWord.Visible = $true

    $msg = "Opening document"
    $Percent = 15
    Update-WPFProgressBarAndStatus $msg $percent

    $Doc = $MSWord.Documents.Open($Docfile)

    if ($Doc) {
        LogMag "$DocFile opened successfully..."
    } else {
        $Msg = "$DocFile opening failed !!"
        $MSWord.Quit()
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$MSword)
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        Remove-Variable MSword
        LogMag $msg
        [System.Windows.MessageBox]::Show($msg)
        $wpf.$FormName.IsEnabled = $true
        StatusLabel "Ready !"
        Return #Trying to return to GUI...
    }

    
    $msg = "Word doc Text Form Field validation"
    $Percent = 25
    Update-WPFProgressBarAndStatus $msg $percent

    LogGreen "Loading all Document's Form Field..."
    $FormFields = $Doc.FormFields
    LogGreen "Initializing collection variable..."
    $WordDocFormFieldsCollection = @()
    LogMag "Beginning parsing all Doc formfield parsing..."
    Foreach ($FF in $FormFields){
        $FFType = Switch ($FF.Type) {
            70 {"Text"}
            Default {"Other"}
        }
        $FFPSObj = [PSCustomObject]@{
            "FormField Name"   =   $FF.Name
            "FormField Type"   =   $FFType
            "FormField Value"  =   $FF.Result
        }
        $WordDocFormFieldsCollection += $FFPSObj
    }
    LogMag "Doc text Form Fields parsing complete and saved in collection variable."

    $msg = "Comparing each Field to ckeck against Word Doc Fields"
    $Percent = 30
    Update-WPFProgressBarAndStatus $msg $percent

    $CompareArrays = Compare2Arrays -ReqFields $FieldsArray -FieldsToCompareToReqFields $WordDocFormFieldsCollection

    If (-not $CompareArrays){
        $msg = "Missing fields in the Word Doc ... exiting..."
        $Doc.Close()
        $MSWord.Quit()
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$MSword)
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        Remove-Variable MSword
        LogMag $msg
        [System.Windows.MessageBox]::Show($msg)
        $wpf.$FormName.IsEnabled = $true
        StatusLabel "Ready !"
        Return #Trying to return to GUI...
    }

    $msg = "Excel file opening and parsing..."
    $Percent = 35
    Update-WPFProgressBarAndStatus $msg $percent

    LogGreen "Launching Excel function for department $Department ..."
    LogGreen "Excel file path : $ExcelInputfile"
    
    $FormFieldsFromExcel = $null
    $FormFieldsFromExcel = Get-E2016ReportValues -Department $Department -ExcelInput $ExcelInputFile

    # Below couple of lines is for debug purposes if for some reasons your Excel call returns nothing...
    # logmag "FORM FIELDS = FALSE ?"
    # $FormFieldsFromExcel | out-host

    if ((-not $FormFieldsFromExcel) -or ($FormFieldsFromExcel -eq "NoTab") -or ($FormFieldsFromExcel -eq "FileNotFound")){
        If ($FormFieldsFromExcel -like "*NoTab*")
        {
            $msg = "No tab with $Department name in $ExcelInputFile or TAB not loaded !`nClick [Load names] button, or create a TAB named $Department in your file or specify the correct Excel input template`nand click the [Load names] button again..."
        } elseif ($FormFieldsFromExcel -like "*FileNotFound*"){
            $msg = "Excel file not found ..."
        }
        $Doc.Close()
        $MSWord.Quit()
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$MSword)
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        Remove-Variable MSword

        LogMag $msg
        [System.Windows.MessageBox]::Show($msg)
        $wpf.$FormName.IsEnabled = $true
        StatusLabel "Ready !"
        Return #Trying to return to GUI...
    }

    $FormFieldsFromExcel | Out-File

    $msg = "Like for the Word document, creating custom PS object with Excel values"
    $Percent = 50
    Update-WPFProgressBarAndStatus $msg $percent
    
    LogGreen "Loading all Document's Form Field..."
    LogGreen "Initializing collection variable..."
    $ExcelFormFieldsCollection = @()
    LogMag "Beginning parsing all Doc formfield parsing..."
    Foreach ($FF in $FormFieldsFromExcel){
        $FFPSObj = [PSCustomObject]@{
            "FormField Name"   =   $FF.BookMark
            "FormField Value"  =   $FF.Value
        }
        $ExcelFormFieldsCollection += $FFPSObj
    }
    LogMag "Excel input parsing complete and saved in collection variable."
    
        #Uncomment the below to check the Excel headers if needed to debug
        <# $ExcelFormFieldsCollection

        # $Doc.Close()
        # $MSWord.Quit()
        # $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$MSword)
        # [gc]::Collect()
        # [gc]::WaitForPendingFinalizers()
        # Remove-Variable MSword
        # exit
        #>

    $msg = "Comparing each Field to ckeck against Excel Doc Fields"
    $Percent = 60
    Update-WPFProgressBarAndStatus $msg $percent

    $CompareArrays = Compare2Arrays -ReqFields $FieldsArray -FieldsToCompareToReqFields $ExcelFormFieldsCollection

    If (-not $CompareArrays){
        $msg = "Missing fields on the Excel file ... exiting..."
        $Doc.Close()
        $MSWord.Quit()
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$MSword)
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        Remove-Variable MSword
        LogMag $msg
        [System.Windows.MessageBox]::Show($msg)
        $wpf.$FormName.IsEnabled = $true
        StatusLabel "Ready !"
        Return #Trying to return to GUI...
    }
   

    Write-Host "Total fields gotten from Excel $Department tab:" -b yellow -f blue
    Write-Host "$($FormFieldsFromExcel.count)" -b yellow -f blue

    Write-Host "Total fields from Word :" -BackgroundColor Blue
    Write-Host $FormFields.count -BackgroundColor Blue

    Write-Host 

    $msg = "Updating all fields in Word Document with the Excel input values..."
    $Percent = 70
    Update-WPFProgressBarAndStatus $msg $percent

    LogGreen $msg

    Foreach ($FF in $FormFieldsFromExcel){
        $Doc.FormFields($($FF.Bookmark)).TextInput.Default = $($FF.Value)
        $FF.Bookmark | out-host
    }

    $msg = "Fields default updated - now updating FormFields displayed values and cross references..."
    $Percent = 75
    Update-WPFProgressBarAndStatus $msg $percent

    #To update all fields :
    LogYellow $msg
    $MSWord.ActiveDocument.Fields.Update()

    $msg = "FormFields and cross references updated - now updating headers and footers for all sections..."
    $Percent = 80
    Update-WPFProgressBarAndStatus $msg $percent

    #To update header and footer
    # Iterate through Sections
    LogMag "There are $(($Doc.sections).count) sections in the word doc, updating..."
    foreach ($Section in $Doc.Sections)
    {
        # Update Header
        LogBlue "Updating Word sections Headers and Footers..."
        LogBlue $Section.

        $Header = $Section.Headers.Item(1)
        $Header.Range.Fields.Update()

        # Update Footer
        $Footer = $Section.Footers.Item(1)
        $Footer.Range.Fields.Update()
    }

    $msg = "Cross references updated - now updating tables of contents ..."
    $Percent = 85
    Update-WPFProgressBarAndStatus $msg $percent

    #Update table of contents ...
    LogMag "There are $(($Doc.TablesOfContents).count) TOCs in the word doc, updating..."
    foreach ($TOC in $Doc.TablesOfContents){
        $TOC.Update()
    }

    $msg = "Saving updated file, please wait..."
    $Percent = 95
    Update-WPFProgressBarAndStatus $msg $percent

    $OutputSubDir = $PSScriptRoot + "\Output"
    if(!(Test-Path -Path $OutputSubDir )){
        New-Item -ItemType directory -Path $OutputSubDir -Force
        Write-Host "New Output folder $OutputSubDir created..."
    }
    else
    {
      Write-Host "Folder Output already exists, using it."
    }

    $outputFile = $LocalScriptExecPath + "\Output\" + $Department + " - " + $(($DocName.split(".")[0])) + "_" + (Get-Date -Format "dd-MM-yyyy-HH-mm-ss") + ".docx"

    Try {
        $Doc.SaveAs([REF]$outputFile)
        LogYellow "File saved as $outputFile" -B Blue
        LogMag "Removing temporary file"
        Remove-Item $DocFile -Force
    }
    Catch {
        Write-Host "An error occured saving the file" -B red -F Yellow
        Logyellow $outputfile -b red
    }
    
    #$Doc.GoTo(1,1,1)

    $msg = "Processing complete - final user input before ending"
    $Percent = 100
    Update-WPFProgressBarAndStatus $msg $percent

    $Action = [System.Windows.MessageBox]::Show("Do you want to close the $($Doc.Name) Word doc ?","$($Doc.Name)",'YesNo','Warning')

    Switch ($Action){
        "Yes" {LogGreen "Closing the doc and closing Word..."
                $Doc.Close()
                $MSWord.Quit()
                $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$MSword)
                [gc]::Collect()
                [gc]::WaitForPendingFinalizers()
                Remove-Variable MSword
                }
        "No" {
                "Leaving the doc opened"
                $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$MSword)
                [gc]::Collect()
                [gc]::WaitForPendingFinalizers()
                Remove-Variable MSword
            }
    }
    $msg = "Processing complete - final user input before ending"
    Update-WPFProgressBarAndStatus $msg 0 "Ready for the next action."
 }

cls

Get-PowerShellVersion

If ($NoNeedToCheckMSWord){
    LogMag "Assuming you already have Word 2013 or later ... if script does not generate the proper document, please re-run without the -NoNeedToCheckMSWord switch !"
} Else {
    Get-MSWordVersion
}

#========================================================
#region WPF form definition and load controls
#========================================================
# Load a WPF GUI from a XAML file built with Visual Studio
Add-Type -AssemblyName presentationframework, presentationcore
$wpf = @{ }
# NOTE: Either load from a XAML file or paste the XAML file content in a "Here String"
#$inputXML = Get-Content -Path ".\WPFGUIinTenLines\MainWindow.xaml"
$inputXML = @"
<Window x:Name="DTUForm" x:Class="Console_for_Update_E2016Doc.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Console_for_Update_E2016Doc"
        mc:Ignorable="d"
        Title="Document Template Customizator" Height="707.492" Width="815.5" ResizeMode="NoResize">
    <Grid>
        <Rectangle Fill="#FFF4F4F5" HorizontalAlignment="Left" Height="133" Margin="11,293,0,0" Stroke="#FF828282" VerticalAlignment="Top" Width="781"/>
        <Rectangle Fill="#FFF4F4F5" HorizontalAlignment="Left" Height="114" Margin="11,174,0,0" Stroke="#FF828282" VerticalAlignment="Top" Width="781"/>
        <Rectangle Fill="#FFF4F4F5" HorizontalAlignment="Left" Height="114" Margin="11,55,0,0" Stroke="#FF828282" VerticalAlignment="Top" Width="781"/>
        <TextBox x:Name="txtDocFileName" HorizontalAlignment="Left" Height="55" Margin="101,103,0,0" TextWrapping="Wrap" Text="Exchange 2016 CUX Mailbox Server Build Document v2.7 - MS.docx" VerticalAlignment="Top" Width="681" FontSize="14"/>
        <Label x:Name="lblWordDocFileName" Content="Word Document Template File Name" HorizontalAlignment="Left" Margin="10,55,0,0" VerticalAlignment="Top" FontSize="14" FontWeight="Bold"/>
        <Label x:Name="lblMustBeInCurrentDir1" Content="(must be in current directory)" HorizontalAlignment="Left" Margin="9,73,0,0" VerticalAlignment="Top" Foreground="#FFF10A0A" FontWeight="Bold" Height="27" FontStyle="Italic"/>
        <TextBox x:Name="txtExcelFileName" HorizontalAlignment="Left" Height="55" Margin="101,221,0,0" TextWrapping="Wrap" Text="E2016BuildInputs - MS.xlsx" VerticalAlignment="Top" Width="681" FontSize="14"/>
        <Label x:Name="lblExcelDocFileName" Content="Excel file containing the inputs" HorizontalAlignment="Left" Margin="10,174,0,0" VerticalAlignment="Top" FontSize="14" FontWeight="Bold"/>
        <Label x:Name="lblMustBeInCurrentDir2" Content="(must be in current directory)" HorizontalAlignment="Left" Margin="10,192,0,0" VerticalAlignment="Top" Foreground="#FFF10A0A" FontWeight="Bold" Height="26" FontStyle="Italic"/>
        <CheckBox x:Name="chkMonitoring" Content="Monitoring Word and Excel Processes" HorizontalAlignment="Left" Margin="17,431,0,0" VerticalAlignment="Top" Height="26" FontSize="14" IsEnabled="False"/>
        <TextBox x:Name="txtCustomerAcronym" HorizontalAlignment="Left" Height="54" Margin="101,354,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="305" FontSize="14" IsEnabled="False"/>
        <Label x:Name="lblCustomerAcronym" Content="Customer / Department" HorizontalAlignment="Left" Margin="10,289,0,0" VerticalAlignment="Top" FontSize="14" FontWeight="Bold"/>
        <TextBox x:Name="txtCmdLine" HorizontalAlignment="Left" Height="60" Margin="10,457,0,0" TextWrapping="Wrap" Text="Command line..." VerticalAlignment="Top" Width="774" IsReadOnly="True" FontSize="14"/>
        <Button x:Name="btnRun" Content="RUN" HorizontalAlignment="Left" Margin="11,534,0,0" VerticalAlignment="Top" Width="204" Height="44" FontSize="14">
            <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>
        <Label x:Name="lblStatus" Content="Ready !" HorizontalAlignment="Left" Margin="90,586,0,0" VerticalAlignment="Top" Width="634" FontSize="14"/>
        <Button x:Name="btnResetFields" Content="Reset default fields" HorizontalAlignment="Left" Margin="561,534,0,0" VerticalAlignment="Top" Width="221" Height="44" FontSize="14">
            <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>
        <Button x:Name="btnExit" Content="EXIT" HorizontalAlignment="Left" Margin="285,534,0,0" VerticalAlignment="Top" Width="203" Height="44" FontSize="14">
            <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>
        <Label Content="Status:" HorizontalAlignment="Left" Margin="14,586,0,0" VerticalAlignment="Top" FontSize="14"/>
        <Label Content="Current Dir:" HorizontalAlignment="Left" Margin="10,5,0,0" VerticalAlignment="Top" Foreground="#FFA0A0A0" FontStyle="Italic" FontSize="11"/>
        <ListBox x:Name="lstCustomerAcronyms" HorizontalAlignment="Left" Height="108" VerticalAlignment="Top" Width="371" Margin="413,313,0,0"/>
        <Button x:Name="btnLoadAcronyms" Content="Load names" HorizontalAlignment="Left" Margin="17,354,0,0" VerticalAlignment="Top" Width="76" Height="54" Foreground="Black" Background="Teal">
            <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>
        <TextBlock x:Name="txtBlockCurrentDir" HorizontalAlignment="Left" Margin="81,10,0,0" TextWrapping="Wrap" Text="Current dir..." VerticalAlignment="Top" Height="40" Width="665" FontStyle="Italic" Foreground="#FFA0A0A0" Background="White" FontSize="11"/>
        <ProgressBar x:Name="ProgressBar" HorizontalAlignment="Left" Height="37" Margin="14,620,0,0" VerticalAlignment="Top" Width="768" Foreground="#FFC310BB">
            <ProgressBar.Effect>
                <DropShadowEffect/>
            </ProgressBar.Effect>
        </ProgressBar>
        <Button x:Name="btnBrowseExcel" Content="Browse" HorizontalAlignment="Left" Margin="16,221,0,0" VerticalAlignment="Top" Width="76" Height="54" Background="Teal">
            <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>
        <Button x:Name="btnBrowseWord" Content="Browse" HorizontalAlignment="Left" Margin="14,104,0,0" VerticalAlignment="Top" Width="75" Height="54" Background="Teal">
            <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>
        <TextBlock HorizontalAlignment="Left" Margin="17,313,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="295" Height="34" FontWeight="Bold" Foreground="Red" FontStyle="Italic"><Run Text="("/><Run Text="l"/><Run Text="oad the "/><Run Text="Customers/"/><Run Text="Department"/><Run Text="s"/><Run Text=" names "/><Run Text="with the &quot;Load Name&quot; button "/><Run Text="and select one)"/></TextBlock>
    </Grid>
</Window>
"@

$inputXMLClean = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''
[xml]$xaml = $inputXMLClean
$reader = New-Object System.Xml.XmlNodeReader $xaml
$tempform = [Windows.Markup.XamlReader]::Load($reader)
$namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
$namedNodes | ForEach-Object {$wpf.Add($_.Name, $tempform.FindName($_.Name))}

$FormName = $NamedNodes[0].Name
#========================================================
# END of WPF form definition and load controls
#endregion
#========================================================

#========================================================
#region WPF EVENTS definition
#========================================================
#region Buttons
$wpf.txtDocFileName.add_TextChanged({
    Update-Cmd
})

$wpf.txtExcelFileName.add_TextChanged({
    Update-Cmd
})

$wpf.txtCustomerAcronym.add_TextChanged({
    Update-Cmd
})

$wpf.chkMonitoring.add_Click({
    Update-Cmd
})

$wpf.btnRun.add_Click({
    $wpf.$FormName.IsEnabled = $false
    StatusLabel "Working..."
    $ExcelPathAndName = $PSScriptRoot + "\" + $wpf.txtExcelFileName.Text
    $WordPathAndName = $PSScriptRoot + "\" + $wpf.txtDocFileName.Text
    $WordOk = Test-Path -Path $WordPathAndName
    $ExcelOk = Test-Path -Path $ExcelPathAndName
    $MissingFile = @()
    $MissingFilesFlag = $false
    If (-not $WordOk){$MissingFile += "`n- Word document: $($wpf.txtDocFileName.Text)";$MissingFilesFlag = $true}
    If (-not $ExcelOk){$MissingFile += "`n- Excel document: $($wpf.txtExcelFileName.Text)";$MissinFile = $true}
    $MissingFile = $MissingFile -join ""
    If ($MissingFile){
        $msg = "Missing the following: $MissingFile `n`nPlease specify a proper file existing in $PSScriptRoot"
        LogMag $msg
        [System.Windows.MessageBox]::Show($msg,"Missing File(s)","OK","Error")
        $wpf.$FormName.IsEnabled = $true
        StatusLabel "Ready !"
    } Else {
        $CommandToInvoke = $wpf.txtCmdLine.Text
        Invoke-Expression $CommandToInvoke
        $wpf.$FormName.IsEnabled = $true
        $wpf.ProgressBar.Value = 0
        StatusLabel "Ready !"
    }
})

$wpf.btnExit.add_Click({
    $wpf.$FormName.close()
})

$wpf.btnResetFields.add_Click({
    $wpf.txtDocFileName.Text = "Exchange 2016 CUX Mailbox Server Build Document v2.7 - MS.docx"
    $wpf.txtExcelFileName.Text = "E2016BuildInputs - MS.xlsx"
    $wpf.txtCustomerAcronym.Text = ""
})

$wpf.btnLoadAcronyms.add_Click({
    $wpf.$FormName.IsEnabled = $false
    StatusLabel "Please wait while loading acronyms from Excel..."
    $custAcronyms = Get-ExcelWorkSheetsNamesWPFGUIUpdate -ExcelInput $($wpf.txtExcelFileName.Text)
    $wpf.lstCustomerAcronyms.ItemsSource = $custAcronyms
    $wpf.lstCustomerAcronyms.SelectedIndex = 0
    $wpf.txtCustomerAcronym.Text = $wpf.lstCustomerAcronyms.SelectedItem
    $wpf.$FormName.IsEnabled = $true

    Update-WPFProgressBarAndStatus "Ready !" 0 "Waiting for action."
})

$wpf.btnBrowseExcel.add_Click({
    $OpenFileDialog = New-Object Microsoft.Win32.OpenFileDialog
    $OpenFileDialog.FileName = $wpf.txtExcelFileName.Text   # Default file name
    $OpenFileDialog.DefaultExt = ".xlsx"                    # Default file extension
    $OpenFileDialog.Filter = "Excel files (.xlsx)|*.xlsx"   # Filter files by extension
    $OpenFileDialog.InitialDirectory = $PSScriptRoot        # Default directory
    $Result = $OpenFileDialog.ShowDialog()
    if ($Result){
        $FileName = $OpenFileDialog.FileName
        $SimpleFileName = Split-Path -Leaf -Path $FileName
        $wpf.txtExcelFileName.text = $SimpleFileName
    }
})

$wpf.btnBrowseWord.add_Click({
    $OpenFileDialog = New-Object Microsoft.Win32.OpenFileDialog
    $OpenFileDialog.FileName = $wpf.txtDocFileName.Text
    $OpenFileDialog.DefaultExt = ".docx"
    $OpenFileDialog.Filter = "MS Word files (.docx)|*.docx"
    $OpenFileDialog.InitialDirectory = $PSScriptRoot
    $Result = $OpenFileDialog.ShowDialog()
    if ($Result) {
        $FileName = $OpenFileDialog.FileName
        $SimpleFileName = Split-Path -Leaf -Path $FileName
        $wpf.txtDocFileName.text = $SimpleFileName
    }
})

$wpf.lstCustomerAcronyms.add_SelectionChanged({
    $wpf.txtCustomerAcronym.Text = $wpf.lstCustomerAcronyms.SelectedItem
})

#endregion
# End of Buttons region

#region Load, Draw (render) and closing form events
#Things to load when the WPF form is loaded aka in memory
$wpf.$FormName.Add_Loaded({
    Update-Cmd
    $wpf.txtBlockCurrentDir.Text = $PSScriptRoot
    $wpf.$FormName.Title = $wpf.$FormName.Title + (" - v$ScriptVersion")
 
})
#Things to load when the WPF form is rendered aka drawn on screen
$wpf.$FormName.Add_ContentRendered({
    #Update-Cmd
})
$wpf.$FormName.add_Closing({
    $msg = "bye bye !"
    write-host $msg
})

#endregion
# End of load, draw and closing form events region

#endregion
# End of region WPF EVENTS definition
#========================================================

#endregion

#=======================================================
#End of Events from the WPF form
#endregion
#=======================================================

#========================================================
# END of WPF EVENTS definition
#endregion
#========================================================


# Load the form:
# Older way >>>>> $wpf.MyFormName.ShowDialog() | Out-Null >>>>> generates crash if run multiple times
# Newer way >>>>> avoiding crashes after a couple of launches in PowerShell...
# Using method from https://gist.github.com/altrive/6227237 to avoid crashing Powershell after we re-run the script after some inactivity time or if we run it several times consecutively...
$async = $wpf.$FormName.Dispatcher.InvokeAsync({
    $wpf.$FormName.ShowDialog() | Out-Null
})
$async.Wait() | Out-Null