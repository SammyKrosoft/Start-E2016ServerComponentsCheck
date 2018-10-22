Function Check-E2016ComponentStateToActive {
    <#
    .NOTES
    Based on V1.1 08.06.2014  by Adnan Rafique @ExchangeITPro
    Modified by Samuel Drey @Microsoft
    V1 10.OCT.2018
    .SYNOPSIS
    Bring Exchange components to active state.
    .DESCRIPTION
    Bring Exchange components to active state.
    .PARAMETER HybridServer
    Indicates to check 2 additional Server Components that are important for
    Office 365 synchronization between the On-premises environment and the
    Exchange Online environment : "ForwardSyncDaemon" and "ProvisioningRps".
    .PARAMETER CheckOnly
    Indicated the script to only check which Components are inactive before
    attempting anything.
    .EXAMPLE
    .\Start-E2016ServerComponentStateToActive.ps1 -HybridServer -CheckOnly
    Will check all Server Components, including ForwardSyncDaemon and ProvisioningRps
    components, but won't attempt to start these.
    .EXAMPLE
    .\Start-E2016ServerComponentStateToActive.ps1
    Will check all Server Components, excluding the ForwardSyncDaemon and ProvisioningRps,
    and attempt to start these.
    The script will tell you if the operation was successful or not.
    .EXAMPLE
    .\Start-E2016ServerComponentStateToActive.ps1 -HybridServer
    Will check and try to start all Server Components, including ForwardSyncDaemon and ProvisioningRps
    .EXAMPLE
    .\Start-E2016ServerComponentStateToActive.ps1 -CheckOnly
    Will check all Server Components, excluding ForwardSyncDaemon and ProvisioningRps
    components, but won't attept to start these.
    .LINK
    https://blogs.technet.microsoft.com/exchange/2012/09/21/lessons-from-the-datacenter-managed-availability/
    #>

    #Requires -version 3.0

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)][switch]$HybridServer,
        [Parameter(Mandatory = $false)][switch]$CheckOnly
    )

    Function Title1 ([string]$title, $TotalLength = 100, $Back = "Yellow", $Fore = "Black") {
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

    Function Test-ExchTools(){
        <#
        .SYNOPSIS
        This small function will just check if you have Exchange tools installed or available on the
        current PowerShell session.

        .DESCRIPTION
        The presence of Exchange tools are checked by trying to execute "Get-ExBanner", one of the basic Exchange
        cmdlets that runs when the Exchange Management Shell is called.

        Just use Test-ExchTools in your script to make the script exit if not launched from an Exchange
        tools PowerShell session...

        .EXAMPLE
        Test-ExchTools
        => will exit the script/program si Exchange tools are not installed
        #>
        Try
        {
            Get-command Get-MAilbox -ErrorAction Stop
            $ExchInstalledStatus = $true
            $Message = "Exchange tools are present !"
            Write-Host $Message -ForegroundColor Blue -BackgroundColor Red
        }
        Catch [System.SystemException]
        {
            $ExchInstalledStatus = $false
            $Message = "Exchange Tools are not present ! This script/tool need these. Exiting..."
            Write-Host $Message -ForegroundColor red -BackgroundColor Blue
            Exit
        }
        Return $ExchInstalledStatus
    }

    If (!(Test-ExchTools)){exit}

    cls

    If ($CheckOnly) {
        Title1 "Check only specified - will just list inactive components without trying to activate ..."
    } Else {
        Title1 "CheckOnly NOT specified ... will try to activate everything if more than 2 components are inactive..."
    }

    $E2016NamesList = @()
    $E2016 = Get-ExchangeServer | ? {$_.AdminDisplayVersion -match "15.1"} 
    Foreach ($item in $E2016){$E2016NamesList += $($item.Name)}

    $counter = 0
    Foreach ($Server in $E2016){
        Title1 $Server
        write-progress -id 1 -Activity "Activating all components" -Status "Server $Server" -PercentComplete $($Counter/$($E2016.Count)*100)
        $Counter++

        #Get the status of components
        If (!($HybridServer)){
            Write-Host "You didn't specify the -HybridServer switch, meaning that this is an On-Premises only environment (aka not Hybrid, not synchronizing with the cloud). We don't need ForwardSyncDaemon and ProvisioningRPS Components - leaving these as-is"
            $ComponentStateStatus = Get-ServerComponentState ($Server.Name) | ? {$_.Component -ne "ForwardSyncDaemon" -and $_.Component -ne "ProvisioningRps"}
        } Else {
            Write-Host "You specified the -HybridServer parameter, indicating that this is an On-Premises environment syncinc with O365. All Server Components need to be active..."
            $ComponentStateStatus = Get-ServerComponentState ($Server.Name) 
        }

        #$ComponentStateStatus | ft Component,State -Autosize
        $InactiveComponents = $ComponentStateStatus | ? {$_.State -eq "Inactive"}
        $ACtiveComponents = $ComponentStateStatus | ? {$_.State -eq "Active"}
        
        $NbActiveComponents = $ACtiveComponents.Count
        If ($NbActiveComponents -eq $null){$NbActiveComponents = 0}
        $NbInactiveComponents = $InactiveComponents.Count
        If ($NbInactiveComponents -eq $null){$NbInactiveComponents = 0}

        Write-Host "There are $NbActiveComponents active components, and $NbInactiveComponents inactive components on server $($Server.Name)" -BackgroundColor yellow -ForegroundColor red

        If ($NbInactiveComponents -eq 0){
            Write-Host "There are no inactive components, everything looks good ... "
            Continue
        } Else {
            Write-host "Some components are not active - we have $NbInactiveComponents inactive components..."
            $InactiveComponents | ft Component
            If (!($CheckOnly)){
                Write-host "... trying to re-activate all inactive components..." 
                $Counter1 = 0
                Foreach ($Component in $InactiveComponents) {
                    Write-progress -id 2 -ParentId 1 -Activity "Setting component states" -Status "setting $($Component.Component)..." -PercentComplete ($Counter1/$NbInactiveComponents*100)
                    $Command = "Set-ServerComponentState $($Server.Name) -Component $($Component.Component) -State Active -Requester Functional" 
                    Write-host "Running the following command: `n$Command" -BackgroundColor Blue -ForegroundColor White
                    Invoke-Expression $Command
                    $Counter1++
                }
                #Get the new status of components
                If (!($HybridServer)){
                    Write-Host "You didn't specify the -HybridServer switch, meaning that this is an On-Premises only environment (aka not Hybrid, not synchronizing with the cloud). We don't need ForwardSyncDaemon and ProvisioningRPS Components - leaving these as-is"
                    $ComponentStateStatus = Get-ServerComponentState ($Server.Name) | ? {$_.Component -ne "ForwardSyncDaemon" -and $_.Component -ne "ProvisioningRps"}
                } Else {
                    Write-Host "You specified the -HybridServer parameter, indicating that this is an On-Premises environment syncinc with O365. All Server Components need to be active..."
                    $ComponentStateStatus = Get-ServerComponentState ($Server.Name) 
                }
            
                #$ComponentStateStatus | ft Component,State -Autosize
                $InactiveComponents = $ComponentStateStatus | ? {$_.State -eq "Inactive"}
                $ACtiveComponents = $ComponentStateStatus | ? {$_.State -eq "Active"}
                
                $NbActiveComponents = $ACtiveComponents.Count
                If ($NbActiveComponents -eq $null){$NbActiveComponents = 0}
                $NbInactiveComponents = $InactiveComponents.Count
                If ($NbInactiveComponents -eq $null){$NbInactiveComponents = 0}

                Write-Host "There are now $NbActiveComponents active components, and $NbInactiveComponents inactive components"
                If ($NbInactiveComponents -eq 0) {Write-Host "$Server is now completely out of maintenance mode and component are active and functional." -ForegroundColor Yellow} Else {Write-host "There are still some inactive components ... please troubleshoot !" -BackgroundColor Red -ForegroundColor Yellow}
            
            } Else {
                Write-Host "Checking only... here's your list of inactive components:"
                $InactiveComponents | ft Component
            }
        }

    }

    write-progress -id 1 -Activity "Activating all components" -Status "All done !" -PercentComplete $($Counter/$($E2016.Count)*100)
    sleep 1
}

Function Update-cmd {
    
}

# Load a WPF GUI from a XAML file build with Visual Studio
Add-Type -AssemblyName presentationframework, presentationcore
$wpf = @{ }
# NOTE: Either load from a XAML file or paste the XAML file content in a "Here String"
#$inputXML = Get-Content -Path ".\WPFGUIinTenLines\MainWindow.xaml"
$inputXML = @"
<Window x:Name="frmCheckServerComponents" x:Class="Check_E2016ServerComponents.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Check_E2016ServerComponents"
        mc:Ignorable="d"
        Title="Exchange Server Components Checker" Height="450" Width="800">
    <Grid>
        <ComboBox x:Name="comboSelectExchangeVersion" HorizontalAlignment="Left" Margin="10,124,0,0" VerticalAlignment="Top" Width="120" SelectedIndex="1" IsReadOnly="True">
            <ComboBoxItem Content="Exchange 2013"/>
            <ComboBoxItem Content="Exchange 2016"/>
        </ComboBox>
        <CheckBox x:Name="chkCheckOnly" Content="CheckOnly" HorizontalAlignment="Left" Margin="10,151,0,0" VerticalAlignment="Top" IsChecked="True"/>
        <TextBox HorizontalAlignment="Left" Height="79" Margin="144,10,0,0" TextWrapping="Wrap" Text="Exchange 2013/2016 Server Component Checker" VerticalAlignment="Top" Width="518" TextAlignment="Center" VerticalContentAlignment="Center" FontSize="20" FontWeight="Bold" IsReadOnly="True">
            <TextBox.Effect>
                <DropShadowEffect ShadowDepth="10" Color="#FFACD151"/>
            </TextBox.Effect>
        </TextBox>
        <Button x:Name="btnRun" Content="Run" HorizontalAlignment="Left" Margin="10,380,0,0" VerticalAlignment="Top" Width="75"/>
        <Button x:Name="btnQuit" Content="Quit" HorizontalAlignment="Left" Margin="681,380,0,0" VerticalAlignment="Top" Width="75"/>
        <Label Content="List of Exchange components and their state" HorizontalAlignment="Left" Margin="251,168,0,0" VerticalAlignment="Top" Width="246"/>
        <ListView x:Name="ListView" HorizontalAlignment="Left" Height="147" Margin="10,199,0,0" VerticalAlignment="Top" Width="762">
            <ListView.View>
                <GridView>
                    <GridViewColumn/>
                </GridView>
            </ListView.View>
        </ListView>
        <CheckBox Content="HybridServer" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,171,0,0"/>
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
    $msg = "bye bye !"
    write-host $msg
})

#endregion Load, Draw and closing form events
#End of load, draw and closing form events

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