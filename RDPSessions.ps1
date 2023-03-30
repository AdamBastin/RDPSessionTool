#region ADMINCHECK
    #Checks if script was ran with admin priveleges
    #If not, will relaunch with UAC prompt
    #Can be removed if adequate permissions are applied to non-admin account

    param([switch]$Elevated)
    function Test-Admin {
        $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
        $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    }
    
    if ((Test-Admin) -eq $false)  {
        if ($elevated) {
            # tried to elevate, did not work, aborting
        } else {
            #Can use pwsh.exe instead to use powershell 7
            try{Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))}
            catch{Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))}
        }
        exit
    }
#endregion

#region XAML
$inputXML=@"
<Window x:Class="wpfGUI.Window1"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:wpfGUI"
        mc:Ignorable="d"
        Title="RDP Sessions" Height="450" Width="505">
    <Grid>
        <ListView x:Name="lstvDisplay" Margin="0,113,0,72">
            <ListView.View>
                <GridView>
                    <GridViewColumn DisplayMemberBinding="{Binding Path=Server}" Header="Server" Width="200"/>
                    <GridViewColumn DisplayMemberBinding="{Binding Path=Username}" Header="User Name" Width="150"/>
                    <GridViewColumn DisplayMemberBinding="{Binding Path=State}" Header="State" Width="150"/>
                </GridView>
            </ListView.View>
        </ListView>
        <TextBox x:Name="txtComputer" HorizontalAlignment="Left" Margin="113,70,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <CheckBox x:Name="chkActive" Content="Show Active RDP Sessions" HorizontalAlignment="Left" Margin="10,14,0,0" VerticalAlignment="Top"/>
        <Label Content="Computer Name:" HorizontalAlignment="Left" Margin="10,65,0,0" VerticalAlignment="Top"/>
        <Button x:Name="btnSearchAll" Content="Search ALL Servers" HorizontalAlignment="Left" Margin="324,64,0,0" VerticalAlignment="Top" Height="21" Width="152"/>
        <Button x:Name="btnLogoff" Content="Log User off" HorizontalAlignment="Left" Margin="41,0,0,18" VerticalAlignment="Bottom" Height="36" Width="150" FontSize="20"/>
        <CheckBox x:Name="chkServers" Content="Only Search Servers" HorizontalAlignment="Left" Margin="329,44,0,0" VerticalAlignment="Top" IsChecked="True"/>

    </Grid>
</Window>
"@
#endregion

#region XAML reading
$global:ReadmeDisplay = $true
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
 
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{
    $Form=[Windows.Markup.XamlReader]::Load( $reader )
}
catch{
    Write-Warning "Unable to parse XML, with error: $($Error[0])`n Ensure that there are NO SelectionChanged or TextChanged properties in your textboxes (PowerShell cannot process them)"
    throw
}
$xaml.SelectNodes("//*[@Name]") | Foreach-Object{
    try {Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop}
    catch{throw}
}
#endregion

#Gets a list of all servers from AD, based on operating system
$global:servers = Get-ADComputer -Filter "operatingsystem -like '*server*'" | Sort-Object Name

$global:searching = $false

#Allows search for servers
$wpftxtComputer.add_TextChanged({
    $global:searching = $true
    $wpflstvDisplay.items.clear()
    $name = $wpftxtComputer.text
    $computers= Get-ADComputer -Filter "Name -like '*$name*'" -Properties Name -ErrorAction SilentlyContinue
    if(($computers.length -lt 100 -and $computers.length -gt -1) -or ($computers.name -ne '')){
        foreach($computer in $computers){
            $server = [PSCustomObject]@{
                Server = $computer.name
            }
            $wpflstvDisplay.items.add($server)
        }
    }    
})

[System.Collections.ArrayList]$global:sessions=@()

#Allows selection of server from listbox
$wpflstvDisplay.add_SelectionChanged({
    if ($global:searching -and $wpflstvDisplay.selectedindex -gt -1){
        $name = $wpflstvDisplay.selecteditem.server
        $global:sessions =@()
        $global:sessions += Get-RDPSessions(Get-ADComputer -Filter "Name -eq '$name'")
        $wpflstvDisplay.items.clear()
        for ($i=0;$i -lt $global:sessions.Count; $i++){
            $wpflstvDisplay.items.add($global:sessions[$i])
        }
        $global:searching = $false
    }
})

#Toggles between searching either SERVERS or ALL COMPUTERS
$WPFchkServers.add_Click({
    if($WPFchkServers.ischecked){
        $WPFbtnSearchAll.Content = "Search ALL Servers"
    }else{
        $WPFbtnSearchAll.Content = "Search ALL Computers"
    }
})

#Searches all SERVER or all COMPUTERS based on $WPFchkServers checkbox
$WPFbtnSearchAll.add_Click({
    if ($WPFchkServers.IsChecked){
        $global:sessions = Get-RDPSessions($global:servers)
    }else{
        $computers = Get-ADComputer -Filter *    
        $global:sessions = Get-RDPSessions($computers)
    }
    $wpflstvDisplay.items.clear()
    for ($i=0;$i -lt $global:sessions.Count; $i++){
        $wpflstvDisplay.items.add($global:sessions[$i])
    }

    #Creates a toast notification when done searching 
    [reflection.assembly]::loadwithpartialname('System.Windows.Forms')
    [reflection.assembly]::loadwithpartialname('System.Drawing')
    $notify = new-object system.windows.forms.notifyicon
    $notify.icon = [System.Drawing.SystemIcons]::Information
    $notify.visible = $true
    $notify.showballoontip(10,'Finished','Finished searching!',[system.windows.forms.tooltipicon]::None)
    start-sleep -s 1
    $notify.Dispose()
})

function Get-RDPSessions($computers){
    $allsessions =@()
    
    #Cycles each computer specified and tests the connection
    #If computer is online, will query all RDP Sessions
    foreach($computer in $computers){    
        if(Test-Connection $computer.Name -Quiet -count 1){
            $name = $computer.Name
            $sessions = query session /server:$name        
            if ($sessions.count -gt 0){
                Write-Host $name
                foreach($session in $sessions){
                    $row = $session -split "`n"    
                    if ($WPFchkActive.ischecked){
                        $regex = "Disc|Active" 
                    }else{
                        $regex = "Disc"
                    }
                    if ($row -NotMatch "services|console" -and $row -match $regex) {
                        $session = $($row -Replace ' {2,}', ',').split(',')       
                        $obj = [PSCustomObject]@{
                            Server = $name
                            Username = $session[1]
                            ID=$session[2]
                            State=$session[3]
                        }
                        $allsessions += $obj
                    }
                }            
            }
        }
    }
    return $allsessions
}

$WPFbtnLogoff.add_Click({
    #After search completes, populates listbox with all current sessions
    #Selecting a user allows to log off and removes from listbox
    if (($global:searching)-or ($wpflstvDisplay.SelectedIndex -lt 0)){
        Write-Warning "Select a session first"
        return
    }
    $id = $global:sessions[$wpflstvDisplay.selectedindex].id
    $server =  $global:sessions[$wpflstvDisplay.selectedindex].server
    logoff $id /server:$server
    Write-Host 'Logged off '$global:sessions[$wpflstvDisplay.selectedindex].Username'from '$global:sessions[$wpflstvDisplay.selectedindex].server
    $global:sessions.remove($global:sessions[$wpflstvDisplay.selectedindex])
    $wpflstvDisplay.items.remove($wpflstvDisplay.selecteditem)   
})


$form.ShowDialog() |Out-Null