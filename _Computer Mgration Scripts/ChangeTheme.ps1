#--------------------------------------------------------------------------------- 
#The script is not supported under any Microsoft standard support 
#program or service. The script is provided AS IS without warranty  
#of any kind. Microsoft further disclaims all implied warranties including,  
#without limitation, any implied warranties of merchantability or of fitness for 
#a particular purpose. The entire risk arising out of the use or performance of  
#the sample scripts and documentation remains with you. In no event shall 
#Microsoft, its authors, or anyone else involved in the creation, production, or 
#delivery of the scripts be liable for any damages whatsoever (including, 
#without limitation, damages for loss of business profits, business interruption, 
#loss of business information, or other pecuniary loss) arising out of the use 
#of or inability to use the sample scripts or documentation, even if Microsoft 
#has been advised of the possibility of such damages 
#--------------------------------------------------------------------------------- 

#requires -Version 3.0

Function Change-Theme
{
<#
 	.SYNOPSIS
        Change-Theme is an advanced function which can be used to change the color scheme, wallpaper and startmenu background on Windows 8.
        It also allows you to save the current configuration and to load settings from a saved Config.csv file.
    .DESCRIPTION
    .PARAMETER  <Color>
        Changes the color scheme
    .PARAMETER <MenuBackground>
        Changes the background on the startmenu
    .PARAMETER <Background>
        Changes the desktopbackground
    .PARAMETER <SaveConfiguration>
        Saves the current configuration to C:\ThemeConfiguration\Config.csv
    .PARAMETER <LoadConfiguration>
        Loads configuration from Config.csv

    .EXAMPLE
        C:\PS> Change-Theme
		
		This command will change the color, menu background, the wallpaper and ask if you want to save the configuration
    .EXAMPLE
        C:\PS> Change-Theme -Color
		
		This command will change the color scheme and ask if you want to save the configuration
    .EXAMPLE
        C:\PS> Change-Theme -MenuBackground

        This command will change the background on the start menu and ask if you want to save the configuration
    .EXAMPLE
        C:\PS> Change-Theme -Background

        This command will change the wallpaper and ask if you want to save the configuration
    .EXAMPLE
        C:\PS> Change-Theme -SaveConfiguration

        This command will save the current configuration to C:\ThemeConfiguration\Config.csv
    .EXAMPLE
        C:\PS> Change-Theme -LoadConfiguration

        This command will allow you to load a Config.csv file and apply the values stored

#>

    [CmdletBinding(SupportsShouldProcess,DefaultParameterSetName="__AllParameterSets")]
    Param
    (
        [Parameter(Position=0,Mandatory,ParameterSetName="ChangeColor")]
        [Alias('changecolor')][Switch]$Color,
        [Parameter(Position=1,Mandatory,ParameterSetName="ChangeMenuBackground")]
        [Alias('changemenubackground')][Switch]$MenuBackground,
        [Parameter(Position=2,Mandatory,ParameterSetName="ChangeBackground")]
        [Alias('changebackground')][Switch]$Background,
        [Parameter(Position=3,Mandatory,ParameterSetName="Save")]
        [Alias('save')][Switch]$SaveConfiguration,
        [Parameter(Position=4,Mandatory,ParameterSetName="Load")]
        [Alias('load')][Switch]$LoadConfiguration
    )

    Begin
    {
        $Shell = New-Object -ComObject Shell.Application
	    $Desktop = $Shell.NameSpace(0X0)
        $WshShell = New-Object -comObject WScript.Shell
    }

    Process
    {
        If($Color)
        {
            Color
            SaveConfiguration
        }
        ElseIf($MenuBackground)
        {
            MenuBackground
            SaveConfiguration
        }
        ElseIf($Background)
        {
            Background
            SaveConfiguration
        }
        ElseIf($SaveConfiguration)
        {
            SaveConfiguration
        }
        ElseIf($LoadConfiguration)
        {
            LoadConfiguration
        }
        Else
        {
            Color
            MenuBackground
            Background
            SaveConfiguration
        }
    }
}

Function Color
{
    Write-Host "Enter value (0-24): "
    $ColorNumberDec = Read-Host

    #Write to registry
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Accent' -Name ColorSet_Version3 -PropertyType Dword -Value $($ColorNumberDec) -Force

    Write-Host "Change applied successfully, please sign out and back in for changes to take effect" -ForegroundColor Green
}

Function MenuBackground
{
    Write-Host "Enter value (100-119): "
    $BackgroundNumberDec = Read-Host

    #Write to registry
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Accent' -Name AccentId_v8.00 -PropertyType Dword -Value $($BackgroundNumberDec) -Force

    Write-Host "Change applied successfully, please sign out and back in for changes to take effect" -ForegroundColor Green
}

Function Background
{
    Write-Host "Enter path to wallpaper (Example C:\Background\MyBackground.jpg): "
    [string]$Wallpaper = Read-Host

    #Write to registry
    New-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name Wallpaper -PropertyType String -Value $($Wallpaper) -Force

    Write-Host "Change applied successfully, please sign out and back in for changes to take effect" -ForegroundColor Green
}

Function SaveConfiguration
{
    Write-Host "Do you want to save configuration to file?" -ForegroundColor White
    Write-Host "Y=Yes" -ForegroundColor Green -NoNewline
    Write-Host "N=No " -ForegroundColor Red -NoNewline
    [string]$answer = Read-Host

    If($($answer) -eq 'Y')
    {
        #Create folder and Configurationfile
        New-Item C:\ThemeConfiguration -ItemType directory -Force
        New-Item C:\ThemeConfiguration\Config.csv -ItemType file -Force
        [string]$DefaultText = "ColorNumber,MenuBackground,Wallpaper"
        $DefaultText >> C:\ThemeConfiguration\Config.csv

        #Get values from registry
        $ColorNumberValue = Get-ChildItem -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\' | ForEach-Object {Get-ItemProperty $_.pspath} | Where-Object {$_.ColorSet_Version3} | ForEach-Object {$_.ColorSet_Version3}
        $MenuBackgroundValue = Get-ChildItem -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\' | ForEach-Object {Get-ItemProperty $_.pspath} | Where-Object {$_.'AccentId_v8.00'} | ForEach-Object {$_.'AccentId_v8.00'}
        [string]$WallpaperValue = Get-ChildItem -Path 'HKCU:\Control Panel\' | ForEach-Object {Get-ItemProperty $_.pspath} | Where-Object {$_.Wallpaper} | ForEach-Object {$_.Wallpaper}

        #Write values to Configuration.txt
        $OutPut = $($ColorNumberValue) , $($MenuBackgroundValue) , $($WallpaperValue)
        $OutPut -join "," >> C:\ThemeConfiguration\Config.csv

        Write-Host "Configuration.csv has been created in C:\ThemeConfiguration\" -ForegroundColor Green
    }
    Else
    {
        Write-Host "Configuration not saved" -ForegroundColor DarkRed
    }
}

Function LoadConfiguration
{
    #Get location of Config.csv
    Write-Host "Enter location osf Config.csv (Example C:\ThemeConfiguration\) "
    [string]$Location = Read-Host
    [string]$path = $($Location) + "Config.csv"

    #Load Config.csv
    $ConfigCsv = Import-Csv -Path $($path)

    #Get values from Config.csv
    foreach($line in $ConfigCsv)
    {
        $ColorNumberLoad = $($line.ColorNumber)
        $MenuBackgroundLoad = $($line.MenuBackground)
        $WallpaperLoad = $($line.Wallpaper)
    }

    #Write to registry
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Accent' -Name ColorSet_Version3 -PropertyType Dword -Value $($ColorNumberLoad) -Force
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Accent' -Name AccentId_v8.00 -PropertyType Dword -Value $($MenuBackgroundLoad) -Force
    New-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name Wallpaper -PropertyType String -Value $($WallpaperLoad) -Force
}