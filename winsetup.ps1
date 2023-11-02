$wshell = New-Object -ComObject Wscript.Shell

#PYTHON
if ($wshell.Popup("Do you want to install Python3.12 ?",0,"Installer",4+32) -eq 6) {
    Start-Job -ScriptBlock { winget install Python.Python.3.12 --override "/passive InstallAllUsers=1 PrependPath=1 Include_test=0" }
}
# PowerToys
if ($wshell.Popup("Do you want to install Powertoys ?",0,"Installer",4+32) -eq 6) {
    Start-Job -ScriptBlock { winget install XP89DCGQ3K6VLD -h --accept-package-agreements ----disable-interactivity }
}
# Google Chrome
if ($wshell.Popup("Do you want to install GOOGLE CHROME ?",0,"Installer",4+32) -eq 6) {
    Start-Job -ScriptBlock { winget install Google.Chrome -h --accept-package-agreements ----disable-interactivity }
}
# Mobatek.MobaXterm
if ($wshell.Popup("Do you want to install MOBAXTERM ?",0,"Installer",4+32) -eq 6) {
    Start-Job -ScriptBlock { install Mobatek.MobaXterm -h --accept-package-agreements ----disable-interactivity }
}
# Whatsapp
if ($wshell.Popup("Do you want to install WhatsApp Desktop ?",0,"Installer",4+32) -eq 6) {
    Start-Job -ScriptBlock { winget install 9NKSQGP7F2NH -h --accept-package-agreements ----disable-interactivity }
}
# notepad++
if ($wshell.Popup("Do you want to install Notepad++ ?",0,"Installer",4+32) -eq 6) {
    Start-Job -ScriptBlock { winget install notepadd++.notepad++ -h --accept-package-agreements ----disable-interactivity }
}
#ADOBE ACROBAT
if ($wshell.Popup("Do you want to install ADOBE ACROBAT ?",0,"Installer",4+32) -eq 6) {
    Start-Job -ScriptBlock { winget install Adobe.Acrobat.Reader.64-bit -h --accept-package-agreements ----disable-interactivity }
}
#VLC MEDIA PLAYER
if ($wshell.Popup("Do you want to install VLC MEDIA PLAYER ?",0,"Installer",4+32) -eq 6) {
    Start-Job -ScriptBlock {  winget install XPDM1ZW6815MQM -h --accept-package-agreements ----disable-interactivity }
}
# Brave Browser
if ($wshell.Popup("Do you want to install Brave Browser ?",0,"Installer",4+32) -eq 6) {
    Start-Job -ScriptBlock { winget install brave.brave-h --accept-package-agreements ----disable-interactivity }
}
# VS CODE
if ($wshell.Popup("Do you want to install VS CODE ?",0,"Installer",4+32) -eq 6) {
    Start-Job -ScriptBlock { winget install XP9KHM4BK9FZ7Q -h --accept-package-agreements ----disable-interactivity }
}
# VESTA
if ($wshell.Popup("Do you want to install VESTA to the Desktop?",0,"Installer",4+32) -eq 6) {
    Start-Job -ScriptBlock {
        # change dir to desktop
        # check if $HOME/Desktop exists
        if (Test-Path $HOME\Desktop) {
            Set-Location $HOME\Desktop
        }elseif (Test-Path $HOME\OneDrive\Desktop) {
            Set-Location $HOME\OneDrive\Desktop
        }else{
            Write-Host "Desktop not found"
            # Popup
            $wshell.Popup("Desktop folder not found by this script",0,"Installer",0+16)
            exit
        }

        Invoke-WebRequest -Uri https://jp-minerals.org/vesta/archives/3.5.8/VESTA-win64.zip -OutFile $PWD/vesta.zip
        Expand-Archive -Path $PWD/vesta.zip -DestinationPath $PWD/vesta
        Remove-Item $PWD/vesta.zip

    }
}
#Google earth pro
if ($wshell.Popup("Do you want to install Google Earth Pro ?",0,"Installer",4+32) -eq 6) {
    Start-Job -ScriptBlock { winget install Google.EarthPro -h --accept-package-agreements ----disable-interactivity }
}
#MS OFFICE
if ($wshell.Popup("Do you want to Install MS OFFICE ? ",0,"Installer",4+32) -eq 6) {
    # MS Office
    $officeJob = Start-Job -ScriptBlock { winget install Microsoft.Office --accept-package-agreements } -Name "Office"

    # Wait for the job to finish
    Wait-Job -Name "Office"

    #ENDNOTE
    if ($wshell.Popup("Do you want to install ENDNOTE ?",0,"Installer",4+32) -eq 6) {
    
        Start-Job -ScriptBlock { 
            #download from https://download.endnote.com/downloads/21/EN21Inst.exe into $PWD
            Invoke-WebRequest -Uri https://download.endnote.com/downloads/21/EN21Inst.exe -OutFile $PWD\EN21Inst.exe
            # install
            Start-Process -FilePath $PWD\EN21Inst.exe
            # remove the installer
            Remove-Item $PWD\EN21Inst.exe
        }
    }

}
    

if ($wshell.Popup("Do you want to show file-extensions in File-Explorer?",0,"File-Explorer",4+32) -eq 6) {
    # Yes
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name 'HideFileExt' -Value 0
    Stop-Process -Name explorer -Force
    Start-Process explorer.exe
}


if ($wshell.Popup("Do you want to Disable Transparancy effects, Reconmended for Performance",0,"File-Explorer",4+32) -eq 6) {
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "EnableTransparency" -Value 0

}
if ($wshell.Popup("Do you want Move Startbutton to the left of the taskbar?",0,"File-Explorer",4+32) -eq 6) {
    New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarAl" -Value 0 -Force
}

# Load the XAML definition
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Option Selector" Width="400" Height="200">
    <Grid>
        <Label Content="Select Windows search taskbar Presentation:" FontSize="16" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="20,20,0,0"/>
        <ComboBox x:Name="OptionComboBox" FontSize="16" HorizontalAlignment="Left" VerticalAlignment="Top" Width="200" Margin="20,60,0,0">
            <ComboBoxItem Content="Hide"/>
            <ComboBoxItem Content="ICON ONLY"/>
            <ComboBoxItem Content="FullSearchbox -Default"/>
        </ComboBox>
        <Button Content="OK" Name="OKButton" FontSize="16" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Height="30" Margin="250,60,0,0"/>
    </Grid>
</Window>
"@

# Create a window from the XAML definition
$Window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))

# Find the controls
$OptionComboBox = $Window.FindName("OptionComboBox")
$OKButton = $Window.FindName("OKButton")

# Define the click event for the OK button
$OKButton.Add_Click({
    $selectedOption = $OptionComboBox.SelectedItem.Content

    # Set the registry value based on the selection
    switch ($selectedOption) {
        "Hide" {
            New-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Search -Name "SearchboxTaskbarMode" -Value 0 -Force
        }
        "ICON ONLY" {
            New-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Search -Name "SearchboxTaskbarMode" -Value 1 -Force
        }
        "FullSearchbox -Default" {
            New-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Search -Name "SearchboxTaskbarMode" -Value 2 -Force
        }
    }

    $Window.Close()
})

# Show the window
$Window.ShowDialog()
