#################################################FORM ATTRIBUTES###############################################
<#
          This section contains data to the form which is the GUI DO NOT ALTER                                
#>
#region ############################SCRIPT CONFIG############################
Add-Type -AssemblyName System.Windows.Forms
$form = New-Object System.Windows.Forms.Form # form instantiation
$Global:form = $form
$size = New-Object System.Drawing.Size(1300, 900)
$form.Size = $size
$form.StartPosition = 'CenterScreen'
$Font = New-Object System.Drawing.Font("Verdana Pro", 9)
$Form.Font = $Font
$Icon = New-Object System.Drawing.Icon ("C:\Users\cereal\Desktop\powershell\checklist\resources\comp.ico")
$form.Icon = $Icon
#endregion
############################################    PRECONFIGURATION   ############################################
<#
         This section may be updated with updated versions of additional .EXEs that need to run               
#>
#region  ############################BACKGROUND TASKS###########################
<#Use the robocopy generator to store packages into the IT folder#>
$itFolder = "C:\IT"
If(!(test-path $itFolder))
{
New-Item -ItemType directory -Path C:\IT
robocopy F:\checklist\custom_packages\cCleaner C:\IT\cCleaner /e /MIR /R:0 /W:0 /NP /FFT /log:"C:\IT\cCleaner_LOG.txt" /NDL

}
Install-Module -Name PSWindowsUpdate
#$runtimes = "C:\Users\cereal\Desktop\powershell\checklist\runtime\install_all.bat"
#start $runtimes
#silent install of alt Chrome that goes to all profiles on PC
#$chrome = "C:\Users\cereal\Desktop\powershell\checklist\ninite\chromesetup.exe"
#start $chrome /silent /install
####Chocolatey is a package management framework (i.e. it installs adobe flash without the verbose code)
#$chocolate = Set-ExecutionPolicy Bypass -Scope Process -Force; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
#start $chocolate
#check for manufactuer: determines the driver utility to install
<#
$manufacturer = (Get-WmiObject -Class:Win32_ComputerSystem).Manufacturer
choco install flashplayerplugin
choco install adobereader
choco install greenshot
choco install irfanview
choco install cutepdf
choco install vlc
choco install winrar
choco install dotnet3.5
choco install dotnet4.5
choco install classic-shell
#>

#endregion 
#region ############################FORM CONTENTS############################
#################################################FORM CONTENTS#################################################
$Form.Text = "EMC IT Solutions Computer Preperation Checklist Application"
#################################################SYSINFO_GROUPBOX CONFIG#######################################
$SYSINFO_GROUPBOX = New-Object System.Windows.Forms.GroupBox
$SYSINFO_GROUPBOX.Location = New-Object System.Drawing.Point(27.5, 32)
$SYSINFO_GROUPBOX.Size = New-Object System.Drawing.Size(775, 325)
$SYSINFO_GROUPBOX.TabIndex = 0
$SYSINFO_GROUPBOX.TabStop = $false
$SYSINFO_GROUPBOX.Text = "System Information"
###############################################################################################################
############Order Label#############
$orderLabel = New-object System.Windows.Forms.Label
$orderLabel.Text = "Order:"
$orderLabel.Location = New-Object System.Drawing.Point(50, 55)
$orderLabel.size = New-Object System.Drawing.Size(100, 55)
$SYSINFO_GROUPBOX.Controls.Add($orderLabel)
############Order TextBox###########
$orderBox = New-Object System.Windows.Forms.TextBox
$orderBox.Location = New-Object System.Drawing.Point(155,55)
$orderBox.Size = New-Object System.Drawing.Size(100, 55)
$SYSINFO_GROUPBOX.Controls.Add($orderBox)
############PC Name Label#############
$pcNameLabel = New-object System.Windows.Forms.Label
$pcNameLabel.Text = "PC Name:"
$pcNameLabel.Location = New-Object System.Drawing.Point(260, 55)
$pcNameLabel.size = New-Object System.Drawing.Size(100, 55)
$SYSINFO_GROUPBOX.Controls.Add($pcNameLabel)
############PC Name TextBox###########
$pcNameBox = New-Object System.Windows.Forms.TextBox
$pcNameBox.Text = $env:COMPUTERNAME
$pcNameBox.Location = New-Object System.Drawing.Point(365,55)
$pcNameBox.Size = New-Object System.Drawing.Size(100, 55)
$SYSINFO_GROUPBOX.Controls.Add($pcNameBox)
############Serial Label#############
$serialLabel = New-object System.Windows.Forms.Label
$serialLabel.Text = "Serial:"
$serialLabel.Location = New-Object System.Drawing.Point(470, 55)
$serialLabel.size = New-Object System.Drawing.Size(100, 55)
$SYSINFO_GROUPBOX.Controls.Add($serialLabel)
############Serial TextBox###########
$serialBox = New-Object System.Windows.Forms.TextBox
$serial = (gwmi win32_bios).SerialNumber
$serialBox.Text = $serial
$serialBox.Location = New-Object System.Drawing.Point(575, 55)
$serialBox.Size = New-Object System.Drawing.Size(100, 55)
$SYSINFO_GROUPBOX.Controls.Add($serialBox)
############RMM Label#############
#$label03 = New-object System.Windows.Forms.Label
#$label03.Text = "RMM:"
#$label03.Location = New-Object System.Drawing.Point(470, 55)
#$label03.size = New-Object System.Drawing.Size(100, 55)
#$SYSINFO_GROUPBOX.Controls.Add($label03)
#############RMM TextBox###########
#$textBox03 = New-Object System.Windows.Forms.TextBox
#$textBox03.Location = New-Object System.Drawing.Point(550,55)
#$textBox03.Size = New-Object System.Drawing.Size(100, 55)
#$fileRMM = "C:\Users\cereal\Desktop\powershell\checklist\rmm.csv"
#$fileSEP = "C:\Users\cereal\Desktop\powershell\checklist\SEP.csv"
#$SYSINFO_GROUPBOX.Controls.Add($textBox03) 
###READING FROM CSV FILE OF LIST OF CLIENTS TO CHECK FOR WHICH AV IT GETS
###MAYBE CHECK TO SEE IF THE PROGRAM IS JUST INSTALLED
############Case Number Label#############
$caseNLabel = New-object System.Windows.Forms.Label
$caseNLabel.Text = "Case Number:"
$caseNLabel.Location = New-Object System.Drawing.Point(50, 115)
$caseNLabel.size = New-Object System.Drawing.Size(100, 55)
$SYSINFO_GROUPBOX.Controls.Add($caseNLabel)
############Case Number TextBox###########
$caseNBox = New-Object System.Windows.Forms.TextBox
$caseNBox.Location = New-Object System.Drawing.Point(155,115) 
$caseNBox.Size = New-Object System.Drawing.Size(100, 55)
$SYSINFO_GROUPBOX.Controls.Add($caseNBox)
############Client Label#############
$caseNLabel = New-object System.Windows.Forms.Label
$caseNLabel.Text = "Client:"
$caseNLabel.Location = New-Object System.Drawing.Point(50, 175)
$caseNLabel.size = New-Object System.Drawing.Size(100, 55)
$SYSINFO_GROUPBOX.Controls.Add($caseNLabel)
############Client TextBox###########
$caseNBox = New-Object System.Windows.Forms.TextBox
$caseNBox.Location = New-Object System.Drawing.Point(155,175) 
$caseNBox.Size = New-Object System.Drawing.Size(100, 55)
$SYSINFO_GROUPBOX.Controls.Add($caseNBox)
############Make Label#############
$makeLabel = New-object System.Windows.Forms.Label
$makeLabel.Text = "Brand:"
$makeLabel.Location = New-Object System.Drawing.Point(260, 115)
$makeLabel.size = New-Object System.Drawing.Size(100, 55)
$SYSINFO_GROUPBOX.Controls.Add($makeLabel)
############Make TextBox###########
$makeBox = New-Object System.Windows.Forms.TextBox
$make = (gwmi Win32_ComputerSystem).Manufacturer
$makeBox.Text = $make
$makeBox.Location = New-Object System.Drawing.Point(365,115)
$makeBox.Size = New-Object System.Drawing.Size(100, 55)
$SYSINFO_GROUPBOX.Controls.Add($makeBox)
############Model Label#############
$modelLabel = New-object System.Windows.Forms.Label
$modelLabel.Text = "Model:"
$modelLabel.Location = New-Object System.Drawing.Point(470, 115)
$modelLabel.size = New-Object System.Drawing.Size(100, 55)
$SYSINFO_GROUPBOX.Controls.Add($modelLabel)
############Model TextBox###########
$modelBox = New-Object System.Windows.Forms.TextBox
$model = (gwmi Win32_ComputerSystem).Model
$modelBox.Text = $model
$modelBox.Location = New-Object System.Drawing.Point(575,115)
$modelBox.Size = New-Object System.Drawing.Size(100, 55)
$SYSINFO_GROUPBOX.Controls.Add($modelBox)
############Time Label#############
$timeLabel = New-object System.Windows.Forms.Label
$timeLabel.Text = "Time:"
$timeLabel.Location = New-Object System.Drawing.Point(260, 175)
$timeLabel.size = New-Object System.Drawing.Size(100, 55)
$SYSINFO_GROUPBOX.Controls.Add($timeLabel)
############Time TextBox###########
$timeBox = New-Object System.Windows.Forms.TextBox
tzutil /s "Eastern Standard Time"
$time = Get-TimeZone
$timeBox.Text = $time
$timeBox.Location = New-Object System.Drawing.Point(360,175) 
$timeBox.Size = New-Object System.Drawing.Size(100, 55)
$SYSINFO_GROUPBOX.Controls.Add($timeBox)

$form.Controls.Add($SYSINFO_GROUPBOX) #<-- loads the GroupBox holding the above items to the form, must load after all items in groupbox
#################################################SYSINFO_GROUPBOX END##################################################
#################################################ADMIN_GROUPBOX CONFIG#################################################
#THIS SECTION CONTAINS INTERACTIVE PROCESSES THAT CANNOT BE JUST RUN#
$ADMIN_GROUPBOX = New-Object System.Windows.Forms.GroupBox
$ADMIN_GROUPBOX.Location = New-Object System.Drawing.Point(27.5, 370)
$ADMIN_GROUPBOX.Size = New-Object System.Drawing.Size(900, 200)
$ADMIN_GROUPBOX.Text = "Administrative Interaction"
#######################################################################################################################
############PC Name Button###########
$renameButton = New-Object System.Windows.Forms.Button
$computerName = Get-WmiObject Win32_ComputerSystem 
$renameButton.Location = New-Object System.Drawing.Point(55,442.5) 
$renameButton.Size = New-Object System.Drawing.Size(100, 55)
$renameButton.Text = "Rename Computer"
$renameButton.TabIndex = 0
$renameButton.Add_Click(
        {    
		[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
        $name = [Microsoft.VisualBasic.Interaction]::InputBox("Enter PC Name")
        $computerName.Rename($name)
        }
    )
$form.Controls.Add($renameButton)
############Enable LUSR Button###########
$adminButton = New-Object System.Windows.Forms.Button
$adminButton.Location = New-Object System.Drawing.Point(160, 442.5) 
$adminButton.Size = New-Object System.Drawing.Size(100, 55)
$adminButton.Text = " Activate Admin" 
$adminButton.TabIndex = 1
$form.Controls.Add($adminButton)
$adminButton.Add_Click(
        { 
		[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
        $name = [Microsoft.VisualBasic.Interaction]::InputBox("Admin Name")
        $pass = [Microsoft.VisualBasic.Interaction]::InputBox("Password")
        $password = $pass | ConvertTo-SecureString -AsPlainText -Force
        New-LocalUser -Name $name -Password $password -Description "Client Specific local admin" 
        }
    )
############UAC Button#############
$uacButton = New-Object System.Windows.Forms.Button
$uacButton.Location = New-Object System.Drawing.Point(265,442.5) 
$uacButton.Size = New-Object System.Drawing.Size(100, 55)
$uacButton.Text = "Disable UAC"
$uacButton.TabIndex = 0
$uacButton.Add_Click(
        {
                ##CONNECT TO REGISTRY VIA PSSESSION
                Set-Service -Name remoteregistry -ComputerName $env:COMPUTERNAME -StartupType Manual
                Get-Service remoteregistry -ComputerName $env:COMPUTERNAME | Start-Service
                $RootKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $env:COMPUTERNAME)
                $numVersion = (Get-CimInstance Win32_OperatingSystem).Version
                $numSplit = $numVersion.split(".")[0]
                if ($numSplit -eq 10) {
                Set-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "ConsentPromptBehaviorAdmin" -Value "0"
                } ElseIf ($numSplit -eq 6) {
                $enumSplit = $numSplit.split(".")[1]
                if ($enumSplit -eq 1 -or $enumSplit -eq 0) {
                Set-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableLUA" -Value "0"
                } Else {
                Set-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "ConsentPromptBehaviorAdmin" -Value "0"
                }
                } ElseIf ($numSplit -eq 5) {
                Set-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableLUA" -Value "0"
                } Else {
                Set-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "ConsentPromptBehaviorAdmin" -Value "0"
                }
        }
)
$form.Controls.Add($uacButton)
############WinUpdate Button#############
$winUpdateButton = New-Object System.Windows.Forms.Button
$winUpdateButton.Location = New-Object System.Drawing.Point(370, 442.5) 
$winUpdateButton.Size = New-Object System.Drawing.Size(100, 55)
$winUpdateButton.Text = "Update"
$winUpdateButton.Add_Click(
        {
           Powershell "C:\Users\cereal\Desktop\powershell\checklist\update-Windows.ps1"
        }
)
$form.Controls.Add($winUpdateButton)
#################################################ADMIN_GROUPBOX CONFIG END#################################################
##################################################SYS_CONFIG CHECK START###################################################
<#Section contains settings that have been modified and require a validation#>
############PowerConfig Label#############
$powerLabel = New-object System.Windows.Forms.Label
$powerLabel.Text = "PowerConfig:"
$powerLabel.Location = New-Object System.Drawing.Point(55, 625)
$powerLabel.size = New-Object System.Drawing.Size(100, 55)
$Form.Controls.Add($powerLabel)
############PowerConfig TextBox###########
#AC - plugged in with cord
#DC - battery (for laptops)
$powerBox = New-Object System.Windows.Forms.label
powercfg -change -monitor-timeout-ac 0
powercfg -change -monitor-timeout-dc 45
powercfg -change disk-timeout-ac 0
powercfg -change disk-timeout-dc 120
powercfg -change standby-timeout-ac 0
powercfg -change standby-timeout-dc 0
powercfg -change hibernate-timeout-ac 0
powercfg -change hibernate-timeout-dc 0
Function Detect-Laptop
{
Param( [string]$computer = “localhost” )
$isLaptop = $false
#The chassis is the physical container that houses the components of a computer. Check if the machine’s chasis type is 9.Laptop 10.Notebook 14.Sub-Notebook
if(Get-WmiObject -Class win32_systemenclosure -ComputerName $computer | Where-Object { $_.chassistypes -eq 9 -or $_.chassistypes -eq 10 -or $_.chassistypes -eq 14})
{ $isLaptop = $true }
#Shows battery status , if true then the machine is a laptop.
if(Get-WmiObject -Class win32_battery -ComputerName $computer)
{ $isLaptop = $true }
$isLaptop
}
If(Detect-Laptop) {$powerBox.Text = "Laptop"}
else {$powerBox.Text = "Desktop"}
$powerBox.Location = New-Object System.Drawing.Point(160, 625)
$powerBox.size = New-Object System.Drawing.Size(100, 55)
$form.Controls.Add($powerBox)
############Updates INFO Label#############
$updateLabel = New-object System.Windows.Forms.Label
$updateLabel.Text = "Build:"
$updateLabel.Location = New-Object System.Drawing.Point(265, 625)
$updateLabel.size = New-Object System.Drawing.Size(100, 55)
$Form.Controls.Add($updateLabel)
############Updates INFO Label2#############
$updateInfo = New-Object System.Windows.Forms.Label
$build = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ReleaseId).ReleaseId
$updateInfo.Text = $build
$updateInfo.Location = New-Object System.Drawing.Point(370, 625)
$updateInfo.Size = New-Object System.Drawing.Size(100, 55)
$form.Controls.Add($updateInfo)
############Submit Button#############
$submitForm = New-Object System.Windows.Forms.Button
$submitForm.Location = New-Object System.Drawing.Size(200, 700)
$submitForm.Size = New-Object System.Drawing.Size(120,23)
$submitForm.Text = "Submit"
$form.Controls.Add($submitForm)
$submitForm.Add_Click(
    {
    Export-Csv
    }
)#>
$Form.ShowDialog()#<--dsiplays the form | must load last
#endregion #FORM CONTENTS#
#region references
#&&&&&&&&&&&&&&&&&&&&&&&&&&&&&#&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&PROGRAM INFORMATION&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&#&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&#
# Drawing.Point is over and down | horizontal than vertical i.e. Drawing.Point(X-axis, Y-Axis) starting from the top left most point.
# X value of text box location should be about 55 points more than the X value of the label depending on word length
# is it possible to host this on a google drive account and make something reference it from a client
#
# add below line for setting default apps
#Default program associations for all users can be defined in: HKEY_LOCAL_MACHINE\Software\Classes
# Circa 2019 Adobe will not allow their flash or Reader programs be automated without a licensing agreement #
#execution policy is set to bypass on an as needed basis, the file itself is set to bypass and the processes executed from it, not the entire policy. 
#&&&&&&&&&&&&&&&&&&#&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&PROGRAM INFORMATION&&&&&&&#&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&#
<#

If this were stored on the server, it could be called from a machine on the network
-the file pathes would not have to be changed, but must be full path names
-then a user would just call the script using another script
       &&&&&&&&&&#&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& MEASUREMENT REFERENCE &&&&&&&&&&&&&&#&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
SIZE
Merged Cell: BCD-4 = 1.3" length x .5" height
1 inch      = 100 points
1/17th inch = 5 points

"actual example" size--> (100, 55)

pattern is label:textbox
example is location(x1, y2) size(x3, y4)
x3 will need to be ~5px more and y2 should be the same as y4 to create same line effect
going over by one x value is roughly adding 105 points to the previous locations's x value
line 1 Y  = 55
line 2 Y  = 115
line 3 Y  = 175
line 4 Y  = 235
line 5 Y  = 295
line 6 Y  = 355
      &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&  REFERENCE &&&&&&&&&&&&&&&&#&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
GROUP BOXES: http://blog.dbsnet.fr/powershell-gui
RMM API INTEGRATION: http://help.logicnow.com/remote-management/helpcontents/index.html?data_extraction_api.htm
workETC integration made by Veetro http://admin.worketc.com/xml
ADO: https://devblogs.microsoft.com/scripting/hey-scripting-guy-how-can-i-write-to-excel-without-using-excel/
updates & possible fixes
PS requires full path names to work properly, any mapping, calling, or assining should be done with full path name, or you risk confusing the powershell.
BIOS tweaks: http://www.systanddeploy.com/2019/03/list-and-change-bios-settings-with.html MANUFACTURER SPECIFIC! 
#>
#endregion #REFERENCES#