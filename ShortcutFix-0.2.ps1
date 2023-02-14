#To run this, save this on a C:\temp folder (create if you dont have one) and on a cmd as administrator write: "powershell -executionpolicy bypass -File "ShortcutFix-0.2.ps1". Please don't forget to run as admin.


#Check for bitness
$office = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\*"
$office_bitness = $office.Bitness

if ($office_bitness -eq "x86") {
    $Outlook = "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
    $Word = "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"
    $Excel = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
    $PowerPoint = "C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE"
    $StartMenu = [Environment]::GetFolderPath("CommonPrograms")
    $Shortcut = New-Object -ComObject WScript.Shell

    #Icon Outlook
    $Link32 = $Shortcut.CreateShortcut("$StartMenu\Outlook.lnk")
    $Link32.TargetPath = $Outlook
    $Link32.Save()
    
    #Icon Word
    $Link32 = $Shortcut.CreateShortcut("$StartMenu\Word.lnk")
    $Link32.TargetPath = $Word
    $Link32.Save()
    #Icon Excel

    $Link32 = $Shortcut.CreateShortcut("$StartMenu\Excel.lnk")
    $Link32.TargetPath = $Excel
    $Link32.Save()

    #Icon Powerpoint
    $Link32 = $Shortcut.CreateShortcut("$StartMenu\PowerPoint.lnk")
    $Link32.TargetPath = $PowerPoint
    $Link32.Save()

} else {

    $Outlook = "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
    $Word = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
    $Excel = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
    $PowerPoint = "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
    $StartMenu = [Environment]::GetFolderPath("CommonPrograms")
    $Shortcut = New-Object -ComObject WScript.Shell

    #Icon Outlook
    $Link64 = $Shortcut.CreateShortcut("$StartMenu\Outlook.lnk")
    $Link64.TargetPath = $Outlook
    $Link64.Save()

    #Icon Word
    $Link64 = $Shortcut.CreateShortcut("$StartMenu\Word.lnk")
    $Link64.TargetPath = $Word
    $Link64.Save()
    #Icon Excel

    $Link64 = $Shortcut.CreateShortcut("$StartMenu\Excel.lnk")
    $Link64.TargetPath = $Excel
    $Link64.Save()

    #Icon Powerpoint
    $Link64 = $Shortcut.CreateShortcut("$StartMenu\PowerPoint.lnk")
    $Link64.TargetPath = $PowerPoint
    $Link64.Save()
}

##############################################################

## This feature has been made invalid by  Microsoft ##
<#

$taskbarPath = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" 

#Icon Outlook
$Link1 = $Shortcut.CreateShortcut("$taskbarPath\Outlook.lnk")
$Link1.TargetPath = $Outlook
$Link1.Save()

#Icon Word
$Link1 = $Shortcut.CreateShortcut("$taskbarPath\Word.lnk")
$Link1.TargetPath = $Word
$Link1.Save()

#Icon Excel
$Link1 = $Shortcut.CreateShortcut("$taskbarPath\Excel.lnk")
$Link1.TargetPath = $Excel
$Link1.Save()

#Icon Powerpoint
$Link1 = $Shortcut.CreateShortcut("$taskbarPath\PowerPoint.lnk")
$Link1.TargetPath = $PowerPoint
$Link1.Save()
#>
