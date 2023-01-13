$Word = "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"
$Excel = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
$PowerPoint = "C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE"
$Outlook = "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"

# Get the path to the Programs folder in the Start Menu
$startMenuPrograms = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs"

# Get the path to the folder where all the taskbars are located
$taskbarPath = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"

# Create the shortcuts
New-Item -ItemType SymbolicLink -Path "$startMenuPrograms\Word.lnk" -Target $Word
New-Item -ItemType SymbolicLink -Path "$startMenuPrograms\Excel.lnk" -Target $Excel
New-Item -ItemType SymbolicLink -Path "$startMenuPrograms\PowerPoint.lnk" -Target $PowerPoint
New-Item -ItemType SymbolicLink -Path "$startMenuPrograms\Outlook.lnk" -Target $Outlook

# Pin the shortcuts to the taskbar of all users
New-Item -ItemType SymbolicLink -Path "$taskbarPath\Word.lnk" -Target $Word
New-Item -ItemType SymbolicLink -Path "$taskbarPath\Excel.lnk" -Target $Excel
New-Item -ItemType SymbolicLink -Path "$taskbarPath\PowerPoint.lnk" -Target $PowerPoint
New-Item -ItemType SymbolicLink -Path "$taskbarPath\Outlook.lnk" -Target $Outlook