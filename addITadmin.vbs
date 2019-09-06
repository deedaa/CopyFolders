on error resume next
strComputer = "."
Set objGroup = GetObject("WinNT://" & strComputer & "/Administrators,group")
Set objDAGroup = GetObject("WinNT://dlcanon/ITadmin")
objGroup.Add(objDAGroup.ADsPath)
