Do
	Set TypeLib = CreateObject("Scriptlet.TypeLib")
	GUID = TypeLib.Guid
	GUID = left(GUID, len(GUID)-3)
	GUID = right(GUID, len(GUID)- 1)
	Set WshShell = WScript.CreateObject("WScript.Shell")
	WshShell.Run "cmd.exe /C echo " & GUID & "| clip", 0, TRUE
Loop until InputBox("Generateur de GUID","GUID",GUID) = ""

Set WshShell = Nothing
Set TypeLib = Nothing