Do
	Set TypeLib = CreateObject("Scriptlet.TypeLib")
	GUID = TypeLib.Guid
	GUID = left(GUID, len(GUID)-3)
	GUID = right(GUID, len(GUID)- 1)
	Set TypeLib = Nothing
Loop until InputBox("Générateur de GUID","GUID",GUID) = ""
