Dim otrd
Dim status

status = "PASS"

Set otrd = CreateObject("DDVariableStore.TestRecordData")

'Test the InputRecord Property
otrd.InputRecord = "InputRecordTest"

If otrd.InputRecord <> "InputRecordTest" Then
	msgbox "InputRecord Test Failed"
	status = "FAIL"
end if

'Test the LineNumber Property
otrd.LineNumber = 1

If otrd.LineNumber <> 1 Then
	msgbox "LineNumber Test Failed"
	status = "FAIL"
end if


'Test the FileName Property
otrd.FileName = "FileNameTest"

If otrd.FileName <> "FileNameTest" Then
	msgbox "FileName Test Failed"
	status = "FAIL"
end if

'Test the Separator Property
otrd.Separator = "SeparatorTest"

If otrd.Separator <> "SeparatorTest" Then
	msgbox "Separator Test Failed"
	status = "FAIL"
end if

'Test the TestLevel Property
otrd.TestLevel = "TestLevelTest"

If otrd.TestLevel <> "TestLevelTest" Then
	msgbox "TestLevel Test Failed"
	status = "FAIL"
end if

'Test the AppMapName Property
otrd.AppMapName = "AppMapNameTest"

If otrd.AppMapName <> "AppMapNameTest" Then
	msgbox "AppMapName Test Failed"
	status = "FAIL"
end if

'Test the Fac Property
otrd.Fac = "FacTest"

If otrd.Fac <> "FacTest" Then
	msgbox "Fac Test Failed"
	status = "FAIL"
end if

'Test the StatusCode Property
otrd.StatusCode = 4

If otrd.StatusCode <> 4 Then
	msgbox "StatusCode Test Failed"
	status = "FAIL"
end if

'Test the StatusInfo Property
otrd.StatusInfo = "StatusInfoTest"

If otrd.StatusInfo <> "StatusInfoTest" Then
	msgbox "StatusInfo Test Failed"
	status = "FAIL"
end if

'Test the WindowName Property
otrd.WindowName = "WindowNameTest"

If otrd.WindowName <> "WindowNameTest" Then
	msgbox "WindowName Test Failed"
	status = "FAIL"
end if

'Test the WindowGUIID Property
otrd.WindowGUIID = "WindowGUIIDTest"

If otrd.WindowGUIID <> "WindowGUIIDTest" Then
	msgbox "WindowGUIID Test Failed"
	status = "FAIL"
end if

'Test the CompName Property
otrd.CompName = "CompNameTest"

If otrd.CompName <> "CompNameTest" Then
	msgbox "CompName Test Failed"
	status = "FAIL"
end if


'Test the CompGUIID Property
otrd.CompGUIID = "CompGUIIDTest"

If otrd.CompGUIID <> "CompGUIIDTest" Then
	msgbox "CompGUIID Test Failed"
	status = "FAIL"
end if

'Test the TestCommand Property
otrd.TestCommand = "TestCommandTest"

If otrd.TestCommand <> "TestCommandTest" Then
	msgbox "TestCommand Test Failed"
	status = "FAIL"
end if

'Test the STAFHelper Property
'set otrd.STAFHelper = Eval("STAFHelperTest")

'If Not IsObject(otrd.STAFHelper) Then
'	msgbox "STAFHelper Test Failed"
'	status = "FAIL"
'end if

'Test the HookTRDID Property
otrd.HookTRDID = "HookTRDIDTest"

If otrd.HookTRDID <> "HookTRDIDTest" Then
	msgbox "HookTRDID Test Failed"
	status = "FAIL"
end if

msgbox "The property test " & status & "ed!"