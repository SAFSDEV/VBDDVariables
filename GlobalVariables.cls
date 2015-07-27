VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A Globally shared DDVariables"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'##############################################################################
'# <a name="globalvariables"></a>
'#
'#    GlobalVariables.CLS  SOURCECODE
'#
'# MODULE DESCRIPTION:
'#
'#      Cache for DDVariables data.
'#      Create an instance of this class through standard ActiveX means.
'#
'#      Although each instance of the class is specific to the process that
'#      creates it, the cache is stored elsewhere and is accessible to all
'#      instances of the class regardless of which process they are in.
'#
'#      This Public class is made available via the DDVariableStore.DLL server.
'#
'#      Typical VisualBasic invocations:
'#
'#          Set oGlobalVars = DDVariableStore.GlobalVariables
'#
'#              OR
'#
'#          Set oGlobalVars = CreateObject("DDVariableStore.GlobalVariables")
'#
'#
'# This library is defined by:
'#
'# GlobalVariables.CLS      ActiveX Class
'# DDVariableStore.DLL      Compiled executable library
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#      JUN 16, 2003    (Carl Nagle) Made SAFSVARS usage available
'#
'# Copyright (2001) SAS Institute
'# GNU General Public License: http://www.opensource.org/licenses/gpl-license.php
'##############################################################################

Private mapper As GlobalMappings

'############################################################################
'# True if we have pushed all our old var values to SAFSVARS after SAFSVARS
'# is launched some time after we started storing values.
'############################################################################
Private bVarsPushed As Boolean

'############################################################################
'# copy ALL locally stored variables to SAFSVARS when we first detect it has
'# been launched.  Will not copy Objects since STAF cannot store them.
'############################################################################
Private Sub copyToSAFSVARS()
    Dim objVar As DDVariable
    Dim varname As String
    Dim varvalue As Variant
    Dim status As Integer
    Dim iindex As Integer
    
    'this must happen first to prevent recursive looping
    bVarsPushed = True
    
    For iindex = 1 To gDDVariables.count
        Set objVar = gDDVariables.Item(iindex)
        varname = objVar.Key
        If IsObject(objVar.value) Then
           'do nothing
        Else
            varvalue = objVar.value
            status = SetVariableValue(varname, varvalue)
        End If
    Next
End Sub

'############################################################################
'# <a name="resetandclear"></a>
'#
'#  Function ResetAndClear() As Long
'#
'# DESCRIPTION:
'#
'#  Removes all global variables.  The Count exits as 0.
'#
'#      lCount = oGlobalVars.ResetAndClear()
'#
'#
'# PARAMETERS:
'#
'#       (none)
'#
'#
'# RETURNS:
'#
'#      The number of items stored after completion.
'#      (Should always return 0.)
'#
'#
'# ERRORS:
'#
'#       (none)
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 06, 2001
'# History:
'#
'#      AUG 06, 2001    Original Release
'#      JUN 16, 2003    (Carl Nagle) Made SAFSVARS usage available
'#
'############################################################################
Public Function ResetAndClear() As Long
Attribute ResetAndClear.VB_Description = "Resets and Clears the Variable Store.  Removes all existing variables.  Thus the Count should become 0.\r\n\r\nReturns the value of Count."
    ResetAndClear = gDDVariables.ResetAndClear
    resetServiceChecks
    bUseVARS = isSAFSVARSAvailable()
    If bUseVARS Then
        Dim status As Long
        message = "RESET"
        status = submit(SAFS_SAFSVARS_SERVICE, message)
    End If
End Function

'############################################################################
'# <a name="setvariablevalue"></a>
'#
'#  Function SetVariableValue (varName As String, Value As Variant) As Integer
'#
'# DESCRIPTION:
'#
'#  Add a new global variable to storage or change the Value of an
'#  existing one.
'#
'#  Example: Store simple string value in item with Key of "password"
'#
'#      Const passKey = "password"
'#      password$ = "MyPassword"
'#      ...
'#      iStatus = oGlobalVars.SetVariableValue(passKey, password)
'#
'#
'# PARAMETERS:
'#
'#      varName     The string Key to uniquely identify each item stored.
'#                  If you pass in a pre-existing varName then that item will
'#                  be REPLACED by the Value provided in this call.
'#                  A new unique varName will ADD the new item to storage.
'#                  An empty varName will cause the routine to exit with error (-1).
'#
'#      Value       a value that can be any legal Variant data type including
'#                  Objects.  The provided Value is stored if
'#                  varName is valid.  Value is left unmodified regardless of any
'#                  error conditions.
'#
'#
'# RETURNS:
'#
'#       0  on success.
'#      -1  on Invalid input parameters.
'#      STAF error codes (1 thru N) may be returned if STAF is in use.
'#
'# ERRORS:
'#
'#       (none)
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 06, 2001
'# History:
'#
'#      AUG 06, 2001    Original Release
'#      JUN 16, 2003    (Carl Nagle) Made SAFSVARS usage available
'#
'############################################################################
Private Function old_SetVariableValue(varname As String, value As Variant) As Integer
Attribute old_SetVariableValue.VB_Description = "Set the variable identified by varName to the value provided.\r\n\r\nReturns 0 on success.\r\n-1 on failure."
    old_SetVariableValue = gDDVariables.SetVariableValue(varname, value)
End Function
Public Function SetVariableValue(varname As String, value As Variant) As Integer
    If Not bUseVARS Then
        SetVariableValue = old_SetVariableValue(varname, value)
        Exit Function
    End If
    
    If Not bVarsPushed Then
        copyToSAFSVARS
    End If
    
    Dim status As Long
    SetVariableValue = -1
    
    If Len(varname) < 1 Then Exit Function
    
    If IsEmpty(value) Then
        message = "SET " & lenMark(varname) & " VALUE """""
        status = submit(SAFS_SAFSVARS_SERVICE, message)
    Else
        If IsObject(value) Then
            status = old_SetVariableValue(varname, value)
        Else
            message = "SET " & lenMark(varname) & " VALUE " & lenMark(CStr(value))
            status = submit(SAFS_SAFSVARS_SERVICE, message)
        End If
    End If
    SetVariableValue = status
End Function


'############################################################################
'# <a name="getvariablevalue"></a>
'#
'#  Function GetVariableValue (varName As String, Value As Variant) As Integer
'#
'# DESCRIPTION:
'#
'#  Retrieve the variable from storage that is associated with the provided
'#  varName Key.  The Value parameter provided must be able to accept the item
'#  from storage.
'#
'#  Example 1: Retrieve string password value from item with Key of "password"
'#
'#      Const passKey = "password"
'#      Dim sPassword As String
'#      ...
'#      iStatus = oGlobalVars.GetVariableValue(passKey, sPassword)
'#
'#
'# PARAMETERS:
'#
'#      varName     The string Key to uniquely identify the item in storage.
'#
'#      Value       The provided Value parameter is assigned the Value stored
'#                  with either the VB Set or Let statements,
'#                  whichever is appropriate.
'#
'#
'# RETURNS:
'#
'#       0      on success.
'#      -1      on failure of any kind.
'#      STAF error codes (1 thru N) may be returned if STAF is in use.
'#
'#
'# ERRORS:
'#
'#       (none)
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 06, 2001
'# History:
'#
'#      AUG 06, 2001    Original Release
'#      JUN 16, 2003    (Carl Nagle) Made SAFSVARS usage available
'#
'############################################################################
Private Function old_GetVariableValue(varname As String, value As Variant) As Integer
Attribute old_GetVariableValue.VB_Description = "Set the provided ""value"" parameter to the value of the variable specified in the varName parameter.  \r\n\r\nReturn 0 on success.\r\n-1 if the variable is not in the collection."
    Dim tempValue As Variant
    Dim status As Integer
    
    old_GetVariableValue = -1
    status = gDDVariables.GetVariableValue(varname, tempValue)
    
    'if we have a stored value, use it
    If status = 0 Then
        old_GetVariableValue = 0
        If IsObject(tempValue) Then
            Set value = tempValue
        Else
            value = tempValue
        End If
        Exit Function
    End If
    
    'if we don't have a stored value, try a lookup in the current AppMap
    tempValue = mapper.GetGUIID(mapper.DefaultAppMapSection, _
                                varname, _
                                mapper.CurrentAppMapPath)
    
    'if we found it in the AppMap, use it
    If Len(tempValue) > 0 Then
        old_GetVariableValue = 0
        value = tempValue
        Exit Function
    End If
    
    'if we don't have it in the AppMap, use any valid input value as a default value
    If Not IsEmpty(value) Then old_GetVariableValue = SetVariableValue(varname, value)

End Function

' #############################################################################
Public Function GetVariableValue(varname As String, value As Variant) As Integer
    
    If Not bUseVARS Then
        GetVariableValue = old_GetVariableValue(varname, value)
        Exit Function
    End If
    
    If Not bVarsPushed Then
        copyToSAFSVARS
    End If
    
    Dim status As Long
    message = "GET " & lenMark(varname)
    status = submit(SAFS_SAFSVARS_SERVICE, message)
    
    'was found?
    If (status = STAF_Ok) Then
        If resultLen > 0 Then
            value = resultStr
        Else
            value = ""
        End If
    
    'was not found?
    Else
        Dim tempValue As Variant
        
        'only retrieve an already set value
        status = gDDVariables.GetVariableValue(varname, tempValue)
        If status = 0 Then
            If Not IsEmpty(tempValue) Then
                If IsObject(tempValue) Then
                    Set value = tempValue
                Else
                    value = tempValue
                End If
            Else
                'should not get here if status = 0?
                value = ""
            End If
        
        Else 'otherwise, use the provided input value of 'value'
            If Not IsEmpty(value) Then
                If IsObject(value) Then
                    status = gDDVariables.SetVariableValue(varname, value)
                Else
                    'confirmed: SetVar must STAY
                    status = SetVariableValue(varname, value)
                End If
            Else 'was empty
                'confirmed: status = -1 must STAY
                status = -1
            End If
        End If
    End If
    GetVariableValue = status
End Function

'############################################################################
'# <a name="count"></a>
'#
'#  Property Count() As Long                                  (read only)
'#
'# DESCRIPTION:
'#
'#  Retrieves the count of items stored.
'#
'#
'# PARAMETERS:
'#
'#       (none)
'#
'#
'# RETURNS:
'#
'#       The count of items stored.
'#
'#
'# ERRORS:
'#
'#       (none)
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 06, 2001
'# History:
'#
'#      AUG 06, 2001    Original Release
'#      JUN 16, 2003    (Carl Nagle) Made SAFSVARS usage available
'#
'############################################################################
Public Property Get count() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Count
    If Not bUseVARS Then
        count = gDDVariables.count
        Exit Function
    End If
    
    Dim status As Long
    Dim value As Long
    
    message = "COUNT"
    status = submit(SAFS_SAFSVARS_SERVICE, message)
    On Error Resume Next
    If (status = STAF_Ok) And (resultLen > 0) Then value = CLng(resultStr)
    value = value + gDDVariables.count  'add "objects"
    count = value
End Property


'############################################################################
'# <a name="remove"></a>
'#
'#  Function Remove(vntIndexKey As Variant) As Integer
'#
'# DESCRIPTION:
'#
'#  Removes an item from storage.  The item to be removed is identified
'#  either by the Key value assigned to the item, or by its Index.  Valid
'#  indexes are 1 thru .Count
'#
'#
'# PARAMETERS:
'#
'#       vntIndexKey    Identifies which item to remove.
'#                      This can be an Index OR the (string?)Key for the item.
'#
'#
'# RETURNS:
'#
'#       0      on success.
'#      -1      on failure. Like providing an invalid Key or Index.
'#      STAF error codes (1 thru N) may be returned if STAF is in use.
'#
'#
'# ERRORS:
'#
'#       (none)
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 06, 2001
'# History:
'#
'#      AUG 06, 2001    Original Release
'#      JUN 16, 2003    (Carl Nagle) Made SAFSVARS usage available
'#
'############################################################################
Private Function old_Remove(vntIndexKey As Variant) As Integer
    old_Remove = gDDVariables.Remove(vntIndexKey)
End Function

Public Function Remove(vntIndexKey As Variant) As Integer
    If Not bUseVARS Then
        Remove = old_Remove(vntIndexKey)
        Exit Function
    End If
    message = "DELETE " & lenMark(CStr(vntIndexKey))
    Remove = CInt(submit(SAFS_SAFSVARS_SERVICE, message))
End Function

'############################################################################
'# <a name="resetstafchecks"></a>
'#
'#   Sub ResetSTAFChecks ()
'#
'# DESCRIPTION:
'#
'#  Recheck STAF services to see if they are available.
'#
'# PARAMETERS:
'#
'#       (none)
'#
'# ERRORS:
'#
'#       (none)
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: DEC 15, 2003
'# History:
'#
'#      DEC 15, 2003    Original Release
'#
'############################################################################
Public Function ResetSTAFChecks()
    resetServiceChecks
    If (isSAFSVARSAvailable) And (handle > 1) Then
        bUseVARS = True
    Else
        bUseVARS = False
    End If
End Function

Private Sub Class_Initialize()
    
    bVarsPushed = False
    
    If gDDVariables Is Nothing Then
        Set gDDVariables = New DDVariables
        glngDDVRefCount = 0    'just to make sure
    End If
    
    ResetSTAFChecks
    
    glngDDVRefCount = glngDDVRefCount + 1
    Set mapper = New GlobalMappings
End Sub

Private Sub Class_Terminate()
    glngDDVRefCount = glngDDVRefCount - 1
    If glngDDVRefCount = 0 Then
        Set gDDVariables = Nothing
    End If
    Set mapper = Nothing
    
    If (glngDDVRefCount + glngDDMRefCount = 0) And (handle > 1) Then
        ProcUnregister handle
    End If
    
End Sub
