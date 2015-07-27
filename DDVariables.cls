VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DDVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Collection of Variables and routines supporting their use."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"DDVariable"
Attribute VB_Ext_KEY = "Member0" ,"DDVariable"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'##############################################################################
'# <a name="ddvariables"></a>
'#
'#    DDVariables.CLS  SOURCECODE
'#
'# MODULE DESCRIPTION:
'#
'#  DDVariables is an ActiveX class that provides an accessible collection of
'#  user-defined items accessed via a user-defined name or key.  The items can
'#  be standard data type items including Objects.  It is up to the user to
'#  properly cast or handle the item returned by the collection.  The items in
'#  the collection do NOT have to all be of the same type.
'#
'#  A DDVariables collection is created from the DDVariableStore Server that
'#  must be already installed and registered on the system.
'#
'#
'#      Set oCollection = New DDVariableStore.DDVariables
'#
'#          or
'#
'#      Set oCollection = CreateObject("DDVariableStore.DDVariables")
'#
'#
'#  Consult the detailed descriptions for the functions available for the
'#  DDVariables collection retrieved.
'#
'#
'#  Members available:
'#
'#      Method      <a href="#SetVariableValue">SetVariableValue</a>
'#      Method      <a href="#GetVariableValue">GetVariableValue</a>
'#      Property    <a href="#Count">Count</a>
'#      Method      <a href="#Remove">Remove</a>
'#      Method      <a href="#ResetAndClear">ResetAndClear</a>
'#
'#
'# This library is defined by:
'#
'# DDVariableStore.CLS      Class Sourcecode for Automation Server
'# DDVariables.CLS          Class Sourcecode and Detailed Documentation
'# DDVariableStore.DLL      Compiled ActiveX DLL library
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 06, 2001
'# History:
'#
'#      AUG 06, 2001    Original Release
'#
'# Copyright (2001) SAS Institute Inc. All rights reserved.
'# NOTICE: Permission to use and distribute this DLL and its
'# documentation, if any, for any purpose and without fee is hereby granted,
'# subject to the following conditions:
'#
'# The above copyright notice and this permission notice must appear in all
'# copies of the code and any related documentation.
'#
'# THIS DLL IS PROVIDED "AS IS", AND SAS MAKES NO REPRESENTATIONS OR WARRANTIES,
'# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO WARRANTIES OF MERCHANTABILITY,
'# FITNESS FOR ANY PARTICULAR PURPOSE, AND NONINFRINGEMENT.  IN NO EVENT SHALL
'# SAS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY ARISING IN CONNECTION
'# WITH THE DLL OR ITS USE.
'##############################################################################

'# PUBLIC DEPENDENCIES

'   COMCAT.MSM              Merge Module for COMCAT.DLL    (DLL, It think)
'   OLEAUT32.MSM            Merge Module for OLEAUT32.DLL  (DLL, It think)
'   MSVBVM60.MSM            Merge Module for MSVBVM60.DLL  (DLL, It think)

'# PRIVATE DEPENDENCIES

'# FORWARD DECLARATIONS

'local variable to hold collection
Private mCol As Collection


'############################################################################
'# <a name="resetandclear"></a>
'#
'#  Function ResetAndClear() As Long
'#
'# DESCRIPTION:
'#
'#  Removes all contents of the collection.  The collection.Count exits as 0.
'#
'#      lCount = oCollection.ResetAndClear()
'#
'#
'# PARAMETERS:
'#
'#       (none)
'#
'#
'# RETURNS:
'#
'#      The number of items in the collection after completion.
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
'#
'############################################################################
Public Function ResetAndClear() As Long
Attribute ResetAndClear.VB_Description = "Resets and Clears the Variable Store.  Removes all existing variables.  Thus the Count should become 0.\r\n\r\nReturns the value of Count."
    Dim index As Long
    On Error GoTo eh
    For index = mCol.count To 1 Step -1
        mCol.Remove (index)
    Next
eh:
    ResetAndClear = 0
End Function


'############################################################################
'# <a name="setvariablevalue"></a>
'#
'#  Function SetVariableValue (varName As String, Value As Variant) As Integer
'#
'# DESCRIPTION:
'#
'#  Add a new Key/Value pair to the collection or change the Value of an existing
'#  Key/Value pair.
'#
'#  Example 1: Store simple string value in item with Key of "password"
'#
'#      Const passKey = "password"
'#      password$ = "MyPassword"
'#      ...
'#      iStatus = oCollection.SetVariableValue(passKey, password)
'#
'#
'#  Example 2: Store a collection of windows in the item with Key of "windows"
'#
'#      Const winKey = "windows"
'#      Dim winCollection  = DDVariableStore.DDVariables
'#      ...
'#      iStatus = oCollection.SetVariableValue(winKey, winCollection)
'#
'#
'# PARAMETERS:
'#
'#      varName     The string Key to uniquely identify each item in the collection.
'#                  If you pass in a pre-existing varName then that item will
'#                  be REPLACED by the Value provided in this call.
'#                  A new unique varName will ADD the new item to the collection.
'#                  An empty varName will cause the routine to exit with error (-1).
'#
'#      Value       a value that can be any legal Variant data type including
'#                  Objects.  The provided Value is sent to the collection if
'#                  varName is valid.  Value is left unmodified regardless of any
'#                  error conditions.
'#
'#
'# RETURNS:
'#
'#       0  on success.
'#      -1  on Invalid input parameters.
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
'#
'############################################################################
Public Function SetVariableValue(varname As String, value As Variant) As Integer
Attribute SetVariableValue.VB_Description = "Set the variable identified by varName to the value provided.\r\n\r\nReturns 0 on success.\r\n-1 on failure."
Attribute SetVariableValue.VB_UserMemId = 0
    
    Dim objNewMember As DDVariable
    Dim newvar As Integer
    
    If Len(varname) = 0 Then
        SetVariableValue = -1
        Exit Function
    End If
    
    On Error Resume Next
    
    Set objNewMember = mCol(varname)
    
    'if it already exists, change its value
    If objNewMember Is Nothing Then
        Set objNewMember = New DDVariable
        'set the properties passed into the method
        objNewMember.Key = varname
        newvar = 1
    End If
    
    If IsObject(value) Then
        Set objNewMember.value = value
    Else
        objNewMember.value = value
    End If
        
    If newvar Then mCol.Add objNewMember, varname

    Set objNewMember = Nothing
    
End Function


'############################################################################
'# <a name="getvariablevalue"></a>
'#
'#  Function GetVariableValue (varName As String, Value As Variant) As Integer
'#
'# DESCRIPTION:
'#
'#  Retrieve the Value from the collection that is associated with the provided
'#  varName Key.  The Value parameter provided must be able to accept the item
'#  from the collection.
'#
'#  Example 1: Retrieve string password value from item with Key of "password"
'#
'#      Const passKey = "password"
'#      Dim sPassword As String
'#      ...
'#      iStatus = oCollection.GetVariableValue(passKey, sPassword)
'#
'#
'#  Example 2: Retrieve a collection of windows from the item with Key of "windows"
'#
'#      Const winKey = "windows"
'#      Dim winCollection  = DDVariableStore.DDVariables
'#      ...
'#      iStatus = oCollection.GetVariableValue(winKey, winCollection)
'#
'#
'# PARAMETERS:
'#
'#      varName     The string Key to uniquely identify the item in the collection.
'#
'#      Value       The provided Value parameter is assigned the Value stored in
'#                  the collection with either the VB Set or Let statements,
'#                  whichever is appropriate.
'#
'#
'# RETURNS:
'#
'#       0      on success.
'#      -1      on failure of any kind.
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
'#
'############################################################################
Public Function GetVariableValue(varname As String, value As Variant) As Integer
Attribute GetVariableValue.VB_Description = "Set the provided ""value"" parameter to the value of the variable specified in the varName parameter.  \r\n\r\nReturn 0 on success.\r\n-1 if the variable is not in the collection."
    
    Dim objNewMember As DDVariable
    
    On Error GoTo GET_ERROR
    Set objNewMember = mCol(varname)
    
    'if it already exists, get the value
    If Not (objNewMember Is Nothing) Then
        If IsObject(objNewMember.value) Then
            Set value = objNewMember.value
        Else
            value = objNewMember.value
        End If
        GetVariableValue = 0        'show success
    End If
        
    Set objNewMember = Nothing
    Exit Function
    
GET_ERROR:
    
    'preset for Failure
    GetVariableValue = -1
    Set objNewMember = Nothing
    Exit Function
        
End Function


'############################################################################
'# <a name="item"></a>
'#
'#  Property Item (vntIndexKey As Variant) As DDVariable          (read only)
'#
'# DESCRIPTION:
'#
'#  Retrieves an item from the collection.  The item returned is identified
'#  either by the Key value assigned to the item, or by its Index.  Valid
'#  indexes are 1 thru .Count
'#
'#
'# PARAMETERS:
'#
'#       vntIndexKey    Identifies which item in the collection to retrieve.
'#                      This can be an Index OR the (string?)Key for the item.
'#
'#
'# RETURNS:
'#
'#      The DDVariable associated with the input parameter if vntIndexKey
'#      is valid.
'#
'#
'# ERRORS:
'#
'#       (none)
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Property Get Item(vntIndexKey As Variant) As DDVariable

    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    On Error Resume Next
    Set Item = mCol(vntIndexKey)
    On Error GoTo 0
    
End Property



'############################################################################
'# <a name="count"></a>
'#
'#  Property Count() As Long                                  (read only)
'#
'# DESCRIPTION:
'#
'#  Retrieves the count of items currently in the collection.
'#
'#
'# PARAMETERS:
'#
'#       (none)
'#
'#
'# RETURNS:
'#
'#       The count of items currently in the collection.
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
'#
'############################################################################
Public Property Get count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    count = mCol.count
End Property


'############################################################################
'# <a name="remove"></a>
'#
'#  Function Remove(vntIndexKey As Variant) As Integer
'#
'# DESCRIPTION:
'#
'#  Removes an item from the collection.  The item to be removed is identified
'#  either by the Key value assigned to the item, or by its Index.  Valid
'#  indexes are 1 thru .Count
'#
'#
'# PARAMETERS:
'#
'#       vntIndexKey    Identifies which item in the collection to remove.
'#                      This can be an Index OR the (string?)Key for the item.
'#
'#
'# RETURNS:
'#
'#       0      on success.
'#      -1      on failure. Like providing an invalid Key or Index.
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
'#
'############################################################################
Public Function Remove(vntIndexKey As Variant) As Integer
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    On Error GoTo REMOVE_ERROR
    mCol.Remove vntIndexKey
    Exit Function
    
REMOVE_ERROR:
    Remove = -1
    Exit Function
    
End Function


'Public Property Get NewEnum() As IUnknown
    'this property allows you to enumerate
    'this collection with the For...Each syntax
'    Set NewEnum = mCol.[_NewEnum]
'End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

