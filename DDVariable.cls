VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DDVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Holds a variable value"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'##############################################################################
'# <a name="ddvariable"></a>
'#
'#    DDVariable.CLS  SOURCECODE
'#
'# MODULE DESCRIPTION:
'#
'#      Data storage class for the DDVariables, GlobalVariables, and
'#      GlobalMappings collections.
'#
'#      The collections classes mentioned above will create their own DDVariable
'#      class instances as necessary.
'#
'#      Users can create instances of this class through standard ActiveX
'#      means as needed for their own purposes.
'#
'#      The DDVariable class merely provides to properties:
'#
'#          Key         String      (read/write once)
'#          value       Variant     (read/write)
'#
'#
'#      This Public class is made available via the DDVariableStore.DLL server.
'#
'#      Typical VisualBasic invocations:
'#
'#          Set oStringUtils = DDVariableStore.DDVariable
'#
'#              OR
'#
'#          Set oStringUtils = CreateObject("DDVariableStore.DDVariable")
'#
'#
'# This library is defined by:
'#
'# DDVariable.CLS           ActiveX Class
'# DDVariableStore.DLL      Compiled executable library
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'#
'# Copyright (2001) SAS Institute Inc. All rights reserved.
'# NOTICE: Permission to use, copy, modify and distribute this code and its
'# documentation, if any, for any purpose and without fee is hereby granted,
'# subject to the following conditions:
'#
'# The above copyright notice and this permission notice must appear in all
'# copies of the code and any related documentation.
'#
'# THIS CODE IS PROVIDED "AS IS", AND SAS MAKES NO REPRESENTATIONS OR WARRANTIES,
'# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO WARRANTIES OF MERCHANTABILITY,
'# FITNESS FOR ANY PARTICULAR PURPOSE, AND NONINFRINGEMENT.  IN NO EVENT SHALL
'# SAS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY ARISING IN CONNECTION
'# WITH THE CODE OR ITS USE.
'##############################################################################


'# FORWARD DECLARATIONS

Private mKey As String

'local variable(s) to hold property value(s)
Private mvarValue As Variant 'local copy


' ###########################################################################
Public Property Let value(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Value = 5
    mvarValue = vData
End Property


' ###########################################################################
Public Property Set value(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Value = Form1
    Set mvarValue = vData
End Property


'############################################################################
'# <a name="value"></a>
'#
'#  Property value As Variant
'#
'# DESCRIPTION:
'#
'#      The value of the property can be any valid type of VB data type that
'#      can be stored in a VB Variant.  This includes Objects.
'#
'#
'# USAGE:
'#
'#      object.value = varValue     'sets the value of the property
'#
'#      varValue = object.value     'gets a standard data type from the property
'#
'#      (where 'object' is a DDVariable object)
'#
'#
'# DATA TYPE:
'#
'#      VB Variant
'#
'#
'# ERRORS:
'#
'#       none
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Property Get value() As Variant
Attribute value.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Value
    If IsObject(mvarValue) Then
        Set value = mvarValue
    Else
        value = mvarValue
    End If
End Property


'############################################################################
'# <a name="key"></a>
'#
'#  Property Key As String
'#
'# DESCRIPTION:
'#
'#      The property must be a String.
'#      This is the Key used to identify unique objects in Collections.
'#      Once this has been set, it cannot be changed.
'#
'#      This value is usually set by the Collection's method.  Those collections
'#      create their own instances of the DDVariable class as needed.
'#
'#      A user/developer will normally not set this value unless they are
'#      creating their own instances of this DDVariable class for some other
'#      purpose.
'#
'#
'# USAGE:
'#
'#      object.Key = varValue     'sets the value of the property
'#
'#      varValue = object.value     'gets a standard data type from the property
'#
'#      (where 'object' is a DDVariable object)
'#
'#
'# DATA TYPE:
'#
'#      String
'#
'#
'# ERRORS:
'#
'#       none
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Property Let Key(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Value = 5
    If Len(mKey) = 0 Then mKey = vData
End Property


' ###########################################################################
Public Property Get Key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Value
    Key = mKey
End Property



