VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "STAFResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private rccode As Long      'return code received from STAF
Private rcstring As String  'return info received from STAF


'##############################################################################
'#
'#  Public Property Get rc () As Long
'#
'#  DESCRIPTION:
'#
'#     Retrieves the current value of our STAF return code
'#
'#  PARAMETERS:
'#
'#     (none)
'#
'#  RETURNS:
'#
'#      The current value of our STAF return code.
'#      This value is Class initialized to -1
'#
'#  ERRORS:
'#
'#      (none)
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Property Get rc() As Long
    rc = rccode
End Property


'##############################################################################
'#
'#  Public Property Let rc (ByVal lNewValue as Long)
'#
'#  DESCRIPTION:
'#
'#     Sets a new value for our STAF return code
'#
'#  PARAMETERS:
'#
'#     lNewValue   the new value to store.
'#
'#  ERRORS:
'#
'#      (none)
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Property Let rc(ByVal lNewValue As Long)
    rccode = lNewValue
End Property


'##############################################################################
'#
'#  Public Property Get result () As String
'#
'#  DESCRIPTION:
'#
'#     Retrieves the current value of our STAF return info string.
'#
'#  PARAMETERS:
'#
'#     (none)
'#
'#  RETURNS:
'#
'#      The current value of our STAF return info string.
'#      This value is Class initialized to "" (empty).
'#
'#  ERRORS:
'#
'#      (none)
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Property Get result() As String
    result = rcstring
End Property


'##############################################################################
'#
'#  Public Property Let result (ByVal sNewValue as String)
'#
'#  DESCRIPTION:
'#
'#     Sets a new value for our STAF return info string
'#
'#  PARAMETERS:
'#
'#     sNewValue   the new value to store.
'#
'#  ERRORS:
'#
'#      (none)
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Property Let result(ByVal sNewValue As String)
    rcstring = sNewValue
End Property

Public Sub reset()
    rccode = -1
    rcstring = ""
End Sub

Private Sub Class_Initialize()
    reset
End Sub

