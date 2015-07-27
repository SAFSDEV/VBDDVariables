VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRecordData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mInputRecord As String
Private mLineNumber As Long
Private mfilename As String
Private mseparator As String
Private mtestlevel As String
Private mappmapname As String
Private mfac As String
Private mstatuscode As Long
Private mstatusinfo As String
Private mwinname As String
Private owinGUIID As Variant
Private mcompname As String
Private ocompGUIID As Variant
Private mtestcommand As String
Private mrecordtype As String
Private mcompType As String
Private mcompClass As String
Private mcompModule As String
Private mcompHandle As Long
Private mEnvironment As String
Private mstafhelper As Variant
Private mhooktrd As String  'SAFS/HOOK/' by default


Public Function getInputRecord() As String
    getInputRecord = mInputRecord
End Function
Public Sub setInputRecord(ByVal sNewValue As String)
    mInputRecord = sNewValue
End Sub
Public Property Get InputRecord() As String

    InputRecord = mInputRecord
    
End Property

Public Property Let InputRecord(ByVal sNewValue As String)
    
    mInputRecord = sNewValue
    
End Property
Public Property Get RecordType() As String
    RecordType = mrecordtype
End Property
Public Property Let RecordType(ByVal sNewValue As String)
    mrecordtype = sNewValue
End Property

Public Property Get CompType() As String

    CompType = mcompType
    
End Property

Public Property Let CompType(ByVal sNewValue As String)
    
    mcompType = sNewValue
    
End Property

Public Property Get CompClass() As String

    CompClass = mcompClass
    
End Property

Public Property Let CompClass(ByVal sNewValue As String)
    
    mcompClass = sNewValue
    
End Property

Public Property Get CompModule() As String

    CompModule = mcompModule
    
End Property

Public Property Let CompModule(ByVal sNewValue As String)
    
    mcompModule = sNewValue
    
End Property

Public Property Get Environment() As String

    Environment = mEnvironment
    
End Property

Public Property Let Environment(ByVal sNewValue As String)
    
    mEnvironment = sNewValue
    
End Property

Public Property Get CompHandle() As Long

    CompHandle = mcompHandle
    
End Property

Public Property Let CompHandle(ByVal lNewValue As Long)
    
    mcompHandle = lNewValue
    
End Property

Public Function getLineNumber() As Long
    getLineNumber = mLineNumber
End Function
Public Sub setLineNumber(lNewValue As Long)
    mLineNumber = lNewValue
End Sub
Public Property Get LineNumber() As Long

    LineNumber = mLineNumber
    
End Property

Public Property Let LineNumber(ByVal lNewValue As Long)
    
    mLineNumber = lNewValue
    
End Property

Public Function getFileName() As String
    getFileName = mfilename
End Function
Public Sub setFileName(sNewValue As String)
    mfilename = sNewValue
End Sub
Public Property Get FileName() As String
    FileName = mfilename
End Property
Public Property Let FileName(sNewValue As String)
    mfilename = sNewValue
End Property

Public Function getSeparator() As String
    getSeparator = mseparator
End Function
Public Sub setSeparator(sNewValue As String)
    mseparator = sNewValue
End Sub
Public Property Get Separator() As String
    Separator = mseparator
End Property
Public Property Let Separator(sNewValue As String)
    mseparator = sNewValue
End Property

Public Function getTestLevel() As String
    getTestLevel = mtestlevel
End Function
Public Sub setTestLevel(sNewValue As String)
    mtestlevel = sNewValue
End Sub

Public Property Get TestLevel() As String
    TestLevel = mtestlevel
End Property
Public Property Let TestLevel(sNewValue As String)
    mtestlevel = sNewValue
End Property

Public Function getAppMapName() As String
    getAppMapName = mappmapname
End Function
Public Sub setAppMapName(sNewValue As String)
    mappmapname = sNewValue
End Sub
Public Property Get AppMapName() As String
    AppMapName = mappmapname
End Property
Public Property Let AppMapName(sNewValue As String)
    mappmapname = sNewValue
End Property

Public Function getFac() As String
    getFac = mfac
End Function
Public Sub setFac(sNewValue As String)
    mfac = sNewValue
End Sub
Public Property Get Fac() As String
    Fac = mfac
End Property
Public Property Let Fac(sNewValue As String)
    mfac = sNewValue
End Property

Public Function getStatusCode() As Long
    getStatusCode = mstatuscode
End Function
Public Sub setStatusCode(lNewValue As Long)
    mstatuscode = lNewValue
End Sub
Public Property Get StatusCode() As Long
    StatusCode = mstatuscode
End Property
Public Property Let StatusCode(lNewValue As Long)
    mstatuscode = lNewValue
End Property

Public Function getStatusInfo() As String
    getStatusInfo = mstatusinfo
End Function
Public Sub setStatusInfo(sNewValue As String)
    mstatusinfo = sNewValue
End Sub
Public Property Get StatusInfo() As String
    StatusInfo = mstatusinfo
End Property
Public Property Let StatusInfo(sNewValue As String)
    mstatusinfo = sNewValue
End Property

Public Function getWinName() As String
    getWinName = mwinname
End Function
Public Sub setWinName(sNewValue As String)
    mwinname = sNewValue
End Sub
Public Function getWindowName() As String
    getWindowName = mwinname
End Function
Public Sub setWindowName(sNewValue As String)
    mwinname = sNewValue
End Sub
Public Property Get windowname() As String
    windowname = mwinname
End Property
Public Property Let windowname(sNewValue As String)
    mwinname = sNewValue
End Property

Public Function getWinGUIID() As Variant
    If IsObject(owinGUIID) Then
        Set getWinGUIID = owinGUIID
    Else
        getWinGUIID = owinGUIID
    End If
End Function
Public Sub setWinGUIID(vNewValue As Variant)
    If IsObject(vNewValue) Then
        Set owinGUIID = vNewValue
    Else
        owinGUIID = vNewValue
    End If
End Sub
Public Function getWindowGUIID() As Variant
    If IsObject(owinGUIID) Then
        Set getWindowGUIID = owinGUIID
    Else
        getWindowGUIID = owinGUIID
    End If
End Function
Public Sub setWindowGUIID(vNewValue As Variant)
    If IsObject(vNewValue) Then
        Set owinGUIID = vNewValue
    Else
        owinGUIID = vNewValue
    End If
End Sub
Public Property Get WindowGUIID() As Variant
    If IsObject(owinGUIID) Then
        Set WindowGUIID = owinGUIID
    Else
        WindowGUIID = owinGUIID
    End If
End Property
Public Property Let WindowGUIID(vNewValue As Variant)
    If IsObject(vNewValue) Then
        Set owinGUIID = vNewValue
    Else
        owinGUIID = vNewValue
    End If
End Property

Public Function getCompName() As String
    getCompName = mcompname
End Function
Public Sub setCompName(sNewValue As String)
    mcompname = sNewValue
End Sub
Public Property Get CompName() As String
    CompName = mcompname
End Property
Public Property Let CompName(sNewValue As String)
    mcompname = sNewValue
End Property

Public Function getCompGUIID() As Variant
    If IsObject(ocompGUIID) Then
        Set getCompGUIID = ocompGUIID
    Else
        getCompGUIID = ocompGUIID
    End If
End Function
Public Sub setCompGUIID(vNewValue As Variant)
    If IsObject(vNewValue) Then
        Set ocompGUIID = vNewValue
    Else
        ocompGUIID = vNewValue
    End If
End Sub
Public Property Get CompGUIID() As Variant
    If IsObject(ocompGUIID) Then
        Set CompGUIID = ocompGUIID
    Else
        CompGUIID = ocompGUIID
    End If
End Property
Public Property Let CompGUIID(vNewValue As Variant)
    If IsObject(vNewValue) Then
        Set ocompGUIID = vNewValue
    Else
        ocompGUIID = vNewValue
    End If
End Property

Public Function getTestCommand() As String
    getTestCommand = mtestcommand
End Function
Public Sub setTestCommand(sNewValue As String)
    mtestcommand = sNewValue
End Sub
Public Property Get TestCommand() As String
    TestCommand = mtestcommand
End Property
Public Property Let TestCommand(sNewValue As String)
    mtestcommand = sNewValue
End Property

Public Function getSTAFHelper() As Variant
    Set getSTAFHelper = mstafhelper
End Function
Public Sub setSTAFHelper(vNewValue As Variant)
    Set mstafhelper = vNewValue
End Sub
Public Property Get STAFHelper() As Variant
    Set getSTAFHelper = mstafhelper
End Property
Public Property Let STAFHelper(vNewValue As Variant)
    Set mstafhelper = vNewValue
End Property

Public Function getHookTRDID() As String
    getHookTRDID = mhooktrd
End Function
Public Sub setHookTRDID(sNewValue As String)
    mhooktrd = sNewValue
End Sub
Public Property Get HookTRDID() As String
    HookTRDID = mhooktrd
End Property
Public Property Let HookTRDID(sNewValue As String)
    mhooktrd = sNewValue
End Property

Public Sub reinitialize()
    setInputRecord ""
    setFileName ""
    setSeparator ""
    setLineNumber -1
    setFac ""
    setAppMapName ""
    setTestLevel ""
    setStatusCode 99
    setStatusInfo ""
    setWinName ""
    setCompName ""
    setWinGUIID Nothing
    setCompGUIID Nothing
    setTestCommand ""
End Sub


Private Sub Class_Initialize()
    mhooktrd = "SAFS/HOOK/"
    reinitialize
End Sub
