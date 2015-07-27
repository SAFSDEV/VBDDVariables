VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SAFSMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'##############################################################################
'# 
'# DDVariableStore.SAFSMonitor
'# 
'# This ActiveX Class is designed to provide a fallback mechanism for shutting 
'# down items it has been tasked to 'monitor'.  The class doesn't physically 
'# monitor STAF tools or services, but simply provides a means for some other 
'# tool to list items it has started and for the class to shut them down in the 
'# event the owning tool cannot.  For example, if the user has forced an abort 
'# of a test and consequently bypassed all the normal shutdown and cleanup 
'# activities of the test.
'# 
'# The normal process of use is that a tool will instance one of these SAFSMonitor 
'# objects.  If that tool launched STAF itself, then it would call the addSTAFMonitor 
'# function to assert such a flag.  The tool will then ADD the names of STAF 
'# services, SAFS engines, SAFS queues the tool has started via the appropriate 
'# function calls.  
'# 
'#       safsmonitor.addSTAFMonitor ()
'#       safsmonitor.addServiceMonitor (servicename)
'#       safsmonitor.addEngineMonitor  (enginename)
'#       safsmonitor.addQueueMonitor   (queuename)
'# 
'# At some point, the tool that instanced the SAFSMonitor MUST
'# call the useProcessInfo routine with a valid process name and registered 
'# STAF handle number.  Without that, the shutdown activities will not succeed
'# when the SAFSMonitor attempts them.
'# 
'#       safsmonitor.useProcessInfo (processname, stafhandle)
'#
'# As the test winds down that same tool would then REMOVE those same named STAF 
'# services, engines, and queues from the SAFSMonitor as the tool itself shut 
'# them down.  The tool would also removeSTAFMonitor if it shutdown STAF, too.
'#
'#       safsmonitor.removeEngineMonitor  (enginename)
'#       safsmonitor.removeQueueMonitor   (queuename)
'#       safsmonitor.removeServiceMonitor (servicename)
'#       safsmonitor.removeSTAFMonitor    ()
'#       safsmonitor.removeProcessInfo    ()
'# 
'# The automatic shutdown of monitored items by this class is only invoked when 
'# the tool that created an instance of this class does not REMOVE an item from 
'# the lists.  To this class, that is an indication that the tool has faltered 
'# or been interrupted by a user.  In any event, if the instanced SAFSMonitor 
'# has been told to terminate while it still has items to monitor in its list 
'# it will proceed to shutdown those down.
'#
'# Engines are shutdown via the SAFS standard "SHUTDOWN_HOOK" record dispatch.
'# Queues (if any) are shutdown via a "SHUTDOWN" message issued to the queue.
'# Services are shutdown via the STAF standard "remove service" call.
'# STAF is shutdown via the STAF "SHUTDOWN" command to the SHUTDOWN service.
'# 
'##############################################################################

Dim services As New Collection   'list of services to monitor
Dim engines As New Collection    'list of engines to monitor
Dim queues As New Collection     'list of client queues to monitor

Dim processID As String      'the process name registered with STAF
Dim handleID As Long         'the handle registered with STAF
Dim stafMonitor As Long
'Dim fileid As Integer

'##############################################################################
'# Sets the flag that this object should shutdown STAF if the flag is not 
'# reset by the time this object is terminated.
'##############################################################################
Public Sub addSTAFMonitor()
    stafMonitor = -1
    'On Error Resume Next
    'Print #fileid, "Added STAFMonitor"
End Sub

'##############################################################################
'# Removes the flag so this object will not shutdown STAF during terminate.
'##############################################################################
Public Sub removeSTAFMonitor()
    stafMonitor = 0
    'On Error Resume Next
    'Print #fileid, "Removed STAFMonitor"
End Sub

'##############################################################################
'# Adds the provided service name to the collection of services that will be 
'# shutdown when this object is terminated.
'##############################################################################
Public Sub addServiceMonitor(service As String)
    services.Add (service)
    'On Error Resume Next
    'Print #fileid, "Added Service Monitor " & service
End Sub

'##############################################################################
'# Removes the provided service name from the collection of services that will
'# be shutdown when this object is terminated.
'##############################################################################
Public Sub removeServiceMonitor(service As String)
    Dim index As Integer
    Dim match As String
        
    On Error GoTo Err
    For index = 1 To services.count
        match = services(index)
        If StrComp(service, match, vbTextCompare) = 0 Then
            services.Remove (index)
            'On Error Resume Next
            'Print #fileid, "Removed Service Monitor " & service
            Exit For
        End If
    Next
Err:
    Exit Sub
End Sub

'##############################################################################
'# Private routine to shutdown the specified service.
'##############################################################################
Private Sub terminateService(service As String)
    Dim status As Long
    On Error Resume Next
    If handleID > 0 Then
        'Print #fileid, "Terminating Service Monitor " & service
        status = submitHandle(handleID, "service", "remove service " & service)
    End If
End Sub

'##############################################################################
'# Adds the provided engine name to the collection of engines that will be 
'# shutdown when this object is terminated.
'##############################################################################
Public Sub addEngineMonitor(engine As String)
    engines.Add (engine)
    'On Error Resume Next
    'Print #fileid, "Adding Engine Monitor " & engine
End Sub

'##############################################################################
'# Removes the provided service name from the collection of engines that will be 
'# shutdown when this object is terminated.
'##############################################################################
Public Sub removeEngineMonitor(engine As String)
    Dim index As Integer
    Dim match As String
        
    On Error GoTo Err
    For index = 1 To engines.count
        match = engines(index)
        If StrComp(engine, match, vbTextCompare) = 0 Then
            engines.Remove (index)
            'On Error Resume Next
            'Print #fileid, "Removing Engine Monitor " & engine
            Exit For
        End If
    Next
Err:
    Exit Sub

End Sub

'##############################################################################
'# Private routine to shutdown the specified engine.
'##############################################################################
Private Sub terminateEngine(engine As String)
    Dim status As Long
    Dim request As String
    
    On Error Resume Next
    If handleID > 0 Then
        'Print #fileid, "Terminating Engine Monitor " & engine
        status = submitHandle(handleID, "safsvars", "set safs/hook/inputrecord value SHUTDOWN_HOOK")

        If getStafMajorVersion() < 3 Then
            request = "event " & engine & "dispatch pulse"
        Else
            request = "pulse event " & engine & "dispatch"
        End If
        status = submitHandle(handleID, "sem", request)
        
    End If
End Sub

'##############################################################################
'# Adds the provided queue name to the collection of queues that will be 
'# "shutdown" when this object is terminated.
'##############################################################################
Public Sub addQueueMonitor(queue As String)
    queues.Add (queue)
    'On Error Resume Next
    'Print #fileid, "Adding Queue Monitor " & queue
End Sub

'##############################################################################
'# Removes the provided queue name from the collection of queues that will be 
'# "shutdown" when this object is terminated.
'##############################################################################
Public Sub removeQueueMonitor(queue As String)
    Dim index As Integer
    Dim match As String
        
    On Error GoTo Err
    For index = 1 To queues.count
        match = queues(index)
        If StrComp(queue, match, vbTextCompare) Then
            queues.Remove (index)
            'On Error Resume Next
            'Print #fileid, "Removing Queue Monitor " & queue
            Exit For
        End If
    Next
Err:
    Exit Sub

End Sub

'##############################################################################
'# Private routine to shutdown the specified queue.
'##############################################################################
Private Sub terminateQueue(queue As String)
    Dim status As Long
    Dim request As String
    
    On Error Resume Next
    If handleID > 0 Then
        'Print #fileid, "Terminating Queue Monitor " & queue
        If getStafMajorVersion() < 3 Then
            request = "queue name " & queue & " message SHUTDOWN"
        Else
            request = "message SHUTDOWN " & " queue name " & queue
        End If
        
        status = submitHandle(handleID, "queue", request)
    End If
End Sub

'##############################################################################
'# A Process Name and registered STAF handle must be provided before any of the 
'# termination shutdown activities will succeed.
'##############################################################################
Public Sub useProcessInfo(aprocess As String, ahandle As Long)
    processID = aprocess
    handleID = ahandle
    'On Error Resume Next
    'Print #fileid, "Setting Process Info: " & aprocess & " handle: " & Str$(ahandle)
End Sub

'##############################################################################
'# Remove or reset whatever process name and STAF handle was previously provided.
'##############################################################################
Public Sub removeProcessInfo()
    handleID = 0
    processID = ""
    'On Error Resume Next
    'Print #fileid, "Removing Process Info"
End Sub

'##############################################################################
'# Unregister the STAF client using the STAF handle previously provided.
'##############################################################################
Private Sub unRegisterProcessInfo()
    Dim status As Long
    On Error Resume Next
    status = STAFunregister(handleID)
    'Print #fileid, "Unregistering with Process Info: status = " & Str$(status)
End Sub

'##############################################################################
'# Empty the provided collection of all stored elements.
'##############################################################################
Private Sub emptyCollection(collect As Collection)
    On Error Resume Next
    Do While (collect.count > 0)
        collect.Remove (1)
    Loop
    'Print #fileid, "Emptying the Collection"
End Sub

Private Sub Class_Initialize()
    'On Error Resume Next
    'fileid = FreeFile
    'Open "c:\safs\data\safsmonitor.txt" For Output As fileid
End Sub

'##############################################################################
'# We will shutdown monitored engines, queues, and services (in that order) if 
'# they are still present in our collections lists.  We will then shutdown 
'# STAF if the flag to do so is still set.
'##############################################################################
Private Sub Class_Terminate()
    Dim engine As Variant
    Dim queue As Variant
    Dim service As Variant
    Dim status As Long
    
    On Error GoTo que
    For Each engine In engines
        terminateEngine (engine)
    Next
    emptyCollection engines
    'Print #fileid, "That was for ENGINES"
que:
    On Error GoTo serve
    For Each queue In queues
        terminateQueue (queue)
    Next
    emptyCollection queues
    'Print #fileid, "That was for QUEUES"

serve:
    On Error GoTo done
    For Each service In services
        terminateService (service)
    Next
    emptyCollection services
    'Print #fileid, "That was for SERVICES"
    
done:
    On Error Resume Next
    If stafMonitor = -1 Then
        'Print #fileid, "SHUTDOWN of STAF"
        status = submitHandle(handleID, "shutdown", "shutdown")
    End If
    
    On Error Resume Next
    If handleID > 0 Then
        unRegisterProcessInfo
        removeProcessInfo
    End If
    'Close fileid
    Exit Sub
End Sub
