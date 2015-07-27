VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "STAFUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private processID As String      'the process name registered with STAF
Private handleID As Long         'the handle registered with STAF
Private machine  As String       'the default target machine for STAF -- normally "local"


'##############################################################################
'#
'#  Sub useProcessInfo (aprocess as String, ahandle As Long)
'#
'#  DESCRIPTION:
'#
'#     A previously registered process name and handle can be provided for this
'#     instance of STAFUtilities.  Thus, assuming the handle is valid, we do not
'#     have to RegisterProcess with STAF.
'#
'#  PARAMETERS:
'#
'#     aprocess  The process name previously registered with STAF.
'#
'#     ahandle   The handleID for the process previously registered with STAF.
'#
'#  ERRORS:
'#
'#     (none)
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#      DEC 15, 2009    Carl Nagle  Fixed STAF V3 implementation in WaitEvent
'#
'##############################################################################
Public Sub useProcessInfo(ByVal aprocess As String, ByVal ahandle As Long)
    processID = aprocess
    handleID = ahandle
    'On Error Resume Next
    'Print #fileid, "Setting Process Info: " & aprocess & " handle: " & Str$(ahandle)
End Sub


'##############################################################################
'#
'#  Function RegisterNewProcess (aprocess as String) As Long
'#
'#  DESCRIPTION:
'#
'#     Attempt to call RegisterProcess with the provided aprocess name and our
'#     local handleID to receive the registered STAF handle value.
'#
'#  PARAMETERS:
'#
'#     aprocess  The process name to be registered with STAF.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#      Sets our local handleID to the STAF handle value registered.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function RegisterNewProcess(ByVal aprocess As String) As Long
    processID = aprocess
    RegisterNewProcess = registerProcess(aprocess, handleID)
End Function


'##############################################################################
'#
'#  Sub removeProcessInfo ()
'#
'#  DESCRIPTION:
'#
'#     Remove or reset whatever process name and STAF handle were previously set.
'#
'#  PARAMETERS:
'#
'#     (none)
'#
'#  ERRORS:
'#
'#     (none)
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Sub removeProcessInfo()
    handleID = 0
    processID = ""
    'On Error Resume Next
    'Print #fileid, "Removing Process Info"
End Sub


'##############################################################################
'#
'#  Function unRegisterProcess () As Long
'#
'#  DESCRIPTION:
'#
'#     Attempt to unregister our stored handleID from STAF.
'#
'#  PARAMETERS:
'#
'#     (none)
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function unRegisterProcess() As Long
    Dim status As Long
    On Error Resume Next
    unRegisterProcess = CLng(STAFunregister(handleID))
    'Print #fileid, "Unregistering with Process Info: status = " & Str$(status)
End Function


'##############################################################################
'#
'#  Function getHandleID () As Long
'#
'#  DESCRIPTION:
'#
'#     Retrieves the current value of our stored handleID.
'#     Usually any handleID > 1 means we are currently registered with STAF.
'#
'#  PARAMETERS:
'#
'#     (none)
'#
'#  RETURNS:
'#
'#      The current value of our stored handleID
'#      (which may or may not be registered.)
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function getHandleID() As Long
    getHandleID = handleID
End Function


'##############################################################################
'#
'#  Function getProcessID () As String
'#
'#  DESCRIPTION:
'#
'#     Retrieves the current value of our stored processID.
'#
'#  PARAMETERS:
'#
'#     (none)
'#
'#  RETURNS:
'#
'#      The current value of our stored processID
'#      (which may or may not be registered.)
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function getProcessID() As String
    getProcessID = processID
End Function


'##############################################################################
'#
'#  Function getMachineID () As String
'#
'#  DESCRIPTION:
'#
'#     Retrieves the current value of the default machine to use for STAF calls.
'#     By default we use "local" for this value.  This is the "where" argument
'#     for many STAF calls.
'#
'#  PARAMETERS:
'#
'#     (none)
'#
'#  RETURNS:
'#
'#      The current value of our target machineID
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
Public Function getMachineID() As String
    getMachineID = machine
End Function


'##############################################################################
'#
'#  Public Property Get location () As String
'#
'#  DESCRIPTION:
'#
'#     Retrieves the current value of the default machine to use for STAF calls.
'#     By default we use "local" for this value.  This is the "where" argument
'#     for many STAF calls.
'#
'#  PARAMETERS:
'#
'#     (none)
'#
'#  RETURNS:
'#
'#      The current value of our target machineID
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
Public Property Get location() As String
    location = machine
End Property


'##############################################################################
'#
'#  Public Property Let location (slocation as String)
'#
'#  DESCRIPTION:
'#
'#     Sets the new value for the default machine to use for STAF calls.
'#     By default we use "local" for this value.  This is the "where" argument
'#     for many STAF calls.
'#
'#  PARAMETERS:
'#
'#     slocation   the new target machine to use for STAF calls.
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
Public Property Let location(ByVal slocation As String)
    machine = slocation
End Property


'##############################################################################
'#
'#  Function getMachineID () As String
'#
'#  DESCRIPTION:
'#
'#     Retrieves the current value of the default machine to use for STAF calls.
'#     By default we use "local" for this value.  This is the "where" argument
'#     for many STAF calls.
'#
'#  PARAMETERS:
'#
'#     (none)
'#
'#  RETURNS:
'#
'#      The current value of our target machineID
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
Public Sub setMachineID(ByVal sNewValue As String)
    machine = sNewValue
End Sub


'##############################################################################
'#
'#  Function submitSTAFRequest (where As String, service As String,
'#                              amessage As String, astafresult As STAFResult
'#                              ) As Long
'#
'#  DESCRIPTION:
'#
'#     Submit a STAF request to a local or remote machine using our previously
'#     registered handleID.  The results of the call will be stored in the
'#     provided STAFResult object.
'#
'#     Note: VBScript cannot pass Object types.  VBScript can only pass objects
'#     stored as Variants.  Thus, VBScript and other VB users may need to use the
'#     submitSTAFVariantRequest version of this method.
'#
'#  PARAMETERS:
'#
'#     where        the target machine to receive the STAF request.
'#                  Usually this is "local" for the local machine.
'#
'#     service      the running STAF service to receive the request.
'#
'#     amessage     the message or command to issue to the service.
'#
'#     astafresult  the STAFResult object to receive the results of the call.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function submitSTAFRequest(ByVal where As String, ByVal service As String, _
                                  ByVal amessage As String, astafresult As STAFResult) As Long
    If astafresult Is Nothing Then Set astafresult = New STAFResult
    submitSTAFRequest = submitSTAFResultRequest(handleID, where, service, amessage, astafresult)
End Function


'##############################################################################
'#
'#  Function submitSTAFVariantRequest (where As String, service As String,
'#                                     amessage As String, astafresult As Variant
'#                                     ) As Long
'#
'#  DESCRIPTION:
'#
'#     Submit a STAF request to a local or remote machine using our previously
'#     registered handleID.  The results of the call will be stored in the
'#     provided Variant holding a STAFResult object.
'#
'#     Note: VBScript cannot pass Object types.  VBScript can only pass objects
'#     stored as Variants.  Thus, VBScript must pass a Variant that has been
'#     initialized with a STAFResult object.
'#
'#  PARAMETERS:
'#
'#     where        the target machine to receive the STAF request.
'#                  Usually this is "local" for the local machine.
'#
'#     service      the running STAF service to receive the request.
'#
'#     amessage     the message or command to issue to the service.
'#
'#     astafresult  the Variant STAFResult object to receive the results of the call.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function submitSTAFVariantRequest(ByVal where As String, ByVal service As String, _
                                         ByVal amessage As String, astafresult As Variant) As Long
    If astafresult Is Nothing Then Set astafresult = New STAFResult
    submitSTAFVariantRequest = submitSTAFResultVariant(handleID, where, service, amessage, astafresult)
End Function


'##############################################################################
'#
'#  Function getVariable (sVarName as String, sVarValue as String) As Long
'#
'#  DESCRIPTION:
'#
'#     Retrieves the current value of the SAFSVARS sVarName provided from the
'#     STAF location we currently have set -- usually "local".
'#
'#  PARAMETERS:
'#
'#     sVarName   The name of the variable to retrieve from SAFSVARS.
'#
'#     sVarValue  The argument to receive the value of the variable on success.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function getVariable(ByVal sVarName As String, oResult As Variant) As Long
    Dim amessage As String
    Dim status As Long
    Dim resulttype As String
    Dim validresult As Boolean
    
    validresult = False
    resulttype = TypeName(oResult)
    If resulttype = "STAFResult" Then validresult = True
    amessage = "GET " & lenMark(sVarName)
    status = submitWhereRequest(handleID, machine, "SAFSVARS", amessage)
    getVariable = status
    If validresult Then
        oResult.rc = status
        oResult.result = resultStr
    End If
End Function


'##############################################################################
'#
'#  Function setVariable (sVarName as String, sVarValue as String) As Long
'#
'#  DESCRIPTION:
'#
'#     Sets a new value for the SAFSVARS sVarName provided from the
'#     STAF location we currently have set -- usually "local".
'#
'#  PARAMETERS:
'#
'#     sVarName   The name of the variable to set in SAFSVARS.
'#
'#     sVarValue  The value to set for this variable.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function setVariable(ByVal sVarName As String, ByVal sVarValue As String) As Long
    Dim amessage As String
    amessage = "SET " & lenMark(sVarName) & " VALUE " & lenMark(sVarValue)
    setVariable = submitWhereRequest(handleID, machine, "SAFSVARS", amessage)
End Function

  
'##############################################################################
'#
'#  Function getQueueMessage (sQueueName as String, sQueueMessage as String,
'#                            lSeconds As Long) As Long
'#
'#  DESCRIPTION:
'#
'#     Retrieves a message from the named STAF message queue.
'#     Uses the STAF location we currently have set -- usually "local".
'#
'#  PARAMETERS:
'#
'#     sQueueName    The name of the message queue to retrieve a message from.
'#                   If no name is provided then the default queue for our registered
'#                   process is used.
'#
'#     sQueueMessage  The argument to receive the value of the message if successful.
'#                    This string contains all the components of the STAFQueueMessage
'#                    including Priority, Timestamp, Machine, Name, Handle, and
'#                    the message.
'#
'#     lSeconds  the number of seconds to wait for a message in the queue.
'#               If < 0  then the wait is indefinite.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function getQueueMessage(ByVal sQueueName As String, oResult As Variant, lSeconds As Long) As Long
    Dim amessage As String
    Dim resulttype As String
    Dim validresult As Boolean
    Dim status As Long
    
    resulttype = TypeName(oResult)
    validresult = False
    If resulttype = "STAFResult" Then validresult = True
    
    amessage = "GET"
    If Len(sQueueName) > 0 Then
        amessage = amessage & " NAME " & lenMark(sQueueName)
    End If
    amessage = amessage & " WAIT"
    If lSeconds >= 0 Then
        amessage = amessage & " " & Trim$(Str$(1000 * lSeconds))
    End If
    
    status = submitWhereRequest(handleID, machine, "QUEUE", amessage)
    
    getQueueMessage = status
    If validresult Then
        oResult.rc = status
        oResult.result = resultStr
    End If
End Function


'##############################################################################
'#
'#  Function sendQueueMessage (sQueueName as String, sQueueMessage as String) As Long
'#
'#  DESCRIPTION:
'#
'#     Sends a message to the named STAF message queue.
'#     Uses the STAF location we currently have set -- usually "local".
'#
'#  PARAMETERS:
'#
'#     sQueueName    The name of the message queue to recieve the message.
'#
'#     sQueueMessage  The message to send.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function sendQueueMessage(ByVal sQueueName As String, ByVal sQueueMessage As String) As Long
    Dim amessage As String
    amessage = "QUEUE NAME " & lenMark(sQueueName) & " MESSAGE " & lenMark(sQueueMessage)
    sendQueueMessage = submitWhereRequest(handleID, machine, "QUEUE", amessage)
End Function


'##############################################################################
'#
'#  Function pulseEvent (sEvent as String) As Long
'#
'#  DESCRIPTION:
'#
'#     Pulse a STAF SEM service event.
'#     Uses the STAF location we currently have set -- usually "local".
'#
'#  PARAMETERS:
'#
'#     sEvent    The name of the event in the STAF SEM service to pulse.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function pulseEvent(ByVal sEvent As String) As Long
    Dim request As String
    
    If getStafMajorVersion() < 3 Then
        request = "EVENT " & sEvent & " PULSE"
    Else
        request = "PULSE EVENT " & sEvent
    End If
    pulseEvent = submitWhereRequest(handleID, machine, "SEM", request)
End Function


'##############################################################################
'#
'#  Function postEvent (sEvent as String) As Long
'#
'#  DESCRIPTION:
'#
'#     Post a STAF SEM service event.
'#     Uses the STAF location we currently have set -- usually "local".
'#
'#  PARAMETERS:
'#
'#     sEvent    The name of the event in the STAF SEM service to post.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function postEvent(ByVal sEvent As String) As Long
    Dim request As String
    
    If getStafMajorVersion() < 3 Then
        request = "EVENT " & sEvent & " POST"
    Else
        request = "POST EVENT " & sEvent
    End If
    
    postEvent = submitWhereRequest(handleID, machine, "SEM", request)
End Function
  
  
'##############################################################################
'#
'#  Function resetEvent (sEvent as String) As Long
'#
'#  DESCRIPTION:
'#
'#     Reset a STAF SEM service event.
'#     Uses the STAF location we currently have set -- usually "local".
'#
'#  PARAMETERS:
'#
'#     sEvent    The name of the event in the STAF SEM service to reset.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function resetEvent(ByVal sEvent As String) As Long
    Dim request As String
    
    If getStafMajorVersion() < 3 Then
        request = "EVENT " & sEvent & " RESET"
    Else
        request = "RESET EVENT " & sEvent
    End If
    
    resetEvent = submitWhereRequest(handleID, machine, "SEM", request)
End Function

  
'##############################################################################
'#
'#  Function resetHookEvents (sProcessName as String) As Long
'#
'#  DESCRIPTION:
'#
'#     Reset all STAF SEM events for a particular process (engine).
'#     Uses the STAF location we currently have set -- usually "local".
'#
'#     The list of events reset by this are:
'#
'#         sProcessName + "Start"
'#         sProcessName + "Ready"
'#         sProcessName + "Dispatch"
'#         sProcessName + "Running"
'#         sProcessName + "Results"
'#         sProcessName + "Done"
'#         sProcessName + "Shutdown"
'#
'#  PARAMETERS:
'#
'#     sEvent    The name of the event in the STAF SEM service to reset.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function resetHookEvents(ByVal sProcessName As String) As Long
    Dim status As Long
    status = resetEvent(sProcessName & "Start")
    If status = STAF_Ok Then status = resetEvent(sProcessName & "Ready")
    If status = STAF_Ok Then status = resetEvent(sProcessName & "Dispatch")
    If status = STAF_Ok Then status = resetEvent(sProcessName & "Running")
    If status = STAF_Ok Then status = resetEvent(sProcessName & "Results")
    If status = STAF_Ok Then status = resetEvent(sProcessName & "Done")
    If status = STAF_Ok Then status = resetEvent(sProcessName & "Shutdown")
    resetHookEvents = status
End Function

  
'##############################################################################
'#
'#  Function waitEvent (sEvent as String, lSeconds As Long) As Long
'#
'#  DESCRIPTION:
'#
'#     Wait for a STAF SEM service event.
'#     Uses the STAF location we currently have set -- usually "local".
'#
'#  PARAMETERS:
'#
'#     sEvent    The name of the event in the STAF SEM service to wait for.
'#
'#     lSeconds  the number of seconds to wait for the event.
'#               If lSeconds < 0 then the wait is indefinite.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#      DEC 15, 2009    Carl Nagle Fixed STAF V3 implementation
'#
'##############################################################################
Public Function waitEvent(ByVal sEvent As String, ByVal lSeconds As Long) As Long
    Dim amessage As String
    
    If getStafMajorVersion() < 3 Then
        amessage = "EVENT " & sEvent & " WAIT"
        If (lSeconds > -1) Then
            amessage = amessage & " " & Trim$(Str$(1000 * lSeconds))
        End If
    Else
        amessage = "WAIT EVENT " & sEvent
        If (lSeconds > -1) Then
            amessage = amessage & " TIMEOUT " & Trim$(Str$(lSeconds)) & "s"
        End If
    End If
    
    waitEvent = submitWhereRequest(handleID, machine, "SEM", amessage)
End Function


'##############################################################################
'#
'#  Function waitMutex(sMutexName, lSeconds as Long) As Long
'#
'#  DESCRIPTION:
'#
'#      WAIT for a single STAF SEMaphore MUTEX
'#
'#  PARAMETERS:
'#
'#      sMutexName  the name of the mutex
'#
'#      lSeconds    the number of seconds to wait ( < 0 = wait indefinitely)
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function waitMutex(ByVal sMutexName As String, ByVal lSeconds As Long) As Long
    Dim amessage As String
    
    If getStafMajorVersion() < 3 Then
        amessage = "MUTEX " & sMutexName & " REQUEST"
    Else
        amessage = "REQUEST MUTEX " & sMutexName
    End If
    
    If lSeconds >= 0 Then
        amessage = amessage & " " & Trim$(Str$(1000 * lSeconds))
    End If
    waitMutex = submitWhereRequest(handleID, machine, "SEM", amessage)
End Function

'##############################################################################
'#
'#  Function releaseMutex(sMutexName) As Long
'#
'#  DESCRIPTION:
'#
'#      Release a single STAF SEMaphore MUTEX
'#
'#  PARAMETERS:
'#
'#      sMutexName  the name of the mutex
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function releaseMutex(ByVal sMutexName As String) As Long
    Dim request As String

    If getStafMajorVersion() < 3 Then
        request = "MUTEX " & sMutexName & " RELEASE"
    Else
        request = "RELEASE MUTEX " & sMutexName
    End If
    
    releaseMutex = submitWhereRequest(handleID, machine, "SEM", request)
End Function


'##############################################################################
'#
'#  Function getNextHookTestEvent (sProcessName as String) As Long
'#
'#  DESCRIPTION:
'#
'#      Perform the complete engine-side protocol for processing a test record
'#      dispatched from a driver.
'#
'#          waitEvent  sProcessName+"Dispatch";
'#          resetEvent sProcessName+"Ready";
'#          postEvent  sProcessName+"Running"
'#
'#      waitEvent uses indefinite timeout (forever).
'#
'#  PARAMETERS:
'#
'#     sProcessName  The name of the process (engine) whose events we will do.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function getNextHookTestEvent(ByVal sProcessName As String) As Long
    Dim status As Long
    status = waitEvent(sProcessName & "Dispatch", -1)
    If status = STAF_Ok Then status = resetEvent(sProcessName & "Ready")
    If status = STAF_Ok Then status = postEvent(sProcessName & "Running")
    getNextHookTestEvent = status
End Function


'##############################################################################
'#
'#  Function postNextHookTestEventVariant(sProcessName as String,
'#                                        oTRDData as Variant) as Long
'#
'#  DESCRIPTION:
'#
'#      Perform the complete driver-side protocol for dispatching an engine to process
'#      a test record.
'#
'#          waitEvent sProcessName+"Ready"
'#          waitMutex sTRDRoot+"TRD"
'#          set SAFSVARS test record data
'#          postEvent sProcessName+"Dispatch"
'#          waitEvent sProcessName+"Results"
'#          get SAFSVARS test record data results (statuscode, statusinfo)
'#          pulseEvent sProcessName+"Done"
'#          releaseMutex sTRDRoot+"TRD"
'#
'#  PARAMETERS:
'#
'#      sProcessName   the process name (engine) to dispatch for testing.
'#
'#      oTRDData       the TestRecordData object passed in as a Variant.
'#                     This is for VBScript which can only pass objects as Variants.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function postNextHookTestEventVariant(ByVal sProcessName As String, _
                                             oTRDData As Variant) As Long
    Dim status As Long
    status = waitEvent(sProcessName & "Ready", -1)
    status = waitMutex(oTRDData.getHookTRDID() & "TRD", -1)
    
    status = setTestRecordDataVariant(oTRDData)

    'status = pulseEvent(sProcessName & "Dispatch")
    status = postEvent(sProcessName & "Dispatch")
    status = waitEvent(sProcessName & "Running", -1)
    status = resetEvent(sProcessName & "Dispatch")
    
    status = waitEvent(sProcessName & "Results", -1)
        
    status = getTestRecordResultsVariant(oTRDData)

    status = pulseEvent(sProcessName & "Done")
    status = releaseMutex(oTRDData.getHookTRDID() & "TRD")
    
    postNextHookTestEventVariant = status
End Function
    
    
'##############################################################################
'#
'#  Function setTestRecordDataVariant(sTRDRoot As String, oTRDData as Variant) As Long
'#
'#  DESCRIPTION:
'#
'#      Set the SAFSVARS Test Record Data for the given sTRDRoot.
'#      The most common trd_root is "SAFS/HOOK/", which is used by all official
'#      drivers and engines.  The routine does not deal with the TRD Mutex for the
'#      sTRDRoot provided.  It assumes the caller is handling the mutex.
'#
'#      Sets the following record data in SAFSVARS from the oTRDData object:
'#
'#      sTRDRoot+"inputrecord"
'#      sTRDRoot+"separator"
'#      sTRDRoot+"filename"
'#      sTRDRoot+"linenumber"
'#      sTRDRoot+"testlevel"
'#      sTRDRoot+"appmapname"
'#      sTRDRoot+"fac"
'#      sTRDRoot+"statuscode"
'#      sTRDRoot+"statusinfo"
'#
'#  PARAMETERS:
'#
'#      sTRDRoot    The root prefix used for SAFSVARS storage for this data.
'#                  The most common root prefix is "SAFS/HOOK/"
'#
'#      oTRDData    The Variant TestRecordData object containing the data.
'#                  This is for VBScript which can only pass objects as Variants.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function setTestRecordDataVariant(oTRDData As Variant) As Long
    Dim status As Long
    status = setVariable(oTRDData.getHookTRDID() & SAFS_TRD_INPUTRECORD, oTRDData.getInputRecord())
    status = setVariable(oTRDData.getHookTRDID() & SAFS_TRD_LINENUMBER, oTRDData.getLineNumber())
    status = setVariable(oTRDData.getHookTRDID() & SAFS_TRD_SEPARATOR, oTRDData.getSeparator())
    status = setVariable(oTRDData.getHookTRDID() & SAFS_TRD_TESTLEVEL, oTRDData.getTestLevel())
    status = setVariable(oTRDData.getHookTRDID() & SAFS_TRD_FILENAME, oTRDData.getFileName())
    status = setVariable(oTRDData.getHookTRDID() & SAFS_TRD_APPMAPNAME, oTRDData.getAppMapName())
    status = setVariable(oTRDData.getHookTRDID() & SAFS_TRD_FAC, oTRDData.getFac())
    status = setVariable(oTRDData.getHookTRDID() & SAFS_TRD_STATUSCODE, oTRDData.getStatusCode())
    status = setVariable(oTRDData.getHookTRDID() & SAFS_TRD_STATUSINFO, oTRDData.getStatusInfo())
    setTestRecordDataVariant = status
End Function


'##############################################################################
'#
'#  Function sendTestRecordResultsVariant(oTRDData as Variant) As Long
'#
'#  DESCRIPTION:
'#
'#      Set a results subset of the SAFSVARS Test Record Data for the given sTRDRoot.
'#      The most common trd_root is "SAFS/HOOK/", which is used by all official
'#      drivers and engines.  The routine does not deal with the TRD Mutex for the
'#      sTRDRoot provided.  It assumes the caller is handling the mutex.
'#
'#      Sets only the following record data in SAFSVARS from the oTRDData object:
'#
'#      sTRDRoot+"statuscode"
'#      sTRDRoot+"statusinfo"
'#
'#  PARAMETERS:
'#
'#      oTRDData    The Variant TestRecordData object containing the data.
'#                  This is for VBScript which can only pass objects as Variants.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function sendTestRecordResultsVariant(oTRDData As Variant) As Long
    Dim status As Long
    status = setVariable(oTRDData.getHookTRDID() & SAFS_TRD_STATUSCODE, oTRDData.getStatusCode())
    status = setVariable(oTRDData.getHookTRDID() & SAFS_TRD_STATUSINFO, oTRDData.getStatusInfo())
    sendTestRecordResultsVariant = status
End Function


'##############################################################################
'#
'#  Function getTestRecordDataVariant(oTRDData as Variant) As Long
'#
'#  DESCRIPTION:
'#
'#      Get the SAFSVARS Test Record Data into the oTRDData Variant.
'#      The most common trd_root is "SAFS/HOOK/", which is used by all official
'#      drivers and engines.  The routine does not deal with the TRD Mutex for the
'#      sTRDRoot provided.  It assumes the caller is handling the mutex.
'#
'#      Get the following record data from SAFSVARS into the oTRDData object:
'#
'#      sTRDRoot+"inputrecord"
'#      sTRDRoot+"separator"
'#      sTRDRoot+"filename"
'#      sTRDRoot+"linenumber"
'#      sTRDRoot+"testlevel"
'#      sTRDRoot+"appmapname"
'#      sTRDRoot+"fac"
'#      sTRDRoot+"statuscode"
'#      sTRDRoot+"statusinfo"
'#
'#  PARAMETERS:
'#
'#      oTRDData    The Variant TestRecordData object to receive the data.
'#                  This is for VBScript which can only pass objects as Variants.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function getTestRecordDataVariant(oTRDData As Variant) As Long
    Dim oResult As STAFResult
    Dim status As Long
    On Error Resume Next
    Set oResult = New STAFResult
    status = getVariable(oTRDData.getHookTRDID() & SAFS_TRD_INPUTRECORD, oResult)
    oTRDData.setInputRecord (oResult.result)
    oResult.reset
    status = getVariable(oTRDData.getHookTRDID() & SAFS_TRD_LINENUMBER, oResult)
    oTRDData.setLineNumber (CLng(oResult.result))
    oResult.reset
    status = getVariable(oTRDData.getHookTRDID() & SAFS_TRD_SEPARATOR, oResult)
    oTRDData.setSeparator (oResult.result)
    oResult.reset
    status = getVariable(oTRDData.getHookTRDID() & SAFS_TRD_TESTLEVEL, oResult)
    oTRDData.setTestLevel (oResult.result)
    oResult.reset
    status = getVariable(oTRDData.getHookTRDID() & SAFS_TRD_FILENAME, oResult)
    oTRDData.setFileName (oResult.result)
    oResult.reset
    status = getVariable(oTRDData.getHookTRDID() & SAFS_TRD_APPMAPNAME, oResult)
    oTRDData.setAppMapName (oResult.result)
    oResult.reset
    status = getVariable(oTRDData.getHookTRDID() & SAFS_TRD_FAC, oResult)
    oTRDData.setFac (oResult.result)
    oResult.reset
    status = getVariable(oTRDData.getHookTRDID() & SAFS_TRD_STATUSCODE, oResult)
    oTRDData.setStatusCode (CLng(oResult.result))
    oResult.reset
    status = getVariable(oTRDData.getHookTRDID() & SAFS_TRD_STATUSINFO, oResult)
    oTRDData.setStatusInfo (oResult.result)
    Set oResult = Nothing
    getTestRecordDataVariant = status
End Function
    
    
'##############################################################################
'#
'#  Function getTestRecordResultsVariant(oTRDData as Variant) As Long
'#
'#  DESCRIPTION:
'#
'#      Get the results subset of the SAFSVARS Test Record Data into the oTRDData Variant.
'#      The most common trd_root is "SAFS/HOOK/", which is used by all official
'#      drivers and engines.  The routine does not deal with the TRD Mutex for the
'#      sTRDRoot provided.  It assumes the caller is handling the mutex.
'#
'#      Get only the following record data from SAFSVARS into the oTRDData object:
'#
'#      sTRDRoot+"statuscode"
'#      sTRDRoot+"statusinfo"
'#
'#  PARAMETERS:
'#
'#      oTRDData    The Variant TestRecordData object to receive the data.
'#                  This is for VBScript which can only pass objects as Variants.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function getTestRecordResultsVariant(oTRDData As Variant) As Long
    Dim oResult As STAFResult
    Dim status As Long
    Set oResult = New STAFResult
    On Error Resume Next
    status = getVariable(oTRDData.getHookTRDID() & SAFS_TRD_STATUSCODE, oResult)
    oTRDData.setStatusCode (CLng(oResult.result))
    oResult.reset
    status = getVariable(oTRDData.getHookTRDID() & SAFS_TRD_STATUSINFO, oResult)
    oTRDData.setStatusInfo (oResult.result)
    getTestRecordResultsVariant = status
    Set oResult = Nothing
End Function


'##############################################################################
'#
'#  Function setHookTestResultsEvents (sProcessName as String, lSeconds as Long) As Long
'#
'#  DESCRIPTION:
'#
'#      Perform the engine-side protocol for submitting the results from
'#      processing a test record dispatched from a driver.
'#
'#          postEvent  sProcessName+"Results"
'#          waitEvent  sProcessName+"Done"
'#          resetEvent sProcessName+"Results"
'#
'#      waitEvent uses indefinite timeout (forever).
'#
'#  PARAMETERS:
'#
'#     sProcessName  The name of the process (engine) whose events we will do.
'#
'#     lSeconds      The number of seconds to wait for the driver response.
'#                   This will never be indefinite.  If < 0 then 0.
'#                   This is to prevent lockups from lost drivers.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function setHookTestResultsEvents(ByVal sProcessName As String, _
                                   ByVal lSeconds As Long) As Long
    Dim status As Long
    If lSeconds < 0 Then lSeconds = 0
    status = postEvent(sProcessName & "Results")
    status = waitEvent(sProcessName & "Done", lSeconds)
    status = resetEvent(sProcessName & "Results")
    setHookTestResultsEvents = status
End Function


'##############################################################################
'#
'#  Function getAppMapItem (appMapID as String, section as String,
'#                          item as String, oResult as Variant) As Long
'#
'#  DESCRIPTION:
'#
'#     Retrieves the value of the SAFSMAPS item in the given map from the given
'#     section.  Returns the value in the sValue argument.
'#
'#  PARAMETERS:
'#
'#     appMapID   The ID of the open App Map in SAFSMAPS to retrieve from.
'#
'#     section    The section in the app map to search.
'#
'#     item       The item in the section to find.
'#
'#     oResult    a STAFResult variant to receive the results.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function getAppMapItem(ByVal appmapid As String, _
                              ByVal section As String, ByVal Item As String, _
                              oResult As Variant) As Long
    Dim amessage As String
    Dim trdtype As String
    Dim trdvalid As Boolean
    Dim status As Long
    
    trdvalid = False
    trdtype = TypeName(oResult)
    If trdtype = "STAFResult" Then trdvalid = True
    amessage = "GETITEM " & lenMark(appmapid) & " SECTION " & lenMark(section) + " ITEM " + lenMark(Item)
    status = submitWhereRequest(handleID, machine, "SAFSMAPS", amessage)
    getAppMapItem = status
    If trdvalid Then
        oResult.rc = status
        oResult.result = resultStr
    End If
End Function
    
    
'##############################################################################
'#
'#  Function logMessage (sFac as String, sMessage as String,
'#                       sDesc as String, lType as Long, oResult as Variant) As Long
'#
'#  DESCRIPTION:
'#
'#     Log a message to the given logFac in SAFSLOGS.
'#
'#  PARAMETERS:
'#
'#     sFac       The ID of the open LogFac in SAFSLOGS to receive the message.
'#
'#     sMessage   The message to log to the logFac in SAFSLOGS.
'#
'#     sDesc      Additional detail info to log with the message.
'#                If this is empty it will not be logged.
'#
'#     lType      The Type of the message being logged:
'#
'#                DEBUG_MESSAGE 7
'#                GENERIC_MESSAGE 0
'#                FAILED_MESSAGE 1024
'#                PASSED_MESSAGE 2048
'#                WARNING_MESSAGE 4096
'#                FAILED_OK_MESSAGE 1025
'#                WARNING_OK_MESSAGE 4097
'#                START_LOGGING 16
'#                STOP_LOGGING 32
'#                START_DATATABLE 3
'#                END_DATATABLE 512
'#                START_PROCEDURE 1
'#                END_PROCEDURE 2
'#                START_TESTCASE 4
'#                END_TESTCASE 8
'#                START_SUITE 5
'#                END_SUITE 6
'#                START_CYCLE 9
'#                END_CYCLE 10
'#                START_COUNTER 11
'#                END_COUNTER 12
'#                SUSPEND_STATUS_COUNTS 13
'#                RESUME_STATUS_COUNTS 14
'#                STATUS_REPORT_START 17
'#                STATUS_REPORT_RECORDS 18
'#                STATUS_REPORT_SKIPPED 19
'#                STATUS_REPORT_TESTS 20
'#                STATUS_REPORT_TEST_PASSES 21
'#                STATUS_REPORT_TEST_WARNINGS 22
'#                STATUS_REPORT_TEST_FAILURES 23
'#                STATUS_REPORT_GENERAL_WARNINGS 24
'#                STATUS_REPORT_GENERAL_FAILURES 25
'#                STATUS_REPORT_IO_FAILURES 26
'#                STATUS_REPORT_FUTURE_RESERVED 27
'#                STATUS_REPORT_FUTURE_RESERVED 28
'#                STATUS_REPORT_FUTURE_RESERVED 29
'#                STATUS_REPORT_FUTURE_RESERVED 30
'#                STATUS_REPORT_END 31
'#                START_REQUIREMENT 64
'#                END_REQUIREMENT 128
'#                SKIPPED_TEST_MESSAGE 256
'#                CUSTOM_MESSAGE 10000
'#
'#                Consult the documentation for SAFSLOGS for additional MSGTYPE
'#                values (if any).
'#
'#      oResult   STAFResult Variant to receive the string result of the call.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#  ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#
'##############################################################################
Public Function logMessage(ByVal sFac As String, ByVal sMessage As String, _
                           ByVal sDescription As String, ByVal lType As Long, _
                           oResult As Variant) As Long
    Dim amessage As String
    Dim trdtype As String
    Dim trdvalid As Boolean
    Dim status As Long
    
    trdvalid = False
    trdtype = TypeName(oResult)
    If trdtype = "STAFResult" Then trdvalid = True
    
    amessage = "LOGMESSAGE " & lenMark(sFac) & " MESSAGE " & lenMark(sMessage)
    If Len(sDescription) > 0 Then
        amessage = amessage & " DESCRIPTION " & lenMark(sDescription)
    End If
    If lType < 0 Then lType = 0
    amessage = amessage & " MSGTYPE " & Trim$(Str$(lType))
    status = submitWhereRequest(handleID, machine, "SAFSLOGS", amessage)
    logMessage = status
    If trdvalid Then
        oResult.rc = status
        oResult.result = resultStr
    End If
End Function

Public Function logGenericMessage(ByVal sFac As String, ByVal sMessage As String, _
                                  ByVal sDescription As String, oResult As Variant) As Long
    logGenericMessage = logMessage(sFac, sMessage, sDescription, CLng(0), oResult)
End Function

Public Function logDebugMessage(ByVal sFac As String, ByVal sMessage As String, _
                                  ByVal sDescription As String, oResult As Variant) As Long
    logDebugMessage = logMessage(sFac, sMessage, sDescription, CLng(7), oResult)
End Function

Public Function logFailedMessage(ByVal sFac As String, ByVal sMessage As String, _
                                  ByVal sDescription As String, oResult As Variant) As Long
    logFailedMessage = logMessage(sFac, sMessage, sDescription, CLng(1024), oResult)
End Function

Public Function logPassedMessage(ByVal sFac As String, ByVal sMessage As String, _
                                  ByVal sDescription As String, oResult As Variant) As Long
    logPassedMessage = logMessage(sFac, sMessage, sDescription, CLng(2048), oResult)
End Function

Public Function logWarningMessage(ByVal sFac As String, ByVal sMessage As String, _
                                  ByVal sDescription As String, oResult As Variant) As Long
    logWarningMessage = logMessage(sFac, sMessage, sDescription, CLng(4096), oResult)
End Function


Private Sub Class_Initialize()
    processID = ""
    handleID = 0
    machine = "local"

    'Un-comment the following line to log debug message to "C:\Debug_ddvariable.txt" and "C:\Debug_ddvariableBinary.txt"
    'DEBUG_ENABLED = True
    openDebugFile
    
End Sub

Private Sub Class_Terminate()

    On Error Resume Next
    closeDebugFile
    
End Sub

Public Function getStafVersion() As Integer
    getStafVersion = getStafMajorVersion()
End Function
