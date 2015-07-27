Attribute VB_Name = "STAFModule"
Option Explicit

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Any, _
                                                             ByVal lpSrc As Any, _
                                                             ByVal length As Long)
                                                             
'requires our STAFWrap DLL for VB
Declare Function STAFregister Lib "STAFWrap" Alias "stdSTAFRegister" (ByVal processName As String, handle As Long) As Long
Declare Function STAFunregister Lib "STAFWrap" Alias "stdSTAFUnRegister" (ByVal handle As Long) As Long
Declare Function STAFSubmit Lib "STAFWrap" Alias "stdSTAFSubmit" (ByVal handle As Long, _
                            ByVal where As String, ByVal service As String, _
                            ByVal request As String, ByVal requestLength As Long, _
                            resultPtr As Any, resultLength As Long) As Long
'where, service and request will be declared as Long, they indicate the address of an UTF8 byte array (terminate by null)
Declare Function STAFSubmitUTF8 Lib "STAFWrap" Alias "stdSTAFSubmitUTF8" (ByVal handle As Long, _
                            ByVal where As Long, ByVal service As Long, _
                            ByVal request As Long, ByVal requestLength As Long, _
                            resultPtr As Any, resultLength As Long) As Long
                            
Declare Function STAFFree Lib "STAFWrap" Alias "stdSTAFFree" (ByVal handle As Long, ByVal resultPtr As Any) As Long

'**** Cannot seem to call STAF.DLL directly ****
'Declare Function STAFregister Lib "STAF" Alias "STAFRegister" (ByVal processName As String, handle As Long) As Long
'Declare Function STAFunregister Lib "STAF" Alias "STAFUnRegister" (ByVal handle As Long) As Long
'Declare Function STAFsubmit Lib "STAF" Alias "STAFSubmit" (ByVal handle As Long, _
'                            ByVal where As String, ByVal service As String, _
'                            ByVal request As String, ByVal requestLength As Long, _
'                            resultPtr As Any, resultLength As Long) As Long
'Declare Function STAFfree Lib "STAF" Alias "STAFFree" (ByVal handle As Long, ByVal resultPtr As Any) As Long

'Call native function lstrlenA of kernel32.dll, instead of VB function Len() to get the length of message passed
'to function STAFSubmit() to solve the problem of Chinese string
'BE CAREFUL!!! to call this, if parameter string contains non-ansi-character, the parameter itself will be modified
'the non-ansi-character will replaced by ?
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Public handle As Long       'our registered STAF handle
Public message As String    'messages we send as requests to STAF
Public messageLen As Long   'we calculate this in "submit"
Public result As Long       'pointer returned from STAF to STAF response string
Public resultStr As String  'our CopyMem'd actual STAF response string
Public resultLen As Long    'the expected length of our copied resultStr

Public safsvars As Boolean
Public safsmaps As Boolean
Public mapschecked As Boolean
Public varschecked As Boolean

Dim errnum As Integer
Dim errdsc As String
Dim errsrc As String
Dim stafMajorVersion As Integer

'##############################################################################
'#
'#  Function RegisterProcess   (aprocess As String, ahandle As Long) As Long
'#
'#  DESCRIPTION:
'#
'#      Register a named process with STAF if it is not already registered.
'#      The routine will return the STAF handle in ahandle.
'#
'#      Also, if the DLL's internal 'handle' has not already been
'#      set by any previous use then the routine will set that
'#      handle to match the output ahandle.
'#
'#  PARAMETERS:
'#
'#      aprocess    The (unique?) string name to identify the process in STAF.
'#
'#      ahandle     The variable to receive the handleID once registered.
'#                  Also, if the DLL's internal 'handle' has not already been
'#                  set by any previous use then the routine will set that
'#                  handle to match the output ahandle.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'# ERRORS:
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
Public Function registerProcess(ByVal aprocess As String, ahandle As Long) As Long
    
    'error if STAF not installed or running
    On Error GoTo eh
    
    If ahandle < 2 Then
        registerProcess = CLng(STAFregister(ByVal aprocess, ahandle))
        If handle < 2 Then handle = ahandle
    Else 'already registered?
        registerProcess = CLng(STAF_Ok)
    End If
    Exit Function
eh:
    registerProcess = CLng(STAF_NOT_INSTALLED)
    Exit Function
End Function


'##############################################################################
'#
'#  Function ProcRegister () As Long
'#
'#  DESCRIPTION:
'#
'#      Calls RegisterProcess using the SAFS_DDVDDL_PROCESS constant for the
'#      process name and the DLL's global handle as the variable to receive
'#      the handleID.
'#
'#  PARAMETERS:
'#
'#      (none)
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'# ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: DEC 16, 2003
'# History:
'#
'#      DEC 16, 2003    Original Release
'#      NOV 15, 2005    (Carl Nagle) Modified to use RegisterProcess
'#
'##############################################################################
Public Function ProcRegister() As Long
    ProcRegister = registerProcess(SAFS_DDVDLL_PROCESS, handle)
End Function


'##############################################################################
'#
'#  Function ProcUnregister (ahandle As Long) As Long
'#
'#  DESCRIPTION:
'#
'#      UnRegister our DLL with STAF and set its value to 0.
'#
'#      Also, if the DLL's internal 'handle' has the same value as
'#      the one provided then it will be zeroed as well.
'#
'#  PARAMETERS:
'#
'#      ahandle     The handleID to unregister from STAF.
'#                  Also, if the DLL's internal 'handle' has the same value as
'#                  the one provided then it will be zeroed as well.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'# ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: DEC 16, 2003
'# History:
'#
'#      DEC 16, 2003    Original Release
'#
'##############################################################################
Public Function ProcUnregister(ahandle As Long) As Long
    On Error GoTo eh
    resetServiceChecks
    If ahandle > 1 Then
        ProcUnregister = CLng(STAFunregister(ahandle))
    Else
        ProcUnregister = CLng(STAF_InvalidHandle)
    End If
    If ahandle = handle Then handle = 0
    Exit Function
eh:
    ProcUnregister = CLng(STAF_NOT_INSTALLED)
    ahandle = 0
    Exit Function
End Function

'##############################################################################
'#
'#  Function submitWhereRequest(ahandle As Long,
'#                              where As String,
'#                              service As String,
'#                              amessage As String) As Long
'#
'#  DESCRIPTION:
'#
'#  Submit a request to STAF.  This may be on the "local" or a remote machine.
'#
'#  PARAMETERS:
'#
'#      ahandle     The registered STAF handle id used for this request.
'#
'#      where       The machine location to receive the STAF request.
'#                  Use "local" for the local machine.
'#
'#      service     The name of the service to receive the STAF request.
'#
'#      amessage    The message to send to the service.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'# ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: NOV 15, 2005
'# History:
'#
'#      NOV 15, 2005    Original Release
'#      OCT 06, 2006    (Carl Nagle) Mod for Unicode and NLS testing
'#      JUL 30, 2012    (Lei Wang) Use function lstrlen() instead of Len() to calculate the length of String.
'#                               Otherwise STAFSubmit() can't handle NLS message correctly.
'#      OCT 25, 2013    (Lei Wang) Use STAFSubmitUTF8() instead of STAFSubmit() to handle all kinds of characters.
'#
'##############################################################################
Public Function submitWhereRequest(ahandle As Long, where As String, service As String, amessage As String) As Long
    result = 0
    resultLen = 0
    resultStr = ""
    
    Dim utf8Result() As Byte
    Dim utf8Where() As Byte
    Dim utf8Service() As Byte
    Dim utf8Message() As Byte
    
    Dim newline(0) As Byte
    Dim charp(0) As Byte
    
    On Error GoTo eh
    newline(0) = &HA
    charp(0) = &H23 '#

    'convert unicode-string (where, service, amessage) to utf8-encoded-byte-array
    'STAFSubmitUTF8 requires that where, service and amessage are utf8 bytes
    utf8Where = UnicodeStringToUTF8Bytes(where)
    utf8Service = UnicodeStringToUTF8Bytes(service)
    utf8Message = UnicodeStringToUTF8Bytes(amessage)
    'Get the length of array utf8Message
    messageLen = ArraySize(utf8Message)
    'STAFSubmitUTF8 requires that where, service and amessage are terminated by null
    AddNullTerminatorToArray utf8Where
    AddNullTerminatorToArray utf8Service
    AddNullTerminatorToArray utf8Message
    
    If DEBUG_ENABLED Then
        DebugString "" 'Write a blank line
        DebugString "origianl submit: where=" & where & "  $$  service=" & service & "  $$  message=" & amessage & " $$  messageLen=" & CStr(messageLen)
        
        'Log the "staf utf8 request" to a binary debug file
        DebugBinary newline
        DebugBinary utf8Where
        DebugBinary charp
        DebugBinary utf8Service
        DebugBinary charp
        DebugBinary utf8Message
        DebugBinary charp
        DebugBinary newline
    End If
    
    submitWhereRequest = CLng(STAFSubmitUTF8(ahandle, ByVal VarPtr(utf8Where(0)), ByVal VarPtr(utf8Service(0)), ByVal VarPtr(utf8Message(0)), messageLen, result, resultLen))
    
    On Error Resume Next
    If (result <> 0) And (resultLen > 0) Then
        'Be careful to resize the array!!! Dim arr(1) means holding 2 elements
        ReDim utf8Result(resultLen - 1)
    
        'result is the address where "utf8 bytes result" is stored, length is resultLen
        CopyMemory ByVal VarPtr(utf8Result(0)), result, resultLen
        resultStr = UTF8BytesToUnicodeString(utf8Result)
        
        If DEBUG_ENABLED Then
            DebugString "before conversion: utf8Result=:" & resultLen & ":" & CStr(utf8Result)
            
            'Log the result (utf8 bytes) to a binary debug file
            DebugBinary newline
            DebugBinary utf8Result
            DebugBinary newline
    
            DebugString "after conversion: resultStr=" & resultStr
        End If
        
    End If
    If result <> 0 Then STAFFree ahandle, result
    
    Exit Function
eh:
    submitWhereRequest = CLng(STAF_NOT_INSTALLED)
    Exit Function
    
End Function

'##############################################################################
'#
'#  Function submitSTAFResultRequest(ahandle As Long,
'#                                   where As String,
'#                                   service As String,
'#                                   amessage As String,
'#                                   astafresult as STAFResult) As Long
'#
'#  DESCRIPTION:
'#
'#      Submit a request to STAF.  This may be on the "local" or a remote machine.
'#      Fill the provided STAFResult object with the STAFResult.rc and
'#      STAFResult.result of the call.  The routine simply calls submitWhereRequest
'#      and then fills in the STAFResult after the call.
'#
'#      This routine takes a STAFResult object, so VBScript cannot use this
'#      routine directly.  VBScript will need to use submitSTAFResultVariant
'#      which passes the STAFResult object as a Variant.
'#
'#  PARAMETERS:
'#
'#      ahandle     The registered STAF handle id used for this request.
'#
'#      where       The machine location to receive the STAF request.
'#                  "local" for the local machine.
'#
'#      service     The name of the service to receive the STAF request.
'#
'#      amessage    The message to send to the service.
'#
'#      astafresult STAFResult to receive STAF return code and response string.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'# ERRORS:
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
Public Function submitSTAFResultRequest(ahandle As Long, where As String, service As String, amessage As String, astafresult As STAFResult) As Long
    
    If astafresult Is Nothing Then Set astafresult = New STAFResult
    astafresult.rc = submitWhereRequest(ahandle, where, service, amessage)
    astafresult.result = resultStr
    submitSTAFResultRequest = astafresult.rc

End Function


'##############################################################################
'#
'#  Function submitSTAFResultVariant(ahandle As Long,
'#                                   where As String,
'#                                   service As String,
'#                                   amessage As String,
'#                                   astafresult as Variant) As Long
'#
'#  DESCRIPTION:
'#
'#      Submit a request to STAF.  This may be on the "local" or a remote machine.
'#      Fill the provided Variant STAFResult object with the STAFResult.rc and
'#      STAFResult.result of the call.  The routine simply calls submitWhereRequest
'#      and then fills in the Variant STAFResult after the call.
'#
'#      This routine takes a Variant STAFResult Object as would be provided by
'#      VBScript since VBScript cannot pass Objects as anything but Variants.
'#
'#  PARAMETERS:
'#
'#      ahandle     The registered STAF handle id used for this request.
'#
'#      where       The machine location to receive the STAF request.
'#                  "local" for the local machine.
'#
'#      service     The name of the service to receive the STAF request.
'#
'#      amessage    The message to send to the service.
'#
'#      astafresult Variant STAFResult Object to receive the results of the
'#                  STAF call.  The Variant is expected to be a STAFResult
'#                  object.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'# ERRORS:
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
Public Function submitSTAFResultVariant(ahandle As Long, where As String, service As String, amessage As String, astafresult As Variant) As Long
    
    If astafresult Is Nothing Then Set astafresult = New STAFResult
    astafresult.rc = submitWhereRequest(ahandle, where, service, amessage)
    astafresult.result = resultStr
    submitSTAFResultVariant = astafresult.rc

End Function


'##############################################################################
'#
'#  Function submitHandle(ahandle As Long,
'#                        service As String,
'#                        amessage As String) As Long
'#
'#  DESCRIPTION:
'#
'#      Submit a "local" request to STAF using the provided handle.
'#      The routine simply calls submitWhereRequest using the default "local"
'#      where value.
'#
'#      Upon a successful call, our global results variables will contain:
'#
'#          resultStr    will contain any text response from the STAF service.
'#          result       will contain a pointer value used to make resultStr.
'#          resultLen    will contain the expected length of resultStr.
'#
'#      Note that some successful calls do not generate any resultStr response.
'#      For example, a request for the value of a variable whose value is an empty,
'#      zero-length string will have an empty zero-length response.
'#
'#  PARAMETERS:
'#
'#      ahandle     The registered STAF handle id used for this request.
'#
'#      service     The name of the service to receive the STAF request.
'#
'#      amessage    The message to send to the service.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#      Upon a successful call, our global results variables will contain:
'#
'#          resultStr    will contain any text response from the STAF service.
'#          result       will contain a pointer value used to make resultStr.
'#          resultLen    will contain the expected length of resultStr.
'#
'# ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: DEC 16, 2003
'# History:
'#
'#      DEC 16, 2003    Original Release
'#      NOV 15, 2005    (Carl Nagle) Modified to use submitWhereRequest
'#
'##############################################################################
Public Function submitHandle(ahandle As Long, service As String, amessage As String) As Long
    submitHandle = submitWhereRequest(ahandle, "local", service, amessage)
End Function


'##############################################################################
'#
'#  Function submit(service As String,
'#                  amessage As String) As Long
'#
'#  DESCRIPTION:
'#
'#      Submit a "local" request to STAF using the DLL's global handle.
'#      The routine simply calls submitHandle using the DLL's global handle
'#      and the provided parameters.
'#
'#      Upon a successful call, our global results variables will contain:
'#
'#          resultStr    will contain any text response from the STAF service.
'#          result       will contain a pointer value used to make resultStr.
'#          resultLen    will contain the expected length of resultStr.
'#
'#      Note that some successful calls do not generate any resultStr response.
'#      For example, a request for the value of a variable whose value is an empty,
'#      zero-length string will have an empty zero-length response.
'#
'#  PARAMETERS:
'#
'#      service     The name of the service to receive the STAF request.
'#
'#      amessage    The message to send to the service.
'#
'#  RETURNS:
'#
'#      STAF_OK or STAF result codes if a STAF error occurs.
'#      STAF_NOT_INSTALLED if the STAF or STAFWrap DLLs are not found.
'#
'#      Upon a successful call, our global results variables will contain:
'#
'#          resultStr    will contain any text response from the STAF service.
'#          result       will contain a pointer value used to make resultStr.
'#          resultLen    will contain the expected length of resultStr.
'#
'# ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: DEC 16, 2003
'# History:
'#
'#      DEC 16, 2003    Original Release
'#      NOV 15, 2005    (Carl Nagle) Modified to use submitHandle
'#
'##############################################################################
Public Function submit(service As String, amessage As String) As Long
    submit = submitHandle(handle, service, amessage)
End Function

'##############################################################################
'#
'#  Function isToolAvailable (toolname As String) As Boolean
'#
'#  DESCRIPTION:
'#
'#      Check to see if a specific tool or service is already registered and
'#      running in STAF.
'#
'#      The routine queries the STAF HANDLE service to see if the named tool
'#      is running.
'#
'#      This routine uses the DLL's global handle to communicate with STAF.
'#      If the handle has not been registered, then the routine will attempt
'#      to register with STAF first.
'#
'#  PARAMETERS:
'#
'#      toolname    A unique toolname as seen by the STAF HANDLE service.
'#
'#  RETURNS:
'#
'#      True if STAF is installed, running, and the requested tool has a
'#      registered handle.  False otherwise.
'#
'# ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: DEC 16, 2003
'# History:
'#
'#      DEC 16, 2003    Original Release
'#      AUG 26, 2009    LeiWang     Modify to adapt STAF 3
'#
'##############################################################################
Public Function isToolAvailable(toolname As String) As Boolean
    Dim status As Long
    Dim service As String
    Dim request As String
    Dim result As STAFResult

    isToolAvailable = False
    If handle < 2 Then
        ProcRegister
        If handle < 2 Then Exit Function
    End If
    
    service = "handle"
    If getStafMajorVersion() < 3 Then
        request = " query name " & lenMark(toolname)
    Else
        request = " list name " & lenMark(toolname)
    End If
    
    On Error GoTo eh
    
    status = submitSTAFResultRequest(handle, "local", service, request, result)

    If (status = STAF_Ok) And (InStr(result.result, toolname) > 0) Then
        isToolAvailable = True
    End If
eh:
    Exit Function
End Function


'##############################################################################
'#
'#  Function isSAFSVARSAvailable () As Boolean
'#
'#  DESCRIPTION:
'#
'#      Check to see if the SAFSVARS service is already registered and
'#      running in STAF.
'#
'#      Simply calls isToolAvailable with the SAFSVARS process name.
'#
'#  PARAMETERS:
'#
'#      (none)
'#
'#  RETURNS:
'#
'#      True if SAFSVARS is running.  False otherwise.
'#
'# ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: DEC 16, 2003
'# History:
'#
'#      DEC 16, 2003    Original Release
'#
'##############################################################################
Public Function isSAFSVARSAvailable() As Boolean
    If Not varschecked Then
        safsvars = isToolAvailable(SAFS_SAFSVARS_PROCESS)
        varschecked = True
    End If
    isSAFSVARSAvailable = safsvars
End Function


'##############################################################################
'#
'#  Function isSAFSMAPSAvailable () As Boolean
'#
'#  DESCRIPTION:
'#
'#      Check to see if the SAFSMAPS service is already registered and
'#      running in STAF.
'#
'#      Simply calls isToolAvailable with the SAFSMAPS process name.
'#
'#  PARAMETERS:
'#
'#      (none)
'#
'#  RETURNS:
'#
'#      True if SAFSMAPS is running.  False otherwise.
'#
'# ERRORS:
'#
'#      Throws "Error in loading DLL" if STAF is not installed.
'#
'# Orig Author: Carl Nagle
'# Orig   Date: DEC 16, 2003
'# History:
'#
'#      DEC 16, 2003    Original Release
'#
'##############################################################################
Public Function isSAFSMAPSAvailable() As Long
    If Not mapschecked Then
        safsmaps = isToolAvailable(SAFS_SAFSMAPS_PROCESS)
        mapschecked = True
    End If
    isSAFSMAPSAvailable = safsmaps
End Function


'##############################################################################
'#
'#  Sub resetServiceChecks ()
'#
'#  DESCRIPTION:
'#
'#      Reset internal flags that otherwise stop subsequent STAF checks for
'#      running services to occur.
'#
'#      Currently the following services are checked and flagged:
'#
'#          SAFSMAPS
'#          SAFSVARS
'#
'#  PARAMETERS:
'#
'#      (none)
'#
'#
'# ERRORS:
'#
'#      (none)
'#
'# Orig Author: Carl Nagle
'# Orig   Date: DEC 16, 2003
'# History:
'#
'#      DEC 16, 2003    Original Release
'#
'##############################################################################
Public Sub resetServiceChecks()
    varschecked = False
    mapschecked = False
    safsvars = False
    safsmaps = False
End Sub


'##############################################################################
'#
'#  Function lenMark (value As String) As String
'#
'#  DESCRIPTION:
'#
'#      Create a length delimited STAF string from the input string.
'#      The resulting format is STAF compliant, :<len>:value
'#      This prevents STAF from changing or escaping any character values.
'#
'#  PARAMETERS:
'#
'#      value    The string to be submitted to STAF.
'#
'#  RETURNS:
'#
'#      The length denoted string in the STAF :len:value  format.
'#      The value will be unchanged if it is empty (0-length).
'#
'# ERRORS:
'#
'#      (none)
'#
'# Orig Author: Carl Nagle
'# Orig   Date: DEC 16, 2003
'# History:
'#
'#      DEC 16, 2003    Original Release
'#      OCT 06, 2006    (Carl Nagle) Mod for Unicode and NLS testing
'#
'##############################################################################
Public Function lenMark(value As String) As String
    Dim truevalue As String
    truevalue = value
    If Len(truevalue) > 0 Then
        truevalue = ":" & Trim$(Str$(Len(truevalue))) & ":" & value
    End If
    lenMark = truevalue
End Function


'##############################################################################
'#
'#  Function getStafMajorVersion () As Integer
'#
'#  DESCRIPTION:
'#
'#      Get the major version of STAF.
'#      For example, if staf version is 2.6.11, then major version will be 2.
'#
'#  RETURNS:
'#
'#      The major version of STAF.
'#
'# ERRORS:
'#
'#      (none)
'#
'# Orig Author: Lei Wang
'# Orig   Date: AUT 26, 2009
'# History:
'#
'#      AUT 26, 2009    Original Release
'#
'##############################################################################
Public Function getStafMajorVersion() As Integer
    Dim command As String
    Dim staf2request As String
    Dim staf3request As String
    Dim status As Integer
    Dim astafresult As STAFResult
    Dim versionArray() As String
    
    If stafMajorVersion <> 0 Then
        getStafMajorVersion = stafMajorVersion
        Exit Function
    End If
    
    command = "VAR"
    staf2request = " GLOBAL GET STAF/Version "
    staf3request = " GET SYSTEM VAR STAF/Version "
    
    If handle < 2 Then
        ProcRegister
        If handle < 2 Then Exit Function
    End If
    
    status = submitSTAFResultRequest(handle, "local", command, staf2request, astafresult)
    If (status <> STAF_Ok) Then
        status = submitSTAFResultRequest(handle, "local", command, staf3request, astafresult)
    End If
    
    If (status <> STAF_Ok) Then
        getStafMajorVersion = 0
        Exit Function
    End If
    
    'We get the STAF VERSION string
    versionArray = Split(astafresult.result, ".")
    stafMajorVersion = CInt(versionArray(0))
    getStafMajorVersion = stafMajorVersion
    
End Function