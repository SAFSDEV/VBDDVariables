VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalMappings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Globally shared Application Map storage"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'##############################################################################
'# <a name="globalmappings"></a>
'#
'#    GlobalMappings.CLS  SOURCECODE
'#
'# MODULE DESCRIPTION:
'#
'#      Cache for Application Map data.
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
'#          Set oGlobalMap = DDVariableStore.GlobalMappings
'#
'#              OR
'#
'#          Set oGlobalMap = CreateObject("DDVariableStore.GlobalMappings")
'#
'#
'# This library is defined by:
'#
'# GlobalMappings.CLS       ActiveX Class
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
'#
'# Copyright (2001) SAS Institute
'# GNU General Public License: http://www.opensource.org/licenses/gpl-license.php
'##############################################################################

'# FORWARD DECLARATIONS

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


'############################################################################
'# <a name="defaultappmapsection"></a>
'#
'#  Property DefaultAppMapSection As String
'#
'# DESCRIPTION:
'#
'#      The useable fallback lookup section for searching an Application
'#      Map.  If the fallback lookup section value has not been set, we will
'#      return the value of the DEFAULT FALLBACK LOOKUP SECTION.
'#      In this way, the user does not have to perform the test and setting of
'#      the fallback lookup section to use.
'#
'#      The FALLBACK LOOKUP SECTION is that section of the Application Map
'#      that will be searched for a matching componentID in the event that
'#      the specified windowsID section does not exist or does not contain that
'#      componentID.
'#
'#      By default, this FALLBACK LOOKUP SECTION will be the [ApplicationConstants]
'#      section of the application map.  The user can change this setting by
'#      providing a new value.
'#
'#
'# USAGE:
'#
'#      object.DefaultAppMapSection = strValue     'sets the value of the property
'#
'#      strValue = object.DefaultAppMapSection     'gets a standard data type from the property
'#
'#      (where 'object' is a GlobalMappings object)
'#
'#
'# DATA TYPE:
'#
'#      String
'#
'#
'# ERRORS:
'#
'#      (none)
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 07, 2001
'# History:
'#
'#      AUG 07, 2001    Original Release
'#      JUN 16, 2003    (Carl Nagle) Made SAFSVARS usage available
'#
'############################################################################
Public Property Get DefaultAppMapSection() As String
    If Not bUseMAPS Then GoTo eh
    Dim status As Long
    message = "DEFAULTMAPSECTION"
    status = submit(SAFS_SAFSMAPS_SERVICE, message)
    If status = STAF_Ok Then
        gDDDefaultMapSection = resultStr
        DefaultAppMapSection = resultStr
        Exit Property
    End If
eh:
    If gDDDefaultMapSection = "" Then gDDDefaultMapSection = gDD_DEFAULT_MAP_SECTION
    DefaultAppMapSection = gDDDefaultMapSection
    Exit Property
End Property


' ###########################################################################
Public Property Let DefaultAppMapSection(section As String)
    gDDDefaultMapSection = section
    If bUseMAPS Then
        message = "DEFAULTMAPSECTION SECTION " & lenMark(section)
        submit SAFS_SAFSMAPS_SERVICE, message
    End If
End Property


'############################################################################
'# <a name="currentappmappath"></a>
'#
'#  Property CurrentAppMapPath As String
'#
'# DESCRIPTION:
'#
'#      Contains the active AppMap (set by the user).
'#      You MUST provide a full explicit path to the map.  The property will
'#      be an empty string until an AppMap has been set.
'#
'#
'# USAGE:
'#
'#      object.CurrentAppMapPath = strValue     'sets the value of the property
'#
'#      strValue = object.CurrentAppMapPath     'gets a standard data type from the property
'#
'#      (where 'object' is a GlobalMappings object)
'#
'#
'# DATA TYPE:
'#
'#      String
'#
'# ERRORS:
'#
'#       none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 07, 2001
'# History:
'#
'#      AUG 07, 2001    Original Release
'#      JUN 16, 2003    (Carl Nagle) Made SAFSVARS usage available
'#
'############################################################################
Public Property Let CurrentAppMapPath(appMapPath As String)
    Dim path As String
    path = Dir$(appMapPath)
    If Len(path) = 0 Then
        Exit Property
    Else
        If Not bUseMAPS Then
            gDDCurrentAppMap = appMapPath
        Else
            message = "DEFAULTMAP " & lenMark(appMapPath)
            submit SAFS_SAFSMAPS_SERVICE, message
        End If
    End If
End Property


' ###########################################################################
Public Property Get CurrentAppMapPath() As String
    If Not bUseMAPS Then
        CurrentAppMapPath = gDDCurrentAppMap
        Exit Property
    End If
    message = "DEFAULTMAP"
    If submit(SAFS_SAFSMAPS_SERVICE, message) = STAF_Ok Then CurrentAppMapPath = resultStr
End Property


'############################################################################
'# <a name="getguiid"></a>
'#
'#  Function GetGUIID (windowID As String,
'#                       componentID As String,
'#                       appMapPath As String) As String
'#
'# DESCRIPTION:
'#
'#      Retrieves the recognition method for a given component in a given
'#      window.  It will find this value in the collection or in the file at
'#      the provided appMapPath.  This file is of the Windows INI format.
'#
'#      The routine stores all retrieved values into a collection so that
'#      it only has to resort to the file the first time a particular component is
'#      referenced.
'#
'# PARAMETERS:
'#
'#      windowID        String identifying the name of the window in which we
'#                      will find the componentID.
'#                      This is identical to a "Section" of an INI file which
'#                      our map files mimic.
'#
'#                      If no value is provided("") then we will attempt to check the
'#                      section identified by the gDDDefaultMap.
'#                      Similarly, if a search using a WindowID finds no match, we will
'#                      check the gDDDefaultMap location.
'#
'#
'#      componentID     String identifying the name of the component for which
'#                      we are seeking the recognition method (GUIID).
'#                      This is identical to a "Key" of an INI file which our map
'#                      files mimic.
'#
'#
'#      appMapPath      String specifying the path and name to the appmap to use.
'#                      If no path is provided then we will attempt to use the
'#                      CurrentAppMapPath.
'#
'# RETURNS:
'#
'#      If successful, the GUIID string (recognition method) of the specified component.
'#
'#      An empty string on failure ("").  This can be due to invalid values of any
'#      provided parameters or other errors.
'#
'# ERRORS:
'#
'#      none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 07, 2001
'# History:
'#
'#      AUG 07, 2001    Original Release
'#      JUN 16, 2003    (Carl Nagle) Made SAFSVARS usage available
'#
'############################################################################
Public Function GetGUIID(windowID As String, componentID As String, appMapPath As String) As String
    If Not bUseMAPS Then
        GetGUIID = old_GetGUIID(windowID, componentID, appMapPath)
        Exit Function
    End If
    
    Dim status As Long
    Dim mapname As String
    Dim windowname As String
    
    GetGUIID = ""
    If componentID = "" Then Exit Function
    
    mapname = appMapPath
    If Len(mapname) = 0 Then mapname = CurrentAppMapPath
    If Len(mapname) = 0 Then
        message = "GETITEM SECTION "
    Else
        message = "GETITEM " & lenMark(mapname) & " SECTION "
    End If
    
    windowname = windowID
    If Len(windowname) = 0 Then windowname = SAFS_SAFSMAPS_DEFAULTMAPSECTION
    
    message = message & lenMark(windowname) & " ITEM " & lenMark(componentID)
    
    If submit(SAFS_SAFSMAPS_SERVICE, message) = STAF_Ok Then
        If resultLen > 0 Then GetGUIID = resultStr
    End If
            
End Function
Private Function old_GetGUIID(windowID As String, componentID As String, appMapPath As String) As String
    
    Dim status As Integer
    Dim mapname As String
    Dim mapPath As String
    Dim windowname As String
    Dim componentName As String
    Dim value As String
    Dim collWin As DDVariables       'collection of Windows inside MapStore collection of Maps
    Dim collCmp As DDVariables       'collection of components inside window
        
    old_GetGUIID = ""             'set failure return value
    
    'validate all required parameters
    If componentID = "" Then Exit Function
    mapPath = appMapPath
    If Len(mapPath) = 0 Then mapPath = CurrentAppMapPath
    If Len(mapPath) = 0 Then Exit Function
    
    windowname = windowID
    
    If windowname = "" Then windowname = DefaultAppMapSection  'or check ApplicationConstants in AppMap?
    
    '***** removed UCASE to see if causing internationalization problems
    'windowName = UCase$(windowName)
    'componentName = UCase$(componentID)
    componentName = componentID
    
    'insure the collection of Maps, windows, and components exist
    status = gDDMappings.GetVariableValue(mapPath, collWin)
    
    'if an error, check provided path
    If status Then
        mapname = Dir$(mapPath)
        If mapname = "" Then
            Exit Function
        End If
    End If
    
    'if the path is OK then we have a new map in use
    If collWin Is Nothing Then
        Set collWin = New DDVariables
        Set collCmp = New DDVariables
    Else
        status = collWin.GetVariableValue(windowname, collCmp)
        If collCmp Is Nothing Then
            Set collCmp = New DDVariables
        End If
    End If
    
    value = ""
    status = collCmp.GetVariableValue(componentName, value)
        
    'if status <> 0 use the WIN32 API to try to locate the value and add it to array
    If status Then

        Const buffersize& = 1024            'Long value
        Dim bufferused As Long              'API sets this to #of bytes in buffer
        Dim buffer As String * buffersize   'API fills this buffer with value
        
        bufferused = GetPrivateProfileString(windowname, componentName, "", _
                                              buffer, buffersize, mapPath)
        
        'if not found then try the default lookup location
        If (bufferused < 1) And (StrComp(windowname, DefaultAppMapSection, 1) <> 0) Then
                bufferused = GetPrivateProfileString(DefaultAppMapSection, componentName, "", _
                                                      buffer, buffersize, mapPath)
        End If
        
        If bufferused > 0 Then
            value = Left$(buffer, bufferused)   'may have to strip trailing chr$(0)s
            status = collCmp.SetVariableValue(componentName, value)
            status = collWin.SetVariableValue(windowname, collCmp)
            status = gDDMappings.SetVariableValue(mapPath, collWin)
        End If
        
    End If
    
    old_GetGUIID = value
    Set collWin = Nothing
    Set collCmp = Nothing

End Function

'############################################################################
'# <a name="resetandclear"></a>
'#
'#  Function ResetAndClear() As Long
'#
'# DESCRIPTION:
'#
'#  Removes all App Map cached mappings.
'#
'#
'# PARAMETERS:
'#
'#       (none)
'#
'#
'# RETURNS:
'#
'#      (Should always return 0.)
'#
'#
'# ERRORS:
'#
'#       (none)
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: DEC 11, 2003
'# History:
'#
'#      DEC 11, 2003    Original Release
'#
'############################################################################
Public Function ResetAndClear() As Long
    ResetAndClear = gDDMappings.ResetAndClear
    ResetSTAFChecks
    If bUseMAPS Then
        Dim status As Long
        message = "CLEARCACHE"
        status = submit(SAFS_SAFSMAPS_SERVICE, message)
    End If
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
    If (isSAFSMAPSAvailable) And (handle > 1) Then
        bUseMAPS = True
    Else
        bUseMAPS = False
    End If
End Function

Private Sub Class_Initialize()
    If gDDMappings Is Nothing Then
        Set gDDMappings = New DDVariables
        glngDDMRefCount = 0    'just to make sure
    End If
    
    ResetSTAFChecks
    
    glngDDMRefCount = glngDDMRefCount + 1
End Sub

Private Sub Class_Terminate()
    glngDDMRefCount = glngDDMRefCount - 1
    If glngDDMRefCount = 0 Then
        Set gDDMappings = Nothing
    End If
    If (glngDDVRefCount + glngDDMRefCount = 0) And (handle > 1) Then
        ProcUnregister handle
    End If
End Sub

