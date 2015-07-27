Attribute VB_Name = "STAFConstants"
Public Const SAFS_HOOK_MUTEX = "SAFS/Hook/TRD"
Public Const SAFS_ROBOTJ_MUTEX = SAFS_HOOK_MUTEX  'deprecated

Public Const SAFS_ENGINE_EVENT_START = "Start"
Public Const SAFS_ENGINE_EVENT_READY = "Ready"
Public Const SAFS_ENGINE_EVENT_DISPATCH = "Dispatch"
Public Const SAFS_ENGINE_EVENT_RUNNING = "Running"
Public Const SAFS_ENGINE_EVENT_RESULTS = "Results"
Public Const SAFS_ENGINE_EVENT_DONE = "Done"
Public Const SAFS_ENGINE_EVENT_SHUTDOWN = "Shutdown"

Public Const SAFS_ROBOTJ_EVENT_START = "SAFS/RobotJStart"       'deprecated
Public Const SAFS_ROBOTJ_EVENT_READY = "SAFS/RobotJReady"       'deprecated
Public Const SAFS_ROBOTJ_EVENT_DISPATCH = "SAFS/RobotJDispatch" 'deprecated
Public Const SAFS_ROBOTJ_EVENT_RUNNING = "SAFS/RobotJRunning"   'deprecated
Public Const SAFS_ROBOTJ_EVENT_RESULTS = "SAFS/RobotJResults"   'deprecated
Public Const SAFS_ROBOTJ_EVENT_DONE = "SAFS/RobotJDone"         'deprecated
Public Const SAFS_ROBOTJ_EVENT_SHUTDOWN = "SAFS/RobotJShutdown" 'deprecated

Public Const SAFS_CYCLE_TRD_PREFIX = "SAFS/Cycle/"
Public Const SAFS_SUITE_TRD_PREFIX = "SAFS/Suite/"
Public Const SAFS_STEP_TRD_PREFIX = "SAFS/Step/"
Public Const SAFS_SHARED_TRD_PREFIX = "SAFS/Hook/"

Public Const SAFS_TRD_FILENAME = "filename"
Public Const SAFS_TRD_LINENUMBER = "linenumber"
Public Const SAFS_TRD_INPUTRECORD = "inputrecord"
Public Const SAFS_TRD_SEPARATOR = "separator"
Public Const SAFS_TRD_TESTLEVEL = "testlevel"
Public Const SAFS_TRD_APPMAPNAME = "appmapname"
Public Const SAFS_TRD_FAC = "fac"
Public Const SAFS_TRD_STATUSCODE = "statuscode"
Public Const SAFS_TRD_STATUSINFO = "statusinfo"

'Abbot
Public Const SAFS_ABBOT_PROCESS = "SAFS/Abbot"
Public Const SAFS_ABBOT_PROCESS_ID = "SAFS/AbbotID"

'Classic
Public Const SAFS_ROBOTC_PROCESS = "SAFS/RobotClassic"
Public Const SAFS_ROBOTC_PROCESS_ID = "SAFS/RobotClassicID"

'RobotJ
Public Const SAFS_ROBOTJ_PROCESS = "SAFS/RobotJ"
Public Const SAFS_ROBOTJ_PROCESS_ID = "SAFS/RobotJID"

'QTP
Public Const SAFS_QTP_PROCESS = "SAFS/QTP"
Public Const SAFS_QTP_PROCESS_ID = "SAFS/QTPID"

'WinRunner
Public Const SAFS_WINRUNNER_PROCESS = "SAFS/WinRunner"
Public Const SAFS_WINRUNNER_PROCESS_ID = "SAFS/WinRunnerID"

'DriverCommands
Public Const SAFS_DRIVERCOMMANDS_PROCESS = "SAFS/DriverCommands"
Public Const SAFS_DRIVERCOMMANDS_PROCESS_ID = "SAFS/DriverCommandsID"


'DDVariableStore.DLL
Public Const SAFS_DDVDLL_PROCESS = "SAFS/DDVDLL"
Public Const SAFS_DDVDLL_PROCESS_ID = "SAFS/DDVDLLID"

Public Const SAFS_STAF_ERROR = "_STAF_ERROR_"

Public Const SAFS_SAFSLOGS_PROCESS = "SAFSLoggingService"
Public Const SAFS_SAFSLOGS_SERVICE = "SAFSLOGS"
Public Const SAFS_SAFSVARS_PROCESS = "SAFSVariableService"
Public Const SAFS_SAFSVARS_SERVICE = "SAFSVARS"
Public Const SAFS_SAFSMAPS_PROCESS = "SAFSAppMapService"
Public Const SAFS_SAFSMAPS_SERVICE = "SAFSMAPS"
Public Const SAFS_SAFSINPUT_PROCESS = "SAFSInputService"
Public Const SAFS_SAFSINPUT_SERVICE = "SAFSINPUT"
Public Const SAFS_SAFSMAPS_DEFAULTMAPSECTION = "DEFAULTMAPSECTION"

Public Const SAFS_HOOK_SHUTDOWN_COMMAND = "SHUTDOWN_HOOK"

'STAF result/error codes
Public Const STAF_NOT_INSTALLED = -1
Public Const STAF_Ok = 0
Public Const STAF_InvalidAPI = 1
Public Const STAF_UnknownService = 2
Public Const STAF_InvalidHandle = 3
Public Const STAF_HandleAlreadyExists = 4
Public Const STAF_HandleDoesNotExist = 5
Public Const STAF_UnknownError = 6
Public Const STAF_InvalidRequestString = 7
Public Const STAF_InvalidServiceResult = 8
Public Const STAF_REXXError = 9
Public Const STAF_BaseOSError = 10
Public Const STAF_ProcessAlreadyComplete = 11
Public Const STAF_ProcessNotComplete = 12
Public Const STAF_VariableDoesNotExist = 13
Public Const STAF_UnResolvableString = 14
Public Const STAF_InvalidResolveString = 15
Public Const STAF_NoPathToMachine = 16
Public Const STAF_FileOpenError = 17
Public Const STAF_FileReadError = 18
Public Const STAF_FileWriteError = 19
Public Const STAF_FileDeleteError = 20
Public Const STAF_STAFNotRunning = 21
Public Const STAF_CommunicationError = 22
Public Const STAF_TrusteeDoesNotExist = 23
Public Const STAF_InvalidTrustLevel = 24
Public Const STAF_AccessDenied = 25
Public Const STAF_STAFRegistrationError = 26
Public Const STAF_ServiceConfigurationError = 27
Public Const STAF_QueueFull = 28
Public Const STAF_NoQueueElement = 29
Public Const STAF_NotifieeDoesNotExist = 30
Public Const STAF_InvalidAPILevel = 31
Public Const STAF_ServiceNotUnregisterable = 32
Public Const STAF_ServiceNotAvailable = 33
Public Const STAF_SemaphoreDoesNotExist = 34
Public Const STAF_NotSemaphoreOwner = 35
Public Const STAF_SemaphoreHasPendingRequests = 36
Public Const STAF_Timeout = 37
Public Const STAF_JavaError = 38
Public Const STAF_ConverterError = 39
Public Const STAF_ServiceAlreadyExists = 40
Public Const STAF_InvalidObject = 41
Public Const STAF_InvalidParm = 42
Public Const STAF_RequestNumberNotFound = 43
Public Const STAF_InvalidAsynchOption = 44
Public Const STAF_RequestNotComplete = 45
Public Const STAF_ProcessAuthenticationDenied = 46
Public Const STAF_InvalidValue = 47
Public Const STAF_DoesNotExist = 48
Public Const STAF_AlreadyExists = 49
Public Const STAF_DirectoryNotEmpty = 50
Public Const STAF_DirectoryCopyError = 51

Public Const STAF_UserDefined = 4000

' STAF request submission syncOption
Public Const STAF_ReqSync = 0
Public Const STAF_ReqFireAndForget = 1
Public Const STAF_ReqQueue = 2
Public Const STAF_ReqRetain = 3
Public Const STAF_ReqQueueRetain = 4

