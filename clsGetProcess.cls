VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGetProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'In a module

Private Declare Function GetProcessMemoryInfo Lib "PSAPI.DLL" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcesses Lib "PSAPI.DLL" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)



Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const MAX_PATH = 260
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const TH32CS_SNAPPROCESS = &H2&
Private Const hNull = 0
Private Const WIN95_System_Found = 1
Private Const WINNT_System_Found = 2
Private Const Default_Log_Size = 10000000
Private Const Default_Log_Days = 0
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF
Private Const STANDARD_RIGHTS_ALL = &H1F0000


Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type


Private Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type


Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260 ' MAX_PATH
    End Type


Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long '1 = Windows 95.
    '2 = Windows NT
    szCSDVersion As String * 128
End Type


Public Function GetProcesses(ByVal EXEName As String) As String

    Dim booResult As Boolean
    Dim lngLength As Long
    Dim lngProcessID As Long
    Dim strProcessName As String
    Dim lngSnapHwnd As Long
    Dim udtProcEntry As PROCESSENTRY32
    Dim lngCBSize As Long 'Specifies the size, In bytes, of the lpidProcess array
    Dim lngCBSizeReturned As Long 'Receives the number of bytes returned
    Dim lngNumElements As Long
    Dim lngProcessIDs() As Long
    Dim lngCBSize2 As Long
    Dim lngModules(1 To 200) As Long
    Dim lngReturn As Long
    Dim strModuleName As String
    Dim lngSize As Long
    Dim lngHwndProcess As Long
    Dim lngLoop As Long
    Dim b As Long
    Dim c As Long
    Dim e As Long
    Dim d As Long
    Dim pmc As PROCESS_MEMORY_COUNTERS
    Dim lret As Long
    Dim strProcName2 As String
    Dim strProcName As String

    On Error GoTo Error_handler

   booResult = False

    EXEName = UCase$(Trim$(EXEName))
    lngLength = Len(EXEName)
    GetProcesses = ""

    Select Case getVersion()
        Case WIN95_System_Found 'Windows 95/98

        Case WINNT_System_Found 'Windows NT

            lngCBSize = 8 ' Really needs To be 16, but Loop will increment prior to calling API
            lngCBSizeReturned = 96

            Do While lngCBSize <= lngCBSizeReturned
                DoEvents
                'Increment Size
                lngCBSize = lngCBSize * 2
                'Allocate Memory for Array
                ReDim lngProcessIDs(lngCBSize / 4) As Long
                'Get Process ID's
                lngReturn = EnumProcesses(lngProcessIDs(1), lngCBSize, lngCBSizeReturned)
            Loop

            lngNumElements = lngCBSizeReturned / 4

            For lngLoop = 1 To lngNumElements
            DoEvents

            lngHwndProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lngProcessIDs(lngLoop))

            If lngHwndProcess <> 0 Then
                'Get an array of the module handles for the specified process
                lngReturn = EnumProcessModules(lngHwndProcess, lngModules(1), 200, lngCBSize2)

                'If the Module Array is retrieved, Get the ModuleFileName
                If lngReturn <> 0 Then

                    'Buffer with spaces first to allocate memory for byte array
                    strModuleName = Space(MAX_PATH)

                    lngSize = 500

                    lngReturn = GetModuleFileNameExA(lngHwndProcess, lngModules(1), strModuleName, lngSize)

                    'Remove trailing spaces
                    strProcessName = Left(strModuleName, lngReturn)

                    'Check for Matching Upper case result
                    strProcessName = UCase$(Trim$(strProcessName))

                    strProcName2 = GetElement(Trim(Replace(strProcessName, Chr$(0), "")), "\", 0, 0, GetNumElements(Trim(Replace(strProcessName, Chr$(0), "")), "\") - 1)

                    If strProcName2 = EXEName Then

                        'Get the Site of the Memory Structure
                        pmc.cb = LenB(pmc)

                           lret = GetProcessMemoryInfo(lngHwndProcess, pmc, pmc.cb)

                           GetProcesses = CStr(pmc.WorkingSetSize / 1024)

                    End If
                End If
            End If
            lngReturn = CloseHandle(lngHwndProcess)
            DoEvents
        Next

    End Select

IsProcessRunning_Exit:

Exit Function
Error_handler:
    Err.Raise Err, Err.Source, "ProcessInfo", Error
    Resume Next
End Function


Private Function getVersion() As Long

    Dim osinfo As OSVERSIONINFO
    Dim retvalue As Integer

    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)
    getVersion = osinfo.dwPlatformId

End Function


Private Function StrZToStr(s As String) As String
    StrZToStr = Left$(s, Len(s) - 1)
End Function


Public Function GetElement(ByVal strList As String, ByVal strDelimiter As String, ByVal lngNumColumns As Long, ByVal lngRow As Long, ByVal lngColumn As Long) As String

    Dim lngCounter As Long

    ' Append delimiter text to the end of the list as a terminator.
    strList = strList & strDelimiter

    lngColumn = IIf(lngRow = 0, lngColumn, (lngRow * lngNumColumns) + lngColumn)

    For lngCounter = 0 To lngColumn - 1

        strList = Mid$(strList, InStr(strList, strDelimiter) + Len(strDelimiter), Len(strList))

        If Len(strList) = 0 Then
            GetElement = ""
            Exit Function
        End If

    Next lngCounter

    GetElement = Left$(strList, InStr(strList, strDelimiter) - 1)

End Function


Private Function GetNumElements(ByVal strList As String, ByVal strDelimiter As String) As Integer

    Dim intElementCount As Integer

    If Len(strList) = 0 Then
        GetNumElements = 0
        Exit Function
    End If

    strList = strList & strDelimiter

    While InStr(strList, strDelimiter) > 0
        intElementCount = intElementCount + 1
        strList = Mid$(strList, InStr(strList, strDelimiter) + 1, Len(strList))
    Wend

       GetNumElements = intElementCount

End Function



'Private Sub Form_Load()
'    GetProcesses "explorer.exe"
'End Sub




