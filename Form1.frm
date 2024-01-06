VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Process List"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9990
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Enum processes"
      Height          =   360
      Left            =   90
      TabIndex        =   1
      Top             =   165
      Width           =   1470
   End
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   600
      Width           =   9885
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If VBA7 = 0 Then
Private Enum LongPtr
    [_]
End Enum
#End If

Public Enum LMEM
    LMEM_FIXED = 0
    LMEM_MOVEABLE = &H2
    LHND = &H42
    LMEM_ZEROINIT = &H40
    lPtr = &H40
    NONZEROLHND = LMEM_MOVEABLE
    NONZEROLPTR = LMEM_FIXED
End Enum

Private Const INVALID_HANDLE_VALUE = -1&
Private Const MAX_PATH = 260

Private Type UNICODE_STRING
    uLength As Integer
    uMaximumLength As Integer
    pBuffer As LongPtr
End Type

Private Const SystemProcessIdInformation = 88
Private Type SYSTEM_PROCESS_ID_INFORMATION
    ProcessId As LongPtr
    ImageName As UNICODE_STRING
End Type

Private Enum TH32CS_Flags
    TH32CS_SNAPHEAPLIST = &H1
    TH32CS_SNAPPROCESS = &H2
    TH32CS_SNAPTHREAD = &H4
    TH32CS_SNAPMODULE = &H8
    TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
    TH32CS_INHERIT = &H80000000
End Enum
Private Type PROCESSENTRY32W
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As LongPtr
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile(0 To (MAX_PATH - 1)) As Integer
End Type



#If VBA7 Then
Private Declare PtrSafe Function LocalAlloc Lib "kernel32" (ByVal uFlags As LMEM, ByVal uBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function LocalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function NtQuerySystemInformation Lib "ntdll" (ByVal SystemInformationClass As Long, SystemInformation As Any, ByVal SystemInformationLength As Long, ReturnLength As Long) As Long
Private Declare PtrSafe Function QueryDosDeviceW Lib "kernel32" (ByVal lpDeviceName As LongPtr, ByVal lpTargetPath As LongPtr, ByVal ucchMax As Long) As Long
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (lpString As Any) As Long
Private Declare PtrSafe Function SysReAllocStringW Lib "oleaut32" Alias "SysReAllocString" (ByVal pBSTR As LongPtr, Optional ByVal pszStrPtr As LongPtr) As Long
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr)
Private Declare PtrSafe Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As TH32CS_Flags, ByVal th32ProcessID As Long) As LongPtr
Private Declare PtrSafe Function Process32FirstW Lib "kernel32" (ByVal hSnapshot As LongPtr, lppe As PROCESSENTRY32W) As Long
Private Declare PtrSafe Function Process32NextW Lib "kernel32" (ByVal hSnapshot As LongPtr, lppe As PROCESSENTRY32W) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As LongPtr)
#Else
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As LMEM, ByVal uBytes As LongPtr) As LongPtr
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare Function NtQuerySystemInformation Lib "ntdll" (ByVal SystemInformationClass As Long, SystemInformation As Any, ByVal SystemInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function QueryDosDeviceW Lib "kernel32" (ByVal lpDeviceName As LongPtr, ByVal lpTargetPath As LongPtr, ByVal ucchMax As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long
Private Declare Function SysReAllocStringW Lib "oleaut32" Alias "SysReAllocString" (ByVal pBSTR As LongPtr, Optional ByVal pszStrPtr As LongPtr) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr)
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As TH32CS_Flags, ByVal th32ProcessID As Long) As LongPtr
Private Declare Function Process32FirstW Lib "kernel32" (ByVal hSnapshot As LongPtr, lppe As PROCESSENTRY32W) As Long
Private Declare Function Process32NextW Lib "kernel32" (ByVal hSnapshot As LongPtr, lppe As PROCESSENTRY32W) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As LongPtr)
#End If


Private Type VolData
    sLetter As String
    sName As String
End Type
Private VolMap() As VolData
Private bSetVM As Boolean


Private Function GetProcessFullPathEx(pid As Long, pPath As String) As Long
    'Regular method can't get path of SYSTEM process
    'Note: API Returns NT path
    Dim Status As Long 'NTSTATUS
    Dim lpBuffer As LongPtr
    Dim spii As SYSTEM_PROCESS_ID_INFORMATION
    Dim sTemp As String
    Dim cbMax As Long: cbMax = MAX_PATH * 2 'LenB(Of Integer) ' * sizeof(WCHAR)
    Dim cbRet As Long
    lpBuffer = LocalAlloc(LMEM_FIXED, cbMax)
    ZeroMemory ByVal lpBuffer, cbMax
    spii.ProcessId = pid
    spii.ImageName.uMaximumLength = cbMax
    spii.ImageName.pBuffer = lpBuffer
    
    Status = NtQuerySystemInformation(SystemProcessIdInformation, spii, LenB(spii), cbRet)
    If NT_SUCCESS(Status) Then
        sTemp = LPWSTRtoStr(lpBuffer, False)
        If bSetVM = False Then
            MapVolumes
        End If
        pPath = ConvertNtPathToDosPath(sTemp)
    Else
        Debug.Print "GetProcessFullPathEx error, 0x" & Hex$(Status)
    End If
    
    LocalFree lpBuffer
    GetProcessFullPathEx = Status
End Function

Private Function LPWSTRtoStr(lPtr As LongPtr, Optional ByVal fFree As Boolean = True) As String
SysReAllocStringW VarPtr(LPWSTRtoStr), lPtr
If fFree Then
Call CoTaskMemFree(lPtr)
End If

End Function

Private Function NT_SUCCESS(ByVal Status As Long) As Boolean
    NT_SUCCESS = Status >= 0 'STATUS_SUCCESS
End Function

Private Sub MapVolumes()
'Map out \Device\Harddiskblahblah
Dim sDrive As String
Dim i As Long, j As Long
Dim sBuffer As String
ReDim VolMap(0)
Dim tmpMap() As VolData
Dim nMap As Long, nfMap As Long
Dim lIdx As Long
Dim lnMax As Long
Dim cb As Long
For lIdx = 0 To 25
    sDrive = Chr$(65 + lIdx) & ":"
    sBuffer = String$(1000, vbNullChar)
    cb = QueryDosDeviceW(StrPtr(sDrive), StrPtr(sBuffer), Len(sBuffer))
    If cb Then
        ReDim Preserve tmpMap(nMap)
        tmpMap(nMap).sLetter = sDrive
        tmpMap(nMap).sName = TrimNullW(sBuffer)
        nMap = nMap + 1
    End If
Next
'Next we need to sort the array so e.g. 10 will always come before 1
'We'll find the longest ones, add any of that length, then add any
'of 1 char shorter, until we've added all items
For i = 0 To (nMap - 1)
    If Len(tmpMap(i).sName) > lnMax Then lnMax = Len(tmpMap(i).sName)
Next i
ReDim VolMap(nMap - 1)
For i = lnMax To 1 Step -1
    For j = 0 To UBound(tmpMap)
        If Len(tmpMap(j).sName) = i Then
            VolMap(nfMap).sName = tmpMap(j).sName
            VolMap(nfMap).sLetter = tmpMap(j).sLetter
            nfMap = nfMap + 1
        End If
    Next j
    If nfMap = nMap Then Exit For
Next i
bSetVM = True
End Sub

'Then we can convert path names by running through the ones we got and replacing any
'occurences of them. The array is presorted so 10 comes before 1.
Private Function ConvertNtPathToDosPath(sPath As String) As String
If sPath = "" Then Exit Function

Dim i As Long
ConvertNtPathToDosPath = sPath
For i = 0 To UBound(VolMap)
    ConvertNtPathToDosPath = Replace$(ConvertNtPathToDosPath, VolMap(i).sName, VolMap(i).sLetter, 1, 1)
Next
End Function

Private Function TrimNullW(startstr As String) As String
TrimNullW = Left$(startstr, lstrlenW(ByVal StrPtr(startstr)))
End Function
    
Private Sub PostLog(sMsg As String)
'sMsg = "[" & Format$(Now, "Hh:nn:Ss") & "] " & sMsg & vbCrLf
Text1.Text = Text1.Text & sMsg & vbCrLf
End Sub
    
Private Sub Command1_Click()
Command1.Enabled = False
Text1.Text = ""
EnumProcess
Command1.Enabled = True
End Sub

Private Sub EnumProcess()
    Dim tProc As PROCESSENTRY32W
    Dim hSnap As LongPtr
    Dim hr As Long
    Dim sDisp As String
    Dim sPath As String
    Dim nProc As Long
    nProc = 0
    
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    
    If hSnap <> INVALID_HANDLE_VALUE Then
        tProc.dwSize = LenB(tProc)
        hr = Process32FirstW(hSnap, tProc)
        If hr > 0& Then
            Do While hr > 0&
                sDisp = "": sPath = ""
                If (tProc.th32ProcessID = 0) Then
                    PostLog "ProcId 0: [System idle process]"
                ElseIf (tProc.th32ProcessID = 4) Then
                    PostLog "ProcId 4: [System]"
                Else
                    sDisp = WCHARtoStr(tProc.szExeFile)
                    hr = GetProcessFullPathEx(tProc.th32ProcessID, sPath)
                    PostLog "ProcId " & tProc.th32ProcessID & " (" & sDisp & "): " & sPath
                End If

                nProc = nProc + 1
                hr = Process32NextW(hSnap, tProc)
            Loop
            PostLog "Done. Enumerated " & nProc & " processes."
        Else
            PostLog "Error calling Process32First, 0x" & Hex$(Err.LastDllError) & ", hSnapshot=" & hSnap
        End If
        CloseHandle hSnap
    Else
        PostLog "Error creating process snapshot."
    End If
End Sub

Private Function WCHARtoStr(aCh() As Integer) As String
Dim i As Long
Dim sz As String
Dim bStart As Boolean
For i = LBound(aCh) To UBound(aCh)
    If aCh(i) <> 0 Then
        sz = sz & ChrW(CLng(aCh(i)))
        bStart = True
    Else
        If bStart = False Then sz = sz & "0"
    End If
Next
If bStart = False Then
    WCHARtoStr = "<unknown or none>"
Else
    WCHARtoStr = sz
End If
End Function
    
Private Sub Form_Resize()
'Use anchors in tB for a smoother experience
#If TWINBASIC = 0 Then
Text1.Width = Me.Width - 220
Text1.Height = Me.Height - 1130
#End If
End Sub
