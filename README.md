# EnumProc

### Show process full paths without elevation

![image](https://github.com/fafalone/ProcessPaths/assets/7834493/b7d99212-1ce3-42f3-b7fe-f878e0ae9267)

If you've used programs like ProcessHacker or ProcessExplorer you might have noticed that you don't need to run them elevated to see the full paths of all the running processes. If you've ever tried this yourself with the standard methods of either reading the PEB or using the QueryFullProcessImageName API, you might have noticed this will fail for processes running as SYSTEM, because we're denied even the very limited PROCESS_QUERY_LIMITED_INFORMATION access right. So how can we find the full path anyway, when those are denied and the normal enum function CreateToolhelp32Snapshot only returns the file name? One way is to ask Windows itself, since it maintains a list of all process paths internally.

This is done with the NtQuerySystemInformation and the undocumented info class SystemProcessIdInformation. This allows us to specify a SYSTEM_PROCESS_ID_INFORMATION type with the ProcessId filled in, and it will fill a buffer with the full path to the image. 

```vba
    Private Type SYSTEM_PROCESS_ID_INFORMATION
        ProcessId As LongPtr
        ImageName As UNICODE_STRING
    End Type

    Private Type UNICODE_STRING
        uLength As Integer
        uMaximumLength As Integer
        pBuffer As LongPtr
    End Type

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
```

The final piece of the puzzle is the path post-processing: ConvertNtPathToDosPath. This is needed because the paths we receive here are in the format e.g. \Device\HarddiskVolume1\Windows\System32\crss.exe rather than the drive letters we're accustomed to. The way we translate these is by first creating a map with the MapVolumes function of every drive letter to the device path:

```vba
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
```

Then we just run each path through a find/replace of each map name to the corresponding letter.

```vba
    Private Function ConvertNtPathToDosPath(sPath As String) As String
    If sPath = "" Then Exit Function
    
    Dim i As Long
    ConvertNtPathToDosPath = sPath
    For i = 0 To UBound(VolMap)
        ConvertNtPathToDosPath = Replace$(ConvertNtPathToDosPath, VolMap(i).sName, VolMap(i).sLetter, 1, 1)
    Next
    End Function
```

All in all, this is more complicated than the traditional way, but there's plenty of situations where you're not able to run elevated but still want to display a list of processes and their paths--- or just their name, but need their paths to look up their icon.


-----

This code is compatible with both VB6 and twinBASIC, including 64bit compilation in the latter. The .twinproj is included, but you can reimport yourself to see the only change made is to use the smoother built in Anchor resizing rather than Form_Resize, as noted in the code.

There are no dependencies and this should work on all Windows versions XP and later.

UPDATE (2023 Nov 19) - Fix for garbage in system modules with no path. 
