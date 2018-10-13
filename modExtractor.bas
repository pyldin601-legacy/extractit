Attribute VB_Name = "modExtractor"
' Header format
Type container_header
    ch_sessions As Byte
    ch_nmb As Long
    ch_bkname As String * 19
End Type

' Object format
Type proc_format
    pf_name As String * 32
    pf_input As String * 256
    pf_output As String * 256
    pf_version As Integer
    pf_analyze As Integer
    pf_dirs As Integer
    pf_seserved As String * 256
End Type

' Statistic format
Type stat_format
    pf_containers As Double
    pf_files As Double
    pf_traffic As Double
    pf_traffic_ As Double
    pf_errors As Double
    pf_gots As Long
End Type



Global selected_proc As Integer
Global processors() As proc_format
Global statistic() As stat_format
Global settings_cnt As Integer


' Files format
Dim cd_sessionname As String
Dim cd_sessionleng As Long
Dim cd_sessiondata As String
Dim cd_header As container_header
Dim cd_buffer As String

Const cd_snamelen1 = 64
Const cd_snamelen2 = 128
Const BigFileLock = 2

Global fso As New FileSystemObject
Declare Function GetTickCount Lib "kernel32.dll" () As Long

Sub LoadSettings()

    On Error GoTo OnError
    Dim i As Integer, indx As Integer
    Dim l_sett As proc_format
    Dim new_settings_cnt As Integer
    Dim file_size As Integer
    
    settings_cnt = 0
    i = FreeFile
    Open LowPath(App.Path) & "winalfar.cfg" For Binary Access Read As #i
    Get #i, , file_size
    For indx = 1 To file_size
        Get #i, , l_sett
        new_settings_cnt = new_settings_cnt + 1
        ReDim Preserve processors(1 To new_settings_cnt) As proc_format
        ReDim Preserve statistic(1 To new_settings_cnt) As stat_format
        processors(new_settings_cnt) = l_sett
    Next indx
    settings_cnt = new_settings_cnt
    Close #i
    Exit Sub
OnError:
    LogError "load_proc", Err.Description
    Resume Next
    
End Sub


Sub AddProcessor(proc_name As String, proc_input As String, proc_output As String, proc_ver As Integer, proc_dirs As Integer, proc_analyze As Integer)

    On Error Resume Next
    
    Dim new_settings_cnt As Integer
    new_settings_cnt = settings_cnt + 1
    
    ReDim Preserve processors(1 To new_settings_cnt) As proc_format
    ReDim Preserve statistic(1 To new_settings_cnt) As stat_format
    
    With processors(new_settings_cnt)
        ' Copy settings
        .pf_name = proc_name
        .pf_input = proc_input
        .pf_output = proc_output
        .pf_version = proc_ver
        .pf_dirs = proc_dirs
        .pf_analyze = proc_analyze
    End With
    
    settings_cnt = new_settings_cnt

    Call SaveSettings

End Sub

Sub UpdateProcessor(proc_index As Integer, proc_name As String, proc_input As String, proc_output As String, proc_ver As Integer, proc_dirs As Integer, proc_analyze As Integer)

If proc_index > 0 Then
    With processors(proc_index)
        ' Copy settings
        .pf_name = proc_name
        .pf_input = proc_input
        .pf_output = proc_output
        .pf_version = proc_ver
        .pf_dirs = proc_dirs
        .pf_analyze = proc_analyze
    End With
    
    Call SaveSettings
End If

End Sub

Function RemoveProc(arrData() As proc_format, arrIndex As Integer) As Integer

    On Error Resume Next
    
    If arrIndex = UBound(arrData) Then
        ReDim Preserve arrData(arrIndex - 1)
    Else
        If arrIndex < LBound(arrData) Or arrIndex > UBound(arrData) Then
            MsgBox "Index is out of bounds!", vbExclamation
        Else
            Dim indx As Integer
            For indx = arrIndex To UBound(arrData) - 1
                arrData(indx) = arrData(indx + 1)
            Next indx
            ReDim Preserve arrData(UBound(arrData) - 1)
        End If
    End If
    
End Function

Function RemoveStat(arrData() As stat_format, arrIndex As Integer) As Integer

    On Error Resume Next
    
    If arrIndex = UBound(arrData) Then
        ReDim Preserve arrData(arrIndex - 1)
    Else
        If arrIndex < LBound(arrData) Or arrIndex > UBound(arrData) Then
            MsgBox "Index is out of bounds!", vbExclamation
        Else
            Dim indx As Integer
            For indx = arrIndex To UBound(arrData) - 1
                arrData(indx) = arrData(indx + 1)
            Next indx
            ReDim Preserve arrData(UBound(arrData) - 1)
        End If
    End If
    
End Function


Sub DeleteProcessor(proc_index As Integer)
    
    Dim new_settings_cnt As Integer
    
    
    new_settings_cnt = settings_cnt - 1
    Call RemoveProc(processors, proc_index)
    Call RemoveStat(statistic, proc_index)
    settings_cnt = new_settings_cnt
    
    Call SaveSettings
    
End Sub


Sub SaveSettings()
    
    On Error GoTo OnError
    Dim i As Integer, indx As Integer
    Dim file_size As Integer
    
    i = FreeFile
    file_size = settings_cnt
    
    Open LowPath(App.Path) & "winalfar.cfg" For Binary Access Write As #i
    
    Put #i, , file_size
    
    For indx = 1 To settings_cnt
        Put #i, , processors(indx)
    Next indx
    
    Close #i
    Exit Sub
OnError:
    LogError "save_proc", Err.Description
    
End Sub

Function ReadContainer(inFile As String, inVersion As Integer, inAnalyze As Integer, outFile As String) As stat_format

    On Error GoTo OnError
    
    Dim i As Integer
    Dim pktsmax As Integer
    Dim pktsleft As Integer
    Dim indx As Integer
    Dim sessions As Integer
    Dim traffic As Long
    Dim errors As Long
    Dim gots As Long
    
    Dim myData() As String
    
    Dim myDestPath As String
    Dim myFilename As String
    
    Dim filename_len As Integer
    Dim filename_suffix As String
    Dim filename_cut As Integer
    
    If inVersion Then
        filename_len = 128
        filename_suffix = ""
        filename_cut = 0
    Else
        filename_len = 64
        filename_suffix = ".eml"
        filename_cut = 21
    End If
    
    i = FreeFile
    
    
    Open inFile For Binary Access Read As #i
    
    Get #i, , cd_header
    
    Do
        cd_sessionname = Space(filename_len)
        Get #i, , cd_sessionname
        cd_buffer = Dezero(cd_sessionname)
        If Len(cd_buffer) > 0 Then
            myFilename = Left(cd_buffer, Len(cd_buffer) - filename_cut) & filename_suffix
            sessions = sessions + 1
            Get #i, , cd_sessionleng
            traffic = traffic + cd_sessionleng
            
            If cd_sessionleng < 32000 Then
                cd_sessiondata = Space(cd_sessionleng)
                Get #i, , cd_sessiondata
                ReDim myData(1 To 1)
                myData(1) = cd_sessiondata
            Else
                pktsmax = Fix(cd_sessionleng / 32000)
                pktsleft = cd_sessionleng Mod 32000
                If pktsleft > 0 Then
                    ReDim myData(1 To 1)
                    cd_sessiondata = Space(pktsleft)
                    Get #i, , cd_sessiondata
                    myData(1) = cd_sessiondata
                End If
                For indx = 1 To pktsmax
                    ReDim Preserve myData(1 To UBound(myData) + 1)
                    cd_sessiondata = Space(32000)
                    Get #i, , cd_sessiondata
                    myData(UBound(myData)) = cd_sessiondata
                Next indx
            End If
            If inAnalyze = 1 And frmExtract.mnuEnable.Checked Then
                If Not IsOurObject(inData(1), myDestPath) Then GoTo SkipFile
            Else
                myDestPath = outFile
            End If
            Select Case WriteFile(myFilename, myDestPath, myData)
            Case 2
                LogError "save_eml", "Can't save destination file '" & myFilename & "'"
                errors = errors + 1
            Case 1
                LogError "save_eml", "Can't create destination folder '" & outFile & "'"
                errors = errors + 1
            End Select
SkipFile:
            Erase myData
        End If
    Loop While Not EOF(i)
    Close #i
    
    With ReadContainer
        .pf_containers = 1
        .pf_files = sessions
        .pf_traffic = traffic
        .pf_errors = errors
        .pf_gots = gots
    End With
    
    LogExtraction PureName(inFile), sessions, traffic, Err.Description
    
    Kill inFile
    
    Exit Function
    
OnError:
    LogError "read_cont", Err.Description
    errors = errors + 1
    Resume Next
    
End Function

Function LogExtraction(cont_name As String, sess_count As Integer, bytes_total As Long, last_error As String)
    On Error Resume Next
    i = FreeFile
    Open LowPath(App.Path) & "winalfar.log" For Append As #i
    Print #i, "Name: " & cont_name & "    Files: " & Format(sess_count, "000") & "    Size: " & Format(bytes_total, "0")
    Close #i
End Function

Function LogError(section As String, error_description As String)
    On Error Resume Next
    i = FreeFile
    Open LowPath(App.Path) & "errors.log" For Append As #i
    Print #i, Format(Now, "dd.mm.yyyy HH:mm:ss") & " Section: " & section & " > " & error_description
    Close #i
End Function

Function PureName(inPath As String) As String
    Dim Path
    Path = Split(inPath, "\")
    PureName = Path(UBound(Path))
End Function

Function Dezero(inString As String) As String
    If InStr(inString, Chr(0)) Then
        Dezero = Left(inString, InStr(inString, Chr(0)) - 1)
    Else
        Dezero = Trim(inString)
    End If
End Function

Function WriteFile(inName As String, inPath As String, inData() As String) As Long

    On Error GoTo OnError
    Dim i As Integer, indx As Integer
    i = FreeFile
    
    ' MAKE FOLDERS
    If MyNewDir(inPath) Then
        WriteFile = 1
        Exit Function
    End If
    
    ' CLEAR DESTINATION
    Open LowPath(inPath) & inName For Output As #i: Close #i
    
    ' CREATE AND WRITE FILE
    Open LowPath(inPath) & inName For Binary Access Write As #i
    For indx = LBound(inData) To UBound(inData)
        Put #i, , inData(indx)
    Next indx
    Close #i
    
    ' CLOSE FILE
    WriteFile = 0
    Exit Function

OnError:
    WriteFile = 2
    
End Function


Sub Delay(ms As Long)

Dim tmptime As Long

tmptime = GetTickCount

Do: DoEvents: Loop While Not GetTickCount - tmptime > ms

End Sub

Public Function LowPath(inPath As String) As String
    If Right$(inPath, 1) = "\" Then LowPath = inPath Else LowPath = inPath + "\"
End Function

Function MyNewDir(inPath As String) As Long
    
    On Error GoTo OnError
    
    Dim vArray() As String, indx As Integer
    Dim myPath As String
    
    vArray = Split(inPath, "\")
    For indx = LBound(vArray) To UBound(vArray)
        myPath = myPath + vArray(indx) + "\"
        If vArray(indx) > "" Then If Not fso.FolderExists(myPath) Then MkDir myPath
    Next indx

    MyNewDir = 0
    Exit Function
    
OnError:
    MyNewDir = 1

End Function
