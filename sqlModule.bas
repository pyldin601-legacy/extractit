Attribute VB_Name = "sqlModule"
Type objects_type
    ot_object As String
    ot_theme As Integer
End Type

Type themes_type
    tt_theme As Integer
    tt_path As String
End Type

Global object_array() As objects_type
Global object_count As Integer

Global themes_array() As themes_type
Global themes_count As Integer

Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strCN As String
Dim strRS As String
Dim strIT As Integer

Global dbc_conn As String
Global dbc_login As String
Global dbc_passw As String
Global dbc_db As String
Global dbc_spider As String


Sub LoadDBSettings()
    dbc_conn = GetSetting("WinAlfar", "DB", "Connection", "dbs_sc")
    dbc_login = GetSetting("WinAlfar", "DB", "Login", "primary")
    dbc_passw = GetSetting("WinAlfar", "DB", "Password", "primary")
    dbc_db = GetSetting("WinAlfar", "DB", "DB", "spy")
    dbc_spider = GetSetting("WinAlfar", "DB", "Path", "\\MARS\Spider")
End Sub

Sub SaveDBSettings()
    SaveSetting "WinAlfar", "DB", "Connection", dbc_conn
    SaveSetting "WinAlfar", "DB", "Login", dbc_login
    SaveSetting "WinAlfar", "DB", "Password", dbc_passw
    SaveSetting "WinAlfar", "DB", "DB", dbc_db
    SaveSetting "WinAlfar", "DB", "Path", dbc_spider
End Sub


Sub LoadObjects()
    
    On Error GoTo OnError
    
    strCN = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & dbc_conn & ";Mode=Read;Initial Catalog=" & dbc_db
    strRS = "select id_theme,name_object from spider_objects"
    
    cn.Open strCN, dbc_login, dbc_passw
    
    With rs
        .CursorLocation = adUseClient
        .Open strRS, cn, adOpenDynamic, adLockOptimistic
    End With

    object_count = 0
    
    If rs.State = 1 Then
        For strIT = 1 To rs.RecordCount
            object_count = object_count + 1
            ReDim Preserve object_array(1 To object_count)
            With object_array(object_count)
                .ot_object = rs("name_object")
                .ot_theme = rs("id_theme")
            End With
            rs.MoveNext
        Next strIT
        rs.Close
    End If
    
    strRS = "select id_theme,name_theme from spider_themes"
    
    With rs
        .CursorLocation = adUseClient
        .Open strRS, cn, adOpenDynamic, adLockOptimistic
    End With

    themes_count = 0
    
    If rs.State = 1 Then
        For strIT = 1 To rs.RecordCount
            themes_count = themes_count + 1
            ReDim Preserve themes_array(1 To themes_count)
            With themes_array(themes_count)
                .tt_path = rs("name_theme")
                .tt_theme = rs("id_theme")
            End With
            rs.MoveNext
        Next strIT
        rs.Close
    End If
    
    If cn.State > 0 Then cn.Close
    Exit Sub
OnError:
    LogError "sql_connect", Err.Description
    Resume Next

End Sub


Function IsOurObject(inData As String, outPath) As Boolean
    
    On Error Resume Next
    Dim cursr As Integer
    Dim cursr2 As Integer
    
    IsOurObject = False
    
    For cursr = 1 To object_count
        If InStr(inData, object_array(cursr).ot_object) > 0 Then
            For cursr2 = 1 To themes_count
                If themes_array(cursr2).tt_theme = object_array(cursr).ot_theme Then
                    IsOurObject = True
                    outPath = LowPath(dbc_spider) & themes_array(cursr2).tt_path
                    Exit Function
                End If
            Next cursr2
            Exit For
        End If
    Next cursr
    
End Function
