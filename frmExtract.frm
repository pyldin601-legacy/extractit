VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.Form frmExtract 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12510
   Icon            =   "frmExtract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   12510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer trmMonitor 
      Interval        =   1000
      Left            =   8580
      Top             =   7260
   End
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "Start"
      Height          =   435
      Left            =   9720
      TabIndex        =   2
      Top             =   7260
      Width           =   1275
   End
   Begin VB.Timer trmProcMan 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   9120
      Top             =   7260
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   435
      Left            =   11100
      TabIndex        =   1
      Top             =   7260
      Width           =   1275
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6735
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   11880
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   1500
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Input"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Output"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Containers"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Files"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Traffic"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Errors"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Caught"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label lblNoObjects 
      Alignment       =   1  'Right Justify
      Caption         =   "Warning! There are no objects to capture!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8280
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   4035
   End
   Begin VB.Label lblGlobalTraffic 
      AutoSize        =   -1  'True
      Caption         =   "Traffic speed: "
      Height          =   195
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1020
   End
   Begin VB.Label lblFltObjects 
      AutoSize        =   -1  'True
      Caption         =   "Objects: "
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   630
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuImport 
         Caption         =   "Import folders..."
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export folders..."
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Begin VB.Menu mnuProcessors 
         Caption         =   "Processors..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSpider 
         Caption         =   "Spider..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnable 
         Caption         =   "Enable Spider"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuReload 
         Caption         =   "Reload objects"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim curr_proc As Integer
Dim global_traffic_a As Double
Dim global_traffic_b As Double
Dim prev_tick As Long
Dim prog_busy As Boolean

Const onetime_processings_max = 20

Private Sub cmdExit_Click()
    Unload Me
End Sub




Private Sub cmdSwitch_Click()
    
    trmProcMan.Enabled = Not trmProcMan.Enabled
    cmdSwitch.Caption = IIf(trmProcMan.Enabled, "Stop", "Start")
    
End Sub

Private Sub Form_Load()

    Me.Caption = App.ProductName & " ver. " & Format(App.Major, "0") & "." & Format(App.Minor, "0") & "." & Format(App.Revision, "0")
    Call LoadSettings
    Call LoadDBSettings
    Call UpdateList
    Call LoadObjects

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Are your sure want to stop and close WinAlfar?", vbQuestion + vbYesNo) = vbNo Then
        Cancel = 1
    Else
        End
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuEnable_Click()
    mnuEnable.Checked = Not mnuEnable.Checked
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuProcessors_Click()
    frmConfig.Show vbModal, Me

End Sub

Private Sub mnuReload_Click()
    Call LoadObjects
End Sub

Private Sub mnuSpider_Click()
    frmDBConnect.Show vbModal, Me
End Sub

Private Sub trmMonitor_Timer()

    Call UpdateList
    lblFltObjects.Caption = "Objects: " & Format(object_count, "0") & "/" & Format(themes_count, "0")
    lblGlobalTraffic.Caption = "Traffic speed: " & Format((global_traffic_a - global_traffic_b) / 1000 / ((GetTickCount - prev_tick) / 1000), "### ### ### ##0") & " kiB/s"
    global_traffic_b = global_traffic_a
    prev_tick = GetTickCount
    If object_count = 0 And mnuEnable.Checked Then
        frmExtract.lblNoObjects.Visible = True
    Else
        frmExtract.lblNoObjects.Visible = False
    End If


End Sub

Private Sub trmProcMan_Timer()

    If prog_busy Then Exit Sub

    prog_busy = True
    If settings_cnt = 0 Then Exit Sub
    curr_proc = curr_proc Mod settings_cnt
    curr_proc = curr_proc + 1
    Call ProcessThread(curr_proc)
    prog_busy = False
    
End Sub

Sub UpdateList()

    Dim li As ListItem, indx As Integer
    global_traffic_a = 0
    ListView1.ListItems.Clear
    For indx = 1 To settings_cnt
        Set li = ListView1.ListItems.Add
        li.Text = Trim(processors(indx).pf_name)
        li.SubItems(1) = Trim(processors(indx).pf_input)
        li.SubItems(2) = Trim(processors(indx).pf_output)
        li.SubItems(3) = Format(statistic(indx).pf_containers, "0")
        li.SubItems(4) = Format(statistic(indx).pf_files, "0")
        li.SubItems(5) = Format(statistic(indx).pf_traffic / 1000, "### ### ### ##0") & " kiB"
        li.SubItems(6) = Format(statistic(indx).pf_errors, "0")
        li.SubItems(7) = Format(statistic(indx).pf_gots, "0")
        global_traffic_a = global_traffic_a + statistic(indx).pf_traffic
    Next indx
    
    
End Sub

Sub ProcessThread(proc_index As Integer)
    
    On Error GoTo OnError
    
    Dim ifolder As folder
    Dim iFile As File
    Dim inFiles As Integer
    Dim oReturn As stat_format
    Dim subFolder As String
    
    Set ifolder = fso.GetFolder(Trim(processors(proc_index).pf_input))
    
    For Each iFile In ifolder.Files
        ' If InStr(iFile, "101805") Then Stop
        inFiles = inFiles + 1
        With processors(proc_index)
            subFolder = Format((inFiles - 1) Mod IIf(.pf_dirs, .pf_dirs, 1), "00")
            oReturn = ReadContainer(iFile.Path, .pf_version, .pf_analyze, LowPath(Trim(.pf_output)) & subFolder)
        End With
        With statistic(proc_index)
            .pf_containers = .pf_containers + oReturn.pf_containers
            .pf_errors = .pf_errors + oReturn.pf_errors
            .pf_files = .pf_files + oReturn.pf_files
            .pf_traffic = .pf_traffic + oReturn.pf_traffic
            .pf_gots = .pf_gots + oReturn.pf_gots
        End With
        If inFiles = onetime_processings_max Or trmProcMan.Enabled = False Then Exit For
        DoEvents
    Next
    Exit Sub
OnError:
    LogError "proc_thread#" & Format(proc_index, "0"), Err.Description
    Resume Next
    
End Sub
