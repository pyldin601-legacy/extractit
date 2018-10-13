VERSION 5.00
Begin VB.Form frmExtract 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinAlfar v1.1"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Caption         =   "Out:"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "In:"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   615
   End
End
Attribute VB_Name = "frmExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Monitor As Boolean
Public nContainers As Long
Public nFiles As Long
Public nKB As Double

Sub LogEvent(inMessage As String)
Text1.Text = Text1.Text + inMessage + vbCrLf
If Len(Text1.Text) > 10000 Then Text1.Text = Right(Text1.Text, 10000)
Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    Command2.Enabled = True
    Call StartMonitor
End Sub

Private Sub Command2_Click()
    Monitor = False
    Command2.Enabled = False
    Command1.Enabled = True
End Sub

Sub StartMonitor()

    On Error GoTo LogError

    Dim inDir As String
    Dim outDir As String
    Dim FName As String
    Dim fdirs As Integer
    Dim fdirnmb As Integer
    Dim dirEnd As String
    Dim fsoTotal As Long
    
    Dim fsoFile As File
    Dim fsoDir As folder

    
    Monitor = True
    inDir = LowPath(Text2.Text)
    outDir = LowPath(Text3.Text)
    fdirs = Val(Text4.Text)
    fdirnumb = -1
    
    Do
        Set fsoDir = fso.GetFolder(inDir)
	fsoTotal = fsoDir.Files.Count
        For Each fsoFile In fsoDir.Files
            If fso.FileExists(fsoFile.Path) Then
                If fdirs > 0 Then fdirnmb = (fdirnmb + 1) Mod fdirs
                dirEnd = outDir & IIf(fdirs > 0, Format(fdirnmb, "00") & "\", "")
                If Not fso.FolderExists(dirEnd) Then MyMkDir dirEnd
		LogEvent "Processing " & fsoFile.Name
                modExtractor.ReadContainer fsoFile.Path, dirEnd
                lContainers.Caption = Format(nContainers, "0")
                lFiles.Caption = Format(nFiles, "0")
                lKB.Caption = Format(nKB / 1024, "### ### ### ##0.0") & " kB"
		DoEvents
            End If
            If Not Monitor Then Exit Do
            if fsoDir.Files.Count>0 then Picture1.Line (0, 0)-(100 / fsoTotal * fsoDir.Files.Count, 1), , BF
        Next fsoFile
        If Not Monitor Then Exit Do
        modExtractor.Delay 1000
    Loop
    
    Exit Sub
    
LogError:
    LogEvent "Error: " & Err.Description
    Err.Clear
    Resume Next

End Sub

Private Sub Form_Load()
    
    Text2.Text = GetSetting("ExtractIt", "Pathes", "In", "")
    Text3.Text = GetSetting("ExtractIt", "Pathes", "Out", "")
    Text4.Text = GetSetting("ExtractIt", "Pathes", "Cnt", "")

    Me.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ExtractIt", "Pathes", "In", Text2.Text)
    Call SaveSetting("ExtractIt", "Pathes", "Out", Text3.Text)
    Call SaveSetting("ExtractIt", "Pathes", "Cnt", Text4.Text)
    End
End Sub

