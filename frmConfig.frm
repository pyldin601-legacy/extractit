VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Processors"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close"
      Height          =   435
      Left            =   7080
      TabIndex        =   19
      Top             =   4140
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selected folder statistic"
      Height          =   3855
      Left            =   3300
      TabIndex        =   4
      Top             =   120
      Width           =   5055
      Begin VB.CheckBox chkDB 
         Caption         =   "Analyze files in this folder"
         Height          =   255
         Left            =   1020
         TabIndex        =   21
         Top             =   2040
         Width           =   3195
      End
      Begin VB.TextBox txtOutputs 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   4320
         TabIndex        =   20
         Top             =   1200
         Width           =   495
      End
      Begin VB.CheckBox chkVersion 
         Caption         =   "Containers has a new version"
         Height          =   255
         Left            =   1020
         TabIndex        =   18
         Top             =   1680
         Width           =   3195
      End
      Begin VB.TextBox txtTraffic 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0"
         Top             =   3300
         Width           =   1575
      End
      Begin VB.TextBox txtFiles 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0"
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtContainers 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0"
         Top             =   2460
         Width           =   1575
      End
      Begin VB.TextBox txtOutput 
         Height          =   345
         Left            =   1020
         TabIndex        =   9
         Top             =   1200
         Width           =   3195
      End
      Begin VB.TextBox txtInput 
         Height          =   345
         Left            =   1020
         TabIndex        =   8
         Top             =   780
         Width           =   3795
      End
      Begin VB.TextBox txtName 
         Height          =   345
         Left            =   1020
         TabIndex        =   7
         Top             =   360
         Width           =   3795
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Total traffic:"
         Height          =   255
         Left            =   1620
         TabIndex        =   15
         Top             =   3360
         Width           =   1515
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Files extracted:"
         Height          =   195
         Left            =   1620
         TabIndex        =   13
         Top             =   2940
         Width           =   1515
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Processed:"
         Height          =   195
         Left            =   1620
         TabIndex        =   11
         Top             =   2520
         Width           =   1515
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Output:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   1260
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Input:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   420
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input folders"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton Command2 
         Caption         =   "Update"
         Height          =   315
         Left            =   1980
         TabIndex        =   17
         Top             =   3360
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Delete"
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   3360
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add..."
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   3360
         Width           =   855
      End
      Begin VB.ListBox lstFolders 
         Height          =   2985
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub UpdateProcessors()

    Dim indx As Integer
    Dim last_index As Integer
    
    last_index = lstFolders.ListIndex
    lstFolders.Clear
    
    For indx = 1 To settings_cnt
        lstFolders.AddItem Trim(processors(indx).pf_name)
    Next indx
    
    If last_index < lstFolders.ListCount Then lstFolders.ListIndex = last_index
    
    Call UpdateStatistic

End Sub

Sub UpdateStatistic()
    
    If lstFolders.ListIndex < 0 Or lstFolders.ListIndex > settings_cnt - 1 Then
        txtName.Text = "": txtInput.Text = "": txtOutput.Text = "": txtContainers.Text = "": txtFiles.Text = "": txtTraffic.Text = "": chkVersion.Value = 0: txtOutputs.Text = ""
    Else
        With processors(lstFolders.ListIndex + 1)
            txtName.Text = Trim(.pf_name)
            txtInput.Text = Trim(.pf_input)
            txtOutput.Text = Trim(.pf_output)
            chkVersion.Value = IIf(.pf_version, 1, 0)
            chkDB.Value = IIf(.pf_analyze, 1, 0)
            txtOutputs.Text = Format(.pf_dirs, "0")
        End With
        With statistic(lstFolders.ListIndex + 1)
            txtContainers.Text = Format(.pf_containers, "0")
            txtFiles.Text = Format(.pf_files, "0")
            txtTraffic.Text = Format(.pf_traffic / 1024, "### ### ##0.#") & " kB"
        End With
    End If
    Command2.Enabled = False

End Sub

Private Sub chkDB_Click()
    Command2.Enabled = True
End Sub

Private Sub chkVersion_Click()
    Command2.Enabled = True
End Sub

Private Sub cmdAdd_Click()
    frmAdd.Show vbModal, Me
    Call UpdateProcessors
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    If lstFolders.ListIndex >= 0 Then
        If MsgBox("Delete selected item?", vbQuestion + vbYesNo) = vbYes Then
            DeleteProcessor lstFolders.ListIndex + 1
            Call UpdateProcessors
        End If
    End If
End Sub

Private Sub Command2_Click()
    Call UpdateProcessor(lstFolders.ListIndex + 1, txtName.Text, txtInput.Text, txtOutput.Text, chkVersion.Value, Val(txtOutputs.Text), chkDB.Value)
End Sub

Private Sub Form_Load()
    Call UpdateProcessors
End Sub

Private Sub lstFolders_Click()
    Call UpdateStatistic
End Sub

Private Sub txtInput_Change()
    Command2.Enabled = True
End Sub

Private Sub txtName_Change()
    Command2.Enabled = True
End Sub

Private Sub txtOutput_Change()
    Command2.Enabled = True
End Sub

Private Sub txtOutputs_Change()
    Command2.Enabled = True
End Sub
