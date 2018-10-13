VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add new folder"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Flags"
      Height          =   1275
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   6135
      Begin VB.CheckBox chkDBProc 
         Caption         =   "Analyze files in this folder"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox chkVersion 
         Caption         =   "Containers has a new version"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   3600
      TabIndex        =   3
      Top             =   3600
      Width           =   1275
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   435
      Left            =   4980
      TabIndex        =   2
      Top             =   3600
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Folder name:"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.TextBox txtOutput 
         Height          =   315
         Left            =   1140
         TabIndex        =   9
         Top             =   1380
         Width           =   3555
      End
      Begin VB.TextBox txtDirs 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   5460
         TabIndex        =   8
         Top             =   1380
         Width           =   495
      End
      Begin VB.TextBox txtInput 
         Height          =   315
         Left            =   1140
         TabIndex        =   7
         Top             =   960
         Width           =   4815
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1140
         TabIndex        =   1
         Top             =   420
         Width           =   4815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Count:"
         Height          =   195
         Left            =   4860
         TabIndex        =   10
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Output:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Input:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1020
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
    If txtName.Text > "" And txtInput.Text > "" And txtOutput.Text > "" Then
        Call AddProcessor(txtName.Text, txtInput.Text, txtOutput.Text, chkVersion.Value, Val(txtDirs.Text), chkDBProc.Value)
        Unload Me
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

