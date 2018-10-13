VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About application"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   4695
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   435
      Left            =   3180
      TabIndex        =   5
      Top             =   1920
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   180
         Picture         =   "frmAbout.frx":492A
         ScaleHeight     =   555
         ScaleWidth      =   615
         TabIndex        =   1
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "v0.0.000"
         Height          =   195
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WinAlfar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Copyright 2011 Roman Lakhtadyr."
      Height          =   555
      Left            =   240
      TabIndex        =   4
      Top             =   900
      Width           =   4215
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label2.Caption = "v" & Format(App.Major, "0") & "." & Format(App.Minor, "0") & "." & Format(App.Revision, "0")

End Sub
