VERSION 5.00
Begin VB.Form frmDBConnect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Configuration"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   ControlBox      =   0   'False
   Icon            =   "frmDBConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2700
      TabIndex        =   11
      Top             =   1980
      Width           =   1275
   End
   Begin VB.CommandButton txtSave 
      Caption         =   "Save"
      Height          =   435
      Left            =   1320
      TabIndex        =   10
      Top             =   1980
      Width           =   1275
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtDB 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "X"
      TabIndex        =   7
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtLogin 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtConnector 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Spider path:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label Label4 
      Caption         =   "Database:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Login:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Connection:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmDBConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtConnector.Text = dbc_conn
    txtLogin.Text = dbc_login
    txtPassword.Text = dbc_passw
    txtDB.Text = dbc_db
    txtPath.Text = dbc_spider
End Sub

Private Sub txtSave_Click()
    dbc_conn = txtConnector.Text
    dbc_login = txtLogin.Text
    sdb_passw = txtPassword.Text
    dbc_db = txtDB.Text
    dbc_spider = txtPath.Text
    Call SaveDBSettings
    Call LoadObjects
    Unload Me
End Sub
