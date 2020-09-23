VERSION 5.00
Begin VB.Form frmFTPAdd 
   Caption         =   "LiveUpdate configuration"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDontSee 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   4800
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Caption         =   "Before LiveUpdate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   22
      Top             =   3480
      Width           =   4215
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse ..."
         Enabled         =   0   'False
         Height          =   350
         Left            =   3000
         TabIndex        =   15
         Top             =   690
         Width           =   1095
      End
      Begin VB.CheckBox ChkExecute 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtExcecute 
         BackColor       =   &H80000004&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lblCheck 
         Caption         =   "E&xecute"
         Height          =   255
         Left            =   400
         TabIndex        =   12
         Top             =   380
         Width           =   2175
      End
   End
   Begin VB.Frame fraDownload 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   21
      Top             =   2160
      Width           =   4215
      Begin VB.TextBox txtDestination 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         TabIndex        =   11
         Text            =   "c:\MyRep\file.xxx"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtFiles 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         TabIndex        =   9
         Text            =   "Folder/file.xxx"
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label6 
         Caption         =   "&Destination :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "&FTP files :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "FTP config"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Text            =   "MySoft 2.0"
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Text            =   "ftp.server.com"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Text            =   "username"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "&Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   405
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "FTP &Server :"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   765
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "&Password :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1485
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "&User :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1125
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   19
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   18
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblDontSee 
      Caption         =   "See &config button next time"
      Height          =   255
      Left            =   550
      TabIndex        =   16
      Top             =   4830
      Width           =   2295
   End
End
Attribute VB_Name = "frmFTPAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChkExecute_Click()
If Me.ChkExecute = 1 Then
   Me.cmdBrowse.Enabled = True
   Me.txtExcecute.BackColor = 16777215
   Me.txtExcecute.Locked = False
Else
   Me.cmdBrowse.Enabled = False
   Me.txtExcecute.BackColor = 12632256
   Me.txtExcecute.Locked = True
End If
End Sub

Private Sub cmdBrowse_Click()
Dim strFiles As String
strFiles = GetOpenFile(CurDir, "Choisir un fichier", True, False)
If strFiles <> "" Then
   Me.txtExcecute = strFiles
End If
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Call WriteWinIniParam(App.Path & "\LiveUpdate.ini", "LiveUpdate", "USER", txtUser)
  Call WriteWinIniParam(App.Path & "\LiveUpdate.ini", "LiveUpdate", "HOST", txtHost)
  Call WriteWinIniParam(App.Path & "\LiveUpdate.ini", "LiveUpdate", "PASS", Encrypte(txtPass))
  Call WriteWinIniParam(App.Path & "\LiveUpdate.ini", "LiveUpdate", "AppName", txtName)
  
  Call WriteWinIniParam(App.Path & "\LiveUpdate.ini", "LiveUpdate", "FILES1", txtFiles)
  Call WriteWinIniParam(App.Path & "\LiveUpdate.ini", "LiveUpdate", "PATH1", txtDestination)
    
  Call WriteWinIniParam(App.Path & "\LiveUpdate.ini", "LiveUpdate", "EXECUTE", ChkExecute)
  Call WriteWinIniParam(App.Path & "\LiveUpdate.ini", "LiveUpdate", "EXECFILES", txtExcecute)
  Call WriteWinIniParam(App.Path & "\LiveUpdate.ini", "LiveUpdate", "SEECONFIG", chkDontSee)
  frmLiveUpdate.Caption = "LiveUpdate " & txtName
  Unload Me
  
End Sub

Private Sub Form_Load()
  txtName = GetIniParam(App.Path & "\LiveUpdate.ini", "LiveUpdate", "AppName")
  txtUser = GetIniParam(App.Path & "\LiveUpdate.ini", "LiveUpdate", "USER")
  txtHost = GetIniParam(App.Path & "\LiveUpdate.ini", "LiveUpdate", "HOST")
  txtPass = Decrypt(GetIniParam(App.Path & "\LiveUpdate.ini", "LiveUpdate", "PASS"))
  
  txtFiles = GetIniParam(App.Path & "\LiveUpdate.ini", "LiveUpdate", "FILES1")
  txtDestination = GetIniParam(App.Path & "\LiveUpdate.ini", "LiveUpdate", "PATH1")
  
  ChkExecute = GetIniParam(App.Path & "\LiveUpdate.ini", "LiveUpdate", "EXECUTE")
  txtExcecute = GetIniParam(App.Path & "\LiveUpdate.ini", "LiveUpdate", "EXECFILES")
  chkDontSee = GetIniParam(App.Path & "\LiveUpdate.ini", "LiveUpdate", "SEECONFIG")
End Sub

Private Sub lblCheck_Click()
  If Me.ChkExecute = 1 Then
     Me.ChkExecute = 0
  Else
     Me.ChkExecute = 1
  End If
End Sub

Private Sub lblDontSee_Click()
  If Me.chkDontSee = 1 Then
     Me.chkDontSee = 0
  Else
     Me.chkDontSee = 1
  End If
End Sub
