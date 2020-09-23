VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   6540
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   7395
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3864.048
   ScaleMode       =   0  'User
   ScaleWidth      =   6943.504
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   38
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":15162
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":15467
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":158B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":15B9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":15D97
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":1606B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":164E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":16888
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":16C90
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":16DB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":16FE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":171A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":174A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":17776
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":17BA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":17F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":18377
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":1866F
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":1888E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":18CF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":19005
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":192EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":195AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":19BA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":19C49
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":19EEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":1A204
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":1A68F
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":1A924
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":1AC0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":1B08C
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":1B28D
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":1B562
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":1B990
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":1BBF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":1F55C
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":1F847
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":349B9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   4800
      TabIndex        =   6
      Top             =   6120
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   390
      Left            =   6120
      TabIndex        =   5
      Top             =   6120
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Caption         =   "Login Details"
      Height          =   5775
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   4935
      Begin VB.CheckBox Check1 
         Caption         =   "Remember me please"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   4080
         Width           =   2175
      End
      Begin MSComctlLib.ImageCombo ImageCombo1 
         Height          =   330
         Left            =   1800
         TabIndex        =   7
         Top             =   3000
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "ImageList1"
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "#"
         TabIndex        =   1
         Top             =   3480
         Width           =   2235
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label Label1 
         Height          =   2175
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Password:"
         Height          =   270
         Index           =   1
         Left            =   960
         TabIndex        =   3
         Top             =   3480
         Width           =   840
      End
      Begin VB.Label lblLabels 
         Caption         =   "&User Name:"
         Height          =   270
         Index           =   0
         Left            =   840
         TabIndex        =   2
         Top             =   3015
         Width           =   840
      End
   End
   Begin VB.Image Image1 
      Height          =   6750
      Left            =   0
      Picture         =   "frmLogin.frx":49B2B
      Top             =   -120
      Width           =   2250
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Set vars for database
    Dim rs As DAO.Recordset
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    
'login verification true or false var
    Public login As Boolean


Private Sub cmdCancel_Click()
    'unload login screen
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    
    'Move recordset back to first item in database
        rs.MoveFirst
    
    'Setting some varibles
        Dim i As Integer
           
    'Checking to see if we are going to remember the password
        If Check1.Value = "1" Then
            SaveToINI Encryption(txtPassword.Text, 0), "Login", "Password", App.Path & "\settings\inf.ini"
            SaveToINI "1", "Login", "remember", App.Path & "\settings\inf.ini"
            SaveToINI ImageCombo1.Text, "Login", "username", App.Path & "\settings\inf.ini"
        Else
            SaveToINI "No Password Saved", "Login", "Password", App.Path & "\settings\inf.ini"
            SaveToINI "0", "Login", "remember", App.Path & "\settings\inf.ini"
            SaveToINI "No Username", "Login", "username", App.Path & "\settings\inf.ini"
        End If
        
    'We are now running the script to determain we needs to be done
               For i = 1 To rs.RecordCount
                    If rs("username") = ImageCombo1.Text Then
                        If txtPassword.Text = rs("password") Then
                            frmAdd.Show
                            Unload Me
                        Else
                            Label2.Caption = "Incorrect password"
                            txtPassword.SetFocus
                        End If
                    Else
                        rs.MoveNext
                    End If
                Next i
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    'set the login as false to start.  This is basicly the heart of the verification
        login = False
    
    'Load users table
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\db\database.mdb")
        Set rs = db.OpenRecordset("users", dbOpenTable)
    
    'insert text into caption
        Label1.Caption = "Hello," & vbCrLf & vbCrLf & "Before i can let you continue, i need to have you login to the system.  The reason i need for you to login is because the users of this program have worked very long and hard to setup and organize this database.  If you'd like access, the administrator can add you as a new user." & vbCrLf & vbCrLf & "Thank you"


    'load the image combo box
        With ImageCombo1.ComboItems
            Dim Cnt As Integer
            For Cnt = 1 To rs.RecordCount
                .add , rs("username"), rs("username"), 38
                rs.MoveNext
            Next Cnt
        End With
    
    'select first or saved username so we dont have a blank drop down
        ImageCombo1.SelectedItem = ImageCombo1.ComboItems(1)
    
    'Check if you asked to remember you
        Dim LogonType, LogonName, LogonPw As String
            LogonType = GetFromINI("login", "remember", App.Path & "\settings\inf.ini")
            LogonName = GetFromINI("Login", "username", App.Path & "\settings\inf.ini")
            LogonPw = GetFromINI("Login", "password", App.Path & "\settings\inf.ini")
        
        If LogonType = "1" Then
            Check1.Value = "1"
            ImageCombo1.Text = LogonName
            txtPassword.Text = Encryption(LogonPw, 1)
        Else
            Check1.Value = "0"
        End If
End Sub

Private Sub txtPassword_Change()
    Label2.Caption = ""
End Sub
