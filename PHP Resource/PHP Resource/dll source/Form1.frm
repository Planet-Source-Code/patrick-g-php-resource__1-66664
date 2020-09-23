VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFF80&
      Caption         =   "Delete"
      Height          =   195
      Left            =   3720
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      Caption         =   "test"
      Height          =   255
      Left            =   3480
      TabIndex        =   31
      Top             =   6720
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2880
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Mp3s"
      InitDir         =   "C:\"
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFFF80&
      Caption         =   "Read"
      Height          =   195
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFF80&
      Caption         =   "Edit"
      Height          =   195
      Left            =   3000
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF80&
      Caption         =   "Save"
      Height          =   195
      Left            =   1560
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   8
      Left            =   1320
      TabIndex        =   27
      Top             =   4080
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "del"
      Height          =   195
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFF80&
      Caption         =   "Decode"
      Height          =   195
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFF80&
      Caption         =   "Encode"
      Height          =   195
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   7
      Left            =   1320
      TabIndex        =   19
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF80&
      Caption         =   "read"
      Height          =   195
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF80&
      Caption         =   "save"
      Height          =   195
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   6
      Left            =   1320
      TabIndex        =   15
      Top             =   2880
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Caption         =   "&About"
      Height          =   255
      Left            =   3480
      MaskColor       =   &H8000000D&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   5
      Left            =   1320
      TabIndex        =   5
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   4
      Left            =   1320
      TabIndex        =   4
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Database:"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   26
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Create an error"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   120
      MouseIcon       =   "Form1.frx":27A2
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Create custom splash screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   120
      MouseIcon       =   "Form1.frx":2AAC
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   6360
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Create custom about dialog"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   120
      MouseIcon       =   "Form1.frx":2DB6
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Encryption:"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   18
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Registry:"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   14
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "It is just a very basic example of what you can do with a DLL.  If enough votes come in i'll make this 99x better."
      ForeColor       =   &H00FFFF80&
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Windows Uptime:"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   11
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Filesize:"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dir Exists:"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   9
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Computer name:"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File Exists:"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "App Version:"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'To load the DLL in to your project do the following!
'
'Goto Project Menu > Refrences then hit browse
'Look for the MultiFuncDLL.dll file on your computer and hit ok!
'Then hit OK again
'
'Now in your project i'll use a "Form Load EG"
'Dim the following into your project
'
'dim Dll as multifuncdll.extras
'
'public sub form_load()
'set dll = new extras
'
'From there you can call any function like this
'dll.osuptime (And whatever else is in the selection menu
'
'end sub

'Loads the DLL
Dim dll As Extras

Public Function AppDir() As String
'This is just used incase someone installs this into C:\ for some reason
'Saves errors from happening =)
    If Right$(App.Path, 1) <> "/" Then
        AppDir = App.Path & "/"
    Else
        AppDir = App.Path
    End If
End Function
Private Sub Command1_Click()
'Show the about dialog on this dll program
    dll.Show
End Sub

Private Sub Command10_Click()
'This allows you to read entrys from the database very simple eh
Text1.Item(8) = dll.DatabaseRead("firstname")
End Sub

Private Sub Command11_Click()
'This just shows the open dialog and places whatever file selected into text(9) and in text(4) it
'Calculates the filesize to show you how it all works
    CD.ShowOpen
    Text1.Item(9) = CD.FileName
    Text1.Item(4) = dll.FileSize(CD.FileName)
End Sub

Private Sub Command12_Click()
'Such a simple way to stop mp3's.. the 5000 is how long to wait befor stopping it
    dll.StopMP3 5000
End Sub

Private Sub Command15_Click()
'Show open dialog again
    CD.ShowOpen
    Text1.Item(10) = CD.FileName
    Text1.Item(4) = dll.FileSize(CD.FileName)
End Sub

Private Sub Command16_Click()
'Displays access denied message.  Just something some ppl use, thought i'd make it easier heh
    dll.AccessDeniedMessage
End Sub

Private Sub Command17_Click()

dll.CreateStopMsg "Login Failed", "This is a test message"

End Sub

Private Sub Command2_Click()
'This will allow you to delete serton or all settings from registry for your program
    dll.DelRegistry "Folder", "Subfolder", "Keyname"
    Text1.Item(6) = ""
End Sub

Private Sub Command3_Click()
'This is where you can save settings the the registry
    dll.SaveRegistry "Folder", "Subfolder", "Keyname", Text1.Item(6)
End Sub

Private Sub Command4_Click()
'this is where you can grab settings from registry for your program or something
    Text1.Item(6) = dll.GetRegistry("Folder", "Subfolder", "Keyname")
End Sub

Private Sub Command5_Click()
'This allows us to add a new entry to the database
    dll.DatabaseAddNew "firstname", Text1.Item(8)
End Sub

Private Sub Command6_Click()
'This is the area that encrypts the text from the text box very simple
    Text1.Item(7) = dll.Encryption(Text1.Item(7), 0)
End Sub

Private Sub Command7_Click()
'This is the area that de-crypts the text in the text box very simple as well
    Text1.Item(7) = dll.Encryption(Text1.Item(7), 1)
End Sub

Private Sub Command8_Click()
'This will give us the ability to edit current entry in database
    dll.DatabaseEdit
End Sub

Private Sub Command9_Click()
'this will give us the ability to delete current shown entry in the database
    dll.DatabaseDelete "firstname"
End Sub

Public Sub form_load()
'This loads the dll to be used
Set dll = New Extras

'now we will place all items in the text boxes as an example

'This will load the App Version
Text1.Item(0) = dll.AppVersion

'This will center the form
'dll.CenterForm

'This checks if any file exists
If dll.FileCheck("C:\windows\command.com") = True Then
    'If it exists it will say so here
    Text1.Item(1) = "Command.com Exists"
Else
    'If it does not exist it will say so here
    Text1.Item(1) = "Command.com Doesn't Exist"
End If

'This displays any computer name
Text1.Item(2) = dll.Computername

'This checks if a directory exists
If dll.FileDirCheck("C:\windows") = True Then
    'If the dir exists it will say so here
    Text1.Item(3) = "Windows Dir Exists"
Else
    'If the dir does not exist it will say so here
    Text1.Item(3) = "Windows Dir Does Not Exist"
End If

'This here will calculate filesize of any file you can think of
Text1.Item(4) = dll.FileSize("C:\windows\system32\diskcopy.dll")

'This displays the  windows 9.x\me\2K\XP uptime
Text1.Item(5) = dll.OsUptime

'If there is something saved in the TEST registry it will place it in text1.item(6)
Text1.Item(6) = dll.GetRegistry("Folder", "Subfolder", "Keyname")

'This is nothing, just places text in text1.item(7) giving an example on how to encrypt and decrypt
Text1.Item(7) = "Click Encode, to encrypt this message"

'open the database file
dll.DatabaseOpen AppDir & "database.mdb", "testTable"

'Load Custom msgbox which will show a welcome screen basicly
dll.CreateCustomMessageBox "Thanks!", "Hope you like this program!!!", "If you honestly think this program is making progress and think it can get somewhere in helping people, please vote and leave feedback for me, Thanks again people."
End Sub

Private Sub Form_Unload(Cancel As Integer)
'This is how simple you can get an exit message
    dll.DatabaseClose
    dll.ExitMessage "Enter your exit message here", "Message box caption", "If clicked yes", "If clicked no", True
End Sub

Private Sub Label3_Click()
'This is a very useful yet quick way to create an about screen for your program
    dll.CreateAboutScreen "Enter caption here", "Program title", "Description here", "1.0.0", "Enter your warning here"
End Sub

Private Sub Label4_Click()
'This is a very useful yet quick way to create a splash screen for your program
    dll.CreateSplashScreen 5000, vbGreen, "Patrick G.", "Topshotter.com", "(c)2001-2002 Topshotter Inc", "This program is intended for learning purposes only!", "Windows 9.x\me\2k\XP", "Enter program title here", "1.0.0"
End Sub

Private Sub Label5_Click()
'This is just an example of the debug window, below is how
'I crashed the program and debugged it
On Error GoTo fail:
Dim ErrorTime As Integer
    ErrorTime = 1.08972401284702E+33
fail:
    dll.ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "Trying to see if i can crash the program"
End Sub

