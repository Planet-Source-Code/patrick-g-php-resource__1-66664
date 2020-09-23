VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10425
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14175
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10425
   ScaleWidth      =   14175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   360
      Top             =   1920
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10080
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15162
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A2D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F446
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":545B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6972A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7E89C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":93A0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A8B80
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BDCF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D2E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E7FD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FD148
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FF4CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11463C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1297AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13E920
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":153A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":168C04
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   1535
      ButtonWidth     =   2699
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New Project"
            Key             =   "new"
            Object.ToolTipText     =   "Open a new project"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Load Project"
            Key             =   "load"
            Object.ToolTipText     =   "Load a project"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Options"
            Key             =   "options"
            Object.ToolTipText     =   "Options"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Resources"
            Key             =   "resources"
            Object.ToolTipText     =   "Find external resources"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Function"
            Key             =   "add"
            Object.ToolTipText     =   "Add a new function"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Library"
            Key             =   "library"
            Object.ToolTipText     =   "View your current library of functions and sniplets"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Check for updates"
            Key             =   "updates"
            Object.ToolTipText     =   "Check for new databses and updates"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&About"
            Key             =   "about"
            Object.ToolTipText     =   "About PHP Resource"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Exit"
            Key             =   "exit"
            Object.ToolTipText     =   "Quit PHP Resource"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.Timer OS 
      Interval        =   1
      Left            =   10680
      Top             =   6360
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   10155
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   24950
            MinWidth        =   10583
         EndProperty
      EndProperty
   End
   Begin VB.Menu menu 
      Caption         =   "&Main Menu"
      Begin VB.Menu new 
         Caption         =   "&New Project"
         Shortcut        =   ^N
      End
      Begin VB.Menu open 
         Caption         =   "&Open Project"
         Shortcut        =   ^O
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu register 
         Caption         =   "&Product Registration"
         Shortcut        =   ^R
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
         Shortcut        =   ^Z
      End
   End
   Begin VB.Menu options1 
      Caption         =   "&Options"
      Begin VB.Menu ontop 
         Caption         =   "&Always ontop"
         Checked         =   -1  'True
      End
      Begin VB.Menu taskbar 
         Caption         =   "&Show on taskbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu sys 
         Caption         =   "Use System Tray"
         Enabled         =   0   'False
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu options 
         Caption         =   "&Preferences"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu projects 
      Caption         =   "&Projects"
      Begin VB.Menu resources 
         Caption         =   "&Resources"
         Shortcut        =   {F5}
      End
      Begin VB.Menu library 
         Caption         =   "&Library"
         Shortcut        =   {F6}
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu add 
         Caption         =   "&Add Resources"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu manager 
      Caption         =   "&Project Manager"
      Begin VB.Menu db 
         Caption         =   "&Database"
         Begin VB.Menu import 
            Caption         =   "&Import database"
         End
         Begin VB.Menu export 
            Caption         =   "&Export Database"
         End
         Begin VB.Menu line5 
            Caption         =   "-"
         End
         Begin VB.Menu backup 
            Caption         =   "&Backup Database"
            Shortcut        =   {F12}
         End
      End
      Begin VB.Menu backupmanager 
         Caption         =   "&Backup Manager"
         Begin VB.Menu logs 
            Caption         =   "&Backup logs"
         End
         Begin VB.Menu restore 
            Caption         =   "&Restore a backup"
            Shortcut        =   +{F12}
         End
      End
      Begin VB.Menu line9 
         Caption         =   "-"
      End
      Begin VB.Menu filemanager 
         Caption         =   "&File Manager"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu helptopics 
         Caption         =   "&Help Topics"
      End
      Begin VB.Menu helpcontents 
         Caption         =   "&Help Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu line6 
         Caption         =   "-"
      End
      Begin VB.Menu updates 
         Caption         =   "&Check for updates"
      End
      Begin VB.Menu line7 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "&About"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim InfSettings As String

Private Sub exit_Click()
    'i think we all know what this does
        End
End Sub

Private Sub Form_Load()
On Error Resume Next
    
    'If PHP Resource is already running, theres no need to open a second
        If App.PrevInstance = True Then
            MsgBox "You cannot run 2 instances of PHP Resource.  I will now close the extra instsance", vbCritical, "Error"
            Timer1.Enabled = True
        End If
    
    'Setting a strict varible for the INI Path
        InfSettings = App.Path & "/settings/inf.ini"
    
    'Set title
        Me.Caption = "PHP Resource - " & Computername
        
    'window state
        Me.WindowState = GetFromINI("Settings", "windowstate", InfSettings)
        Me.Left = GetFromINI("Settings", "meleft", InfSettings)
        Me.Top = GetFromINI("Settings", "meTop", InfSettings)
        Me.Width = GetFromINI("Settings", "meWidth", InfSettings)
        Me.Height = GetFromINI("Settings", "meHeight", InfSettings)
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
                    
    'save window settings.  Only here for future versions.  This plays no part in this version
    'but if tampered with in the INI file it would make the dialog look ugly
        SaveToINI Me.WindowState, "Settings", "windowstate", InfSettings
        SaveToINI Me.Width, "Settings", "meWidth", InfSettings
        SaveToINI Me.Height, "Settings", "meHeight", InfSettings
        SaveToINI Me.Left, "Settings", "meleft", InfSettings
        SaveToINI Me.Top, "Settings", "meTop", InfSettings
          
    'save os uptime (to record records)
        SaveToINI Uptime, "OS", "Uptime", InfSettings
    
    'Now we kill the application
        End
End Sub

Private Sub OS_Timer()
'insert os uptime into status panel
        StatusBar1.Panels(1).Text = "OS Uptime:   " & Uptime
End Sub

Private Sub Timer1_Timer()
    'This timer is used only for killing a second instance of php resource
        End
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    'We now have to figure out what buttons are being pressed on the main dialog and direct
    'the buttons somewhere.  Nothing special
        Select Case Button.Key
            Case "updates"
                FrmUpdate.Show
            Case "exit"
                End
            Case "about"
                frmAbout.Show
            Case "library"
                frmFunctions.Show
            Case "options"
                FrmFOptions.Show
            Case "add"
                frmLogin.Show
            Case "load"
                CommonDialog1.DialogTitle = "PHP Resource - Please select function sript"
                CommonDialog1.Filter = "PHP Resource File (*.phpr)|*.phpr"
                CommonDialog1.ShowOpen
            Case "new"
                FrmNewProject.Show
            Case "resources"
                frmBrowser.Show
                frmBrowser.brwWebBrowser.Navigate "http://ca3.php.net/manual/en/"
                
        End Select
End Sub

Private Sub updates_Click()
    'Check for new updates or patches or new database
        FrmUpdate.Show
End Sub
