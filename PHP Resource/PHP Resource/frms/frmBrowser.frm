VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowser 
   ClientHeight    =   10695
   ClientLeft      =   3060
   ClientTop       =   3645
   ClientWidth     =   14820
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10695
   ScaleWidth      =   14820
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   10440
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "frmBrowser.frx":15162
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19288
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1236
            MinWidth        =   1236
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1588
            MinWidth        =   1588
            TextSave        =   "1:37 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   1323
      ButtonWidth     =   1852
      ButtonHeight    =   1164
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Go Back"
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Go Forward"
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Stop"
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Refresh"
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Home"
            Key             =   "Home"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Search"
            Key             =   "Search"
            Object.ToolTipText     =   "Search"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   4080
      ExtentX         =   7197
      ExtentY         =   4471
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer timTimer 
      Interval        =   5
      Left            =   3000
      Top             =   2400
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   14820
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   750
      Width           =   14820
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   1920
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2A2D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2A5B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2A898
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2AB7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2AE5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2B13E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu exit 
      Caption         =   "&Exit"
   End
   Begin VB.Menu a 
      Caption         =   ""
   End
   Begin VB.Menu b 
      Caption         =   ""
   End
   Begin VB.Menu c 
      Caption         =   ""
   End
   Begin VB.Menu d 
      Caption         =   ""
   End
   Begin VB.Menu e 
      Caption         =   ""
   End
   Begin VB.Menu f 
      Caption         =   ""
   End
   Begin VB.Menu g 
      Caption         =   ""
   End
   Begin VB.Menu h 
      Caption         =   ""
   End
   Begin VB.Menu ada 
      Caption         =   "&This is strictly for PHP Resource.  To browse the web, use Internet Explorer"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Set our public varibles
    Public StartingAddress As String
    Dim mbDontNavigateNow As Boolean

Private Sub brwWebBrowser_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    On Error Resume Next
    'Set values for the progress bar and show user the page progress
        frmProgress.ProgressBar1.Max = ProgressMax
        frmProgress.ProgressBar1.Value = 0
        
        If frmProgress.ProgressBar1.Value <> ProgressMax Then
            frmProgress.Show
            frmProgress.ProgressBar1.Value = Progress
        Else
            frmProgress.Hide
        End If
End Sub

Private Sub exit_Click()
    'obvious what this does
        Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    'Lets get everything in order on first launch
        Me.Show
        tbToolBar.Refresh
        Form_Resize
End Sub

Private Sub brwWebBrowser_DownloadComplete()
    'change the title bar to what we are viewing just like IE
        Me.Caption = "PHP Resource - " & brwWebBrowser.LocationName
End Sub

Private Sub Form_Resize()
    'Lets resize the form as its resized.  We dont want the browser to stay same size as window changes
        brwWebBrowser.Width = Me.ScaleWidth
        brwWebBrowser.Height = Me.ScaleHeight - (picAddress.Top + picAddress.Height) + 350
End Sub


Private Sub timTimer_Timer()
    'This is where we determain if web browser is busy or not, if not lets tell you it isn't
        If brwWebBrowser.Busy = False Then
            timTimer.Enabled = False
            Me.Caption = brwWebBrowser.LocationName
            StatusBar1.Panels(2).Text = "Done."
        Else
            Me.Caption = "Working..."
            StatusBar1.Panels(2).Text = brwWebBrowser.LocationURL
        End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
    'enable the timer to allow user to see whats happening
        timTimer.Enabled = True
     'this is where we figure out what button is pressed
        Select Case Button.Key
            Case "Back"
                brwWebBrowser.GoBack
            Case "Forward"
                brwWebBrowser.GoForward
            Case "Refresh"
                brwWebBrowser.Refresh
            Case "Home"
                brwWebBrowser.GoHome
            Case "Search"
                brwWebBrowser.GoSearch
            Case "Stop"
                timTimer.Enabled = False
                brwWebBrowser.Stop
                Me.Caption = brwWebBrowser.LocationName
        End Select
End Sub

