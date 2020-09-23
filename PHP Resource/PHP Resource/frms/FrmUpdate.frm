VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PHP Resource - Check for updates"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4905
   Icon            =   "FrmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   9840
      Width           =   255
      ExtentX         =   450
      ExtentY         =   450
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Top             =   10080
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   720
      TabIndex        =   7
      Top             =   9720
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Check for updates"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   5040
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4575
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   8070
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"FrmUpdate.frx":15162
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   9360
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   9000
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   8640
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   8280
      Width           =   1335
   End
End
Attribute VB_Name = "FrmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'set my varibles before we being for the updates
        On Error Resume Next
        Dim answer As Integer
        Dim data, InIPath, AppVer As String
        
    'set the progress bar to 0 and disable button until we are done
        Command1.Enabled = False
        ProgressBar1.Visible = True
        ProgressBar1.Value = 0

    'Check if update has already been preformed
        If Command1.Caption = "Exit" Then
            Unload Me
            Exit Sub
        End If

    'basic richtextbox stuff, we are just putting text letting you know what im doing with the update
        RichTextBox1.SelColor = vbRed
        RichTextBox1.SelText = "Checking For New Updates..." & vbCrLf
        ProgressBar1.Value = 20

    'Setting the URL where we are getting the information for the update
    'this is the heart of the update
        data = GetUrlSource("http://updates.thinkhosted.com")
    
    'I guess you can say we are about 40% done, so we set the progressbar to 40%
        ProgressBar1.Value = 40

    'now of course we want to hide errors, so if something is wrong and i can't connect
    'to the update server to check for updates, instead of crashing the program
    'we tell the user that we cannot connect to the server and to try again later
        If data = "" Then
            RichTextBox1.SelText = "" & vbCrLf & vbCrLf
            RichTextBox1.SelText = "Error Server Down, Please check back at a later time" & vbCrLf & "Sorry for the Inconvience"
            RichTextBox1.SelText = ""
            Command1.Caption = "Exit"
            ProgressBar1.Value = 100
            Command1.Enabled = True
            Exit Sub
        End If

    'We are now telling you more information in the textbox
        ProgressBar1.Value = 60
        RichTextBox1.SelColor = vbRed
        RichTextBox1.SelText = "Update Check Complete..." & vbCrLf
        
        RichTextBox1.SelColor = vbBlack
        RichTextBox1.SelText = "=====================================" & vbCrLf
        
        RichTextBox1.SelColor = vbRed
        RichTextBox1.SelText = "Looking to see if an Update is Needed..." & vbCrLf

        RichTextBox1.SelColor = vbBlack
        RichTextBox1.SelText = "=====================================" & vbCrLf
        
        RichTextBox1.SelColor = vbBlue
        RichTextBox1.SelText = "Reviewing Data" & vbCrLf
        
        RichTextBox1.SelColor = vbBlack
        RichTextBox1.SelText = "=====================================" & vbCrLf & vbclrf & vbCrLf

    'remove old updates (last time we checked) deleting the file
        Kill App.Path & "/settings/updates.ini"
    
    'Now we are going to open the new file that has been downloaded
        Open App.Path & "/settings/updates.ini" For Append As 1
            Print #1, data
        Close 1

    'We are reading the INI file for updates and setting our progressbar to 80%
        ProgressBar1.Value = 80
        Text1.Text = GetFromINI("Updates", "Version", App.Path & "/settings/updates.ini")
        Text2.Text = GetFromINI("Updates", "Programmer", App.Path & "/settings/updates.ini")
        Text3.Text = GetFromINI("Updates", "FileSize", App.Path & "/settings/updates.ini")
        Text4.Text = GetFromINI("Updates", "URL", App.Path & "/settings/updates.ini")
        Text5.Text = GetFromINI("Updates", "date", App.Path & "/settings/updates.ini")
        Text6.Text = GetFromINI("Updates", "comments", App.Path & "/settings/updates.ini")

    'If for some reason the update url is empty we have to give an error
        If Text1.Text = "" Then
            RichTextBox1.SelText = "Invalid URL and or Unable to locate updates file." & vbCrLf & "Please ensure that the server you are trying to access" & vbCrLf & "has updates.ini located somewhere on the server"
            Command1.Caption = "Exit"
            Command1.Enabled = True
            Exit Sub
        End If


    'Back to our richtextbox stuff
        RichTextBox1.SelText = ""
        RichTextBox1.SelColor = &H8000&
        RichTextBox1.SelText = "Version: " & Text1 & vbCrLf
    
        RichTextBox1.SelColor = &H8000&
        RichTextBox1.SelText = "Programmer: " & Text2 & vbCrLf
        
        RichTextBox1.SelColor = &H8000&
        RichTextBox1.SelText = "FileSize: " & Text3 & vbCrLf
        
        RichTextBox1.SelColor = &H8000&
        RichTextBox1.SelText = "Last updated: " & Text5 & vbCrLf
        
        RichTextBox1.SelColor = &H8000&
        RichTextBox1.SelText = "Comments: " & Text6 & vbCrLf
                
    
    'We need to grab the version directly from the file (This program itself)
        AppVer = App.Major & "." & App.Minor & "." & App.Revision

    'now to verify if there is a new version avalible or not.  If there is to allow user to
    'download the update if they'd like to or download the new databases for php resources
        If Text1.Text > AppVer Then
            Notontop Me
                answer = MsgBox("There is a New Update, do you wish to download?", vbQuestion + vbYesNo, "Update PSC Chat?")

                If answer = vbYes Then
                    WebBrowser1.Navigate Text4.Text
                    ProgressBar1.Value = 100
                    Command1.Caption = "Exit"
                    Command1.Enabled = True
                    Exit Sub
                End If
        
                If answer = vbNo Then
                    ProgressBar1.Value = 100
                    Command1.Caption = "Exit"
                    Command1.Enabled = True
                    Exit Sub
                End If
        End If

    'If you have recent version or updated database then we let you know and allow you to exit
    'without the offer of an update
        RichTextBox1.SelText = "" & vbCrLf & vbCrLf
        RichTextBox1.SelText = "You have the most recent version of PHP Resource." & vbCrLf & "Version:( " & AppVer & " )" & vbCrLf
        RichTextBox1.SelText = "There is no need for an update"
        Command1.Caption = "Exit"
        Command1.Enabled = True
        ProgressBar1.Value = 100

End Sub

Private Sub Form_Load()
    'We want to set this form ontop of others because updates are always very important
        ontop Me
End Sub
