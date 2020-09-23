VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFunctions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PHP Library"
   ClientHeight    =   6795
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   11655
   Icon            =   "frmFunctions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   4320
      Top             =   6360
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Close functions"
      Height          =   375
      Left            =   10200
      TabIndex        =   13
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Script"
      Height          =   375
      Left            =   8760
      TabIndex        =   12
      Top             =   6360
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9340
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   10560
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctions.frx":15162
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctions.frx":2A2D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctions.frx":3F446
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctions.frx":545B8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11160
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctions.frx":6972A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctions.frx":7E89C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctions.frx":93A0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctions.frx":A8B80
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctions.frx":AAF02
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctions.frx":C0074
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctions.frx":D51E6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   6
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   5
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   4
         Top             =   300
         Width           =   2055
      End
   End
   Begin MSComctlLib.ImageCombo ImageCombo2 
      Height          =   330
      Left            =   8400
      TabIndex        =   7
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      Text            =   "ImageCombo2"
      ImageList       =   "ImageList2"
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   4680
      TabIndex        =   8
      Top             =   480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      Text            =   "ImageCombo1"
      ImageList       =   "ImageList1"
   End
   Begin VB.Label lblCount 
      Caption         =   "Label4"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   6360
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmFunctions.frx":EA358
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Please select a function from the list below:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Select Type:"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Select view:"
      Height          =   255
      Left            =   8400
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Set vars for database
Dim rs As DAO.Recordset
Dim db As DAO.Database
Dim ws As DAO.Workspace

Private Sub Command1_Click()
    'view selected function from the list
    'I've created a public function in a module to make this a bit quicker for future versions
        If ImageCombo2.SelectedItem.Key = "null" Then
            MsgBox "I'm unable to open that for you.  You need to select a program which you wish to view it in." & vbCrLf & "Please select a viewer from the drop down box at the top", vbInformation, "Error"
            Exit Sub
        End If
        
        FileResource ImageCombo2.SelectedItem.Key, ListView1.SelectedItem.Text
End Sub

Private Sub Command3_Click()
    'Obviously what this does
        Unload Me
End Sub

Private Sub Form_Load()

    Dim rcount As String

    'using my dll open the database
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\db\database.mdb")
        Set rs = db.OpenRecordset("functions", dbOpenTable)
        
    'Count number of records in database so we can create a a loop down below
        rcount = rs.RecordCount

    'initialize the image combo with special keys
        With ImageCombo1.ComboItems
            .add , "null", "Please select a function from the list", 1
            .add , "myfunctions", "My Functions", 2
            .add , "preloaded", "Pre-Loaded functions", 3
            .add , "scripts", "Pre-Loaded scripts", 4
            .add , "database", "Database Functions", 5
            .add , "howto", "How to's", 6
            .add , "all", "View All code and scripts", 7
        End With

    'initialize the combobox with the options of editors with special keys
        With ImageCombo2.ComboItems
            .add , "null", "Please select a view from the list", 2
            .add , "notepad", "I'd like to view this in notepad", 1
            .add , "wordpad", "I'd like to view this in wordpad", 3
            .add , "phpresource", "I don't know?  You pick for me", 4
        End With


    'Select which items in the combo boxes i want to be displayed upon first opening window
        ImageCombo1.SelectedItem = ImageCombo1.ComboItems(1)
        ImageCombo2.SelectedItem = ImageCombo2.ComboItems(4)

    'create the headers for the list
        ListView1.ColumnHeaders.Clear
        ListView1.ColumnHeaders.add , , "ID", ListView1.Width / 11
        ListView1.ColumnHeaders.add , , "Type", ListView1.Width / 8
        ListView1.ColumnHeaders.add , , "Function", ListView1.Width / 6
        ListView1.ColumnHeaders.add , , "Comment", ListView1.Width / 2
        ListView1.ColumnHeaders.add , , "Author", ListView1.Width / 9
        ListView1.View = lvwReport

    'We are now creating the loop to access the database (table: Functions)
    'so that we can place each function in the list.
    'here we go
        Dim iCnt As Integer
            iCnt = GetFromINI("Database", "selection", App.Path & "\Settings\inf.ini")
            ImageCombo1.SelectedItem = ImageCombo1.ComboItems(iCnt)
            
    'Now we call the sort function
        Call ImageCombo1_Click
        
    'Display function count
        lblCount.Caption = ListView1.ListItems.Count & " function(s) listed above"


End Sub


Private Sub Form_Unload(Cancel As Integer)

    'save settings when closing
        Dim SaveVar As Integer
    
        With ImageCombo1.SelectedItem
            If .Key = "null" Then
                SaveVar = 1
            ElseIf .Key = "myfunctions" Then
                SaveVar = 2
            ElseIf .Key = "preloaded" Then
                SaveVar = 3
            ElseIf .Key = "scripts" Then
                SaveVar = 4
            ElseIf .Key = "database" Then
                SaveVar = 5
            ElseIf .Key = "howto" Then
                SaveVar = 6
            ElseIf .Key = "all" Then
                SaveVar = 7
            End If
        End With

    'now that we fiured out which option was left when closed we save it
        SaveToINI SaveVar, "database", "selection", App.Path & "\settings\inf.ini"
        SaveToINI ImageCombo1.SelectedItem.Key, "database", "view", App.Path & "\settings\inf.ini"

End Sub

Private Sub ImageCombo1_Click()
    
    'This is where we sort which type of scripts the user (you) want to view
        Dim Cnt, i As Integer
            ListView1.ListItems.Clear
    
        If ImageCombo1.SelectedItem.Key = "all" Then
            rs.MoveFirst
                For Cnt = 1 To rs.RecordCount
                    With ListView1.ListItems.add(, , Cnt, 1, 1)
                        .SubItems(1) = StrConv(rs("type"), vbProperCase)
                        .SubItems(2) = StrConv(rs("function"), vbProperCase)
                        .SubItems(3) = StrConv(rs("comment"), vbProperCase)
                        .SubItems(4) = StrConv(rs("author"), vbProperCase)
                        rs.MoveNext
                    End With
                Next Cnt
                Exit Sub
        End If
            
    'Move to first record in database and clean the listview
            rs.MoveFirst
            ListView1.ListItems.Clear
            
                For Cnt = 1 To rs.RecordCount
                
                    If rs("type") = ImageCombo1.SelectedItem.Key Then
                    
                        With ListView1.ListItems.add(, , Cnt, 1, 1)
                            .SubItems(1) = StrConv(rs("type"), vbProperCase)
                            .SubItems(2) = StrConv(rs("function"), vbProperCase)
                            .SubItems(3) = StrConv(rs("comment"), vbProperCase)
                            .SubItems(4) = StrConv(rs("author"), vbProperCase)
                        End With
                    End If
                
                    rs.MoveNext
                Next Cnt
        
End Sub

Private Sub ImageCombo2_Click()
    'Instead of getting errors, we are going to try and locate these files
    'if they are selected right away.  never know, some people just delete
    'these files and rather not use them, so that is why i offer the option
    'to open within php resource
    
    On Error Resume Next
    
        If ImageCombo2.SelectedItem.Key = "wordpad" Then
            If FileExists("C:\Program Files\Windows NT\Accessories\wordpad.exe") = False Then
                MsgBox "Wordpad doesn't seem to be installed on this system, please try another editor", vbInformation, "Editor"
                ImageCombo2.SelectedItem = ImageCombo1.ComboItems(1)
                Exit Sub
            End If
        End If
        
        If ImageCombo2.SelectedItem.Key = "notepad" Then
            If FileExists("C:\windows\notepad.exe") = False Then
                MsgBox "Notepad doesn't seem to be installed on this system, please try another editor", vbInformation, "Editor"
                ImageCombo2.SelectedItem = ImageCombo2.ComboItems(1)
                Exit Sub
            End If
        End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    'if clicked we sort the functions
        If ListView1.SortOrder = lvwAscending Then
            ListView1.SortOrder = lvwDescending
        Else
            ListView1.SortOrder = lvwAscending
        End If

End Sub

Private Sub ListView1_DblClick()
    
    'if user double clicks function, open it.  A lot of people have a happit of double clicking
    
        If ImageCombo2.SelectedItem.Key = "null" Then
            MsgBox "I'm unable to open that for you." & vbCrLf & "Please select a viewer from the drop down box at the top", vbInformation, "Error"
            Exit Sub
        End If
    
        FileResource ImageCombo2.SelectedItem.Key, ListView1.SelectedItem.Text
End Sub

Private Sub Timer1_Timer()
    'constantly update the file count to display how many items are displayed
        lblCount.Caption = ListView1.ListItems.Count & " function(s) listed above"
End Sub
