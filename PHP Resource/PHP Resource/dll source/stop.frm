VERSION 5.00
Begin VB.Form frmstop 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "msg.title"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7605
   Icon            =   "stop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label stopmsg 
         Caption         =   "Label1"
         Height          =   1095
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   6255
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "stop.frx":15162
         Top             =   360
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmstop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Ontop Me
End Sub
