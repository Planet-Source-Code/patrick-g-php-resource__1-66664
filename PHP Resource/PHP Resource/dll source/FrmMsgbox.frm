VERSION 5.00
Begin VB.Form FrmMsgbox 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   Icon            =   "FrmMsgbox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "FrmMsgbox.frx":0442
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label lbldesc 
      Alignment       =   2  'Center
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   5175
   End
End
Attribute VB_Name = "FrmMsgbox"
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
