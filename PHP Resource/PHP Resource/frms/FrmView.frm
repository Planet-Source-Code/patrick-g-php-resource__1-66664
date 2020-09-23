VERSION 5.00
Begin VB.Form FrmView 
   Caption         =   "PHP Resource Viewer"
   ClientHeight    =   9165
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12960
   Icon            =   "FrmView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   12960
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.Menu exit 
      Caption         =   "&Exit"
   End
   Begin VB.Menu n 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu fd 
      Caption         =   ""
   End
   Begin VB.Menu ad 
      Caption         =   ""
   End
   Begin VB.Menu a 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu b 
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
   Begin VB.Menu eeee 
      Caption         =   "This is only a file viewer, note editor"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "FrmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Resize()

    Text1.Height = Me.Height - 800
    Text1.Width = Me.Width - 120

End Sub
