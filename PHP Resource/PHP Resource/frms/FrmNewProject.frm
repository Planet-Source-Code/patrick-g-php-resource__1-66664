VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmNewProject 
   Caption         =   "Create new project"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9630
   Icon            =   "FrmNewProject.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2990
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"FrmNewProject.frx":15162
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNewProject.frx":151E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNewProject.frx":2A356
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNewProject.frx":3F4C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNewProject.frx":5463A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNewProject.frx":697AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNewProject.frx":7E91E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu open_work 
         Caption         =   "&Open"
      End
      Begin VB.Menu save_work 
         Caption         =   "&Save Project"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu close 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "FrmNewProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    'resize text box with the form
    With RichTextBox1
        .Width = Me.Width - 120
        .Height = Me.Height - 800
    End With
End Sub


