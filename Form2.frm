VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Send text or data..."
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3510
   LinkTopic       =   "Form2"
   ScaleHeight     =   1800
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Place data on MdiChild"
      Height          =   525
      Left            =   330
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   30
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   30
      Width           =   3375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'sends data to a text control
MDIForm1.TextTMP.Text = Text1.Text
MDIForm1.Timer1.Enabled = True 'initialise timer
'Unload me 'you can add this code to unload send data form "Form2"
End Sub
