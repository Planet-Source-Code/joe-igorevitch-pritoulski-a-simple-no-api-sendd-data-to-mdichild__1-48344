VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Joe's send text or data..."
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   6870
      TabIndex        =   4
      Top             =   4665
      Width           =   6930
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "You can olso use this code to send some setup information to an inactive MdiChild form. If you like this code please vote..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   60
         TabIndex        =   5
         Top             =   30
         Width           =   6765
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1740
      Top             =   1350
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   6870
      TabIndex        =   0
      Top             =   0
      Width           =   6930
      Begin VB.TextBox TextTMP 
         Height          =   345
         Left            =   3900
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   60
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Create MdiChild form"
         Height          =   345
         Left            =   180
         TabIndex        =   2
         Top             =   60
         Width           =   1755
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Show form2"
         Height          =   375
         Left            =   2070
         TabIndex        =   1
         Top             =   30
         Width           =   1695
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************
'I was looking for a simple way to send any type of
'data to an MdiChild from another form of the project
'with no API calls (whitch i personaly dislike).
'In this example I use a initialise event of te timer...
'Not an API, but simple and works for me...
'
'No limits on using, cause this is a freewear...
'
'IF YOU LIKE THIS CODE PLEASE VOTE!
'***************************************************
Private Sub Command1_Click()
'Loads form2
Form2.Show
End Sub

Private Sub Command2_Click()
 create_child
End Sub

Private Sub MDIForm_Load()
Timer1.Enabled = False ' disables timer
 create_child
End Sub

Private Sub Timer1_Timer()
'IMPORTANT - set timer interval more than 0
'this code actualy adds data to a MdiChild
'note if you are using a RichTextBox control change parametr ".Text" on ".TextRTF"
'to save formating, and don't forgrt to set a multiline property to true

ActiveForm.Text1.Text = ActiveForm.Text1.Text & TextTMP.Text
Timer1.Enabled = False
End Sub

Private Sub create_child()
'This peace of code is standart to create MdiChild
 Static lDocumentCount As Long 'Just a counter to make captions of the forms look different
    Dim frmD As Form1
    lDocumentCount = lDocumentCount + 1
    Set frmD = New Form1
    frmD.Caption = "Document " & lDocumentCount
    frmD.Show

End Sub
