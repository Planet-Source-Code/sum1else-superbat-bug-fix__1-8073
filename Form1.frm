VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   5175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox pics 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   720
      TabIndex        =   1
      Text            =   "0"
      Top             =   3240
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   1560
      Top             =   2040
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5175
      ScaleWidth      =   5895
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   2880
         Top             =   1680
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   1560
         Top             =   2040
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
Form1.Height = Picture1.Height

Form1.Width = Picture1.Width
End Sub

Private Sub Image1_Click()
MsgBox ("Email Tyler Robbins at Tyler@robbins.to")
End Sub



Private Sub Timer1_Timer()
pics.Text = pics.Text + 1
Form1.Height = Picture1.Height
Form1.Width = Picture1.Width

Select Case pics.Text
Case 1
picview = "1.bmp"
Case 2
picview = "2.bmp"
Case 3
picview = "3.bmp"
Case 4
picview = "4.bmp"
Case 5
picview = "5.bmp"
Case 6
picview = "6.bmp"
Case 7
picview = "7.bmp"
Case 8
picview = "8.bmp"
Case 9
picview = "9.bmp"
Timer1.Enabled = False
Timer2.Enabled = True

End Select
Set Picture1.Picture = _
LoadPicture(picview)
Image1.Enabled = True



End Sub

Private Sub Timer2_Timer()

Form2.Visible = True
Form2.Enabled = True
Form1.Enabled = False
Form1.Visible = False

End Sub
