VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SuperBAT Dos Commands"
   ClientHeight    =   5160
   ClientLeft      =   3735
   ClientTop       =   4095
   ClientWidth     =   6945
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   6945
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "INSERT COMMAND"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   3960
      TabIndex        =   8
      Top             =   1560
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Browse"
      Enabled         =   0   'False
      Height          =   735
      Left            =   4440
      TabIndex        =   7
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Enabled         =   0   'False
      Height          =   735
      Left            =   840
      TabIndex        =   6
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   3720
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   3720
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Use Selected Command"
      Height          =   1095
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "Form3.frx":0000
      Left            =   0
      List            =   "Form3.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Destination to copy/move to"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Original File"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   3360
      Width           =   2535
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim it
Private Sub Command1_Click()
Command4.Enabled = True
If List1.Selected(0) = True Or List1.Selected(1) = True Or List1.Selected(3) = True Or List1.Selected(4) = True Or List1.Selected(5) = True Then

    Text1.Enabled = True
    Text2.Enabled = True
    Command3.Enabled = True
        Command2.Enabled = True
        If List1.Selected(3) = True Or List1.Selected(4) Or List1.Selected(5) = True = True Then
        
            Command3.Enabled = False
            End If
            
Else
    Text1.Enabled = False
    Text2.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    
End If
End Sub

Private Sub Command2_Click()
If List1.Selected(4) = True Or List1.Selected(5) = True Then
    MsgBox ("You must type the directory in yourself, Sorry :-) ")
Else

CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.Filename
End If

End Sub

Private Sub Command3_Click()
CommonDialog1.ShowSave
Text2.Text = CommonDialog1.Filename
End Sub

Private Sub Command4_Click()
If List1.Selected(0) = True Then
it = 0
End If
If List1.Selected(1) = True Then
it = 1
End If
If List1.Selected(2) = True Then
it = 2
End If
If List1.Selected(3) = True Then
it = 3
End If
If List1.Selected(4) = True Then
it = 4
End If
If List1.Selected(5) = True Then
it = 5
End If

Select Case it
Case 0
Form2.Text1.Text = Form2.Text1.Text & "Move " & Text1.Text & " " & Text2.Text & Chr(10)
Case 1
Form2.Text1.Text = Form2.Text1.Text & "copy " & Text1.Text & " " & Text2.Text & Chr(10)
Case 2

Form2.Text1.Text = Form2.Text1.Text & "RUNDLL.EXE user.exe,exitwindows" & Chr(10)
Case 3
Form2.Text1.Text = Form2.Text1.Text & "Erase " & Text1.Text & Chr(10)
Case 4
Form2.Text1.Text = Form2.Text1.Text & "Deltree " & Text1.Text & Chr(10)
Case 5
Form2.Text1.Text = Form2.Text1.Text & "DIR " & Text1.Text & Chr(10)
End Select
Form3.Visible = False

End Sub

Private Sub Form_Load()
List1.AddItem "MOVE"
List1.AddItem "COPY"
List1.AddItem "SHUT DOWN WINDOWS"
List1.AddItem "ERASE"
List1.AddItem "DELTREE"
List1.AddItem "DIRECTORY"


End Sub

Private Sub Form_Unload(Cancel As Integer)
Command2.Enabled = False
Command3.Enabled = False
Command3.Enabled = False
End Sub
