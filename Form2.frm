VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SuperBat"
   ClientHeight    =   9480
   ClientLeft      =   450
   ClientTop       =   1410
   ClientWidth     =   10110
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   10110
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   2640
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   16748
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Code"
      TabPicture(0)   =   "Form2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "text1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Timer1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Help"
      TabPicture(1)   =   "Form2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label14"
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(3)=   "Label11"
      Tab(1).Control(4)=   "Label10"
      Tab(1).Control(5)=   "Label9"
      Tab(1).Control(6)=   "Label8"
      Tab(1).Control(7)=   "Label7"
      Tab(1).Control(8)=   "Label6"
      Tab(1).Control(9)=   "Label5"
      Tab(1).Control(10)=   "Label4"
      Tab(1).Control(11)=   "Label3"
      Tab(1).Control(12)=   "Label2"
      Tab(1).ControlCount=   13
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3720
         Top             =   2400
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Close Batch File COMPLETELY"
         Height          =   1095
         Left            =   7560
         TabIndex        =   23
         Top             =   8280
         Width           =   2415
      End
      Begin RichTextLib.RichTextBox text1 
         Height          =   9015
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   15901
         _Version        =   393217
         TextRTF         =   $"Form2.frx":0038
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Display Text"
         Height          =   1095
         Left            =   7560
         TabIndex        =   17
         Top             =   7080
         Width           =   2415
      End
      Begin VB.CommandButton Command7 
         Caption         =   "DOS Command Selector"
         Height          =   1095
         Left            =   7560
         TabIndex        =   15
         Top             =   5880
         Width           =   2415
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Clear Screen"
         Height          =   975
         Left            =   7560
         TabIndex        =   12
         Top             =   4800
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Pause - press ANY key to continue"
         Height          =   975
         Left            =   7560
         TabIndex        =   4
         Top             =   3720
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         Caption         =   "EXECUTE PROGRAM"
         Height          =   975
         Left            =   7560
         TabIndex        =   3
         Top             =   2640
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Make DOS Text Commands Visible"
         Height          =   855
         Left            =   7560
         TabIndex        =   2
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Make DOS Text Commands NOT Visible"
         Height          =   855
         Left            =   7560
         TabIndex        =   1
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label14 
         Caption         =   $"Form2.frx":00BA
         Height          =   1695
         Left            =   -69960
         TabIndex        =   21
         Top             =   4440
         Width           =   4575
      End
      Begin VB.Label Label13 
         Caption         =   "Deltree -  Will delete the directory tree, everything in the directory including subtitles"
         Height          =   735
         Left            =   -69960
         TabIndex        =   20
         Top             =   3840
         Width           =   3495
      End
      Begin VB.Label Label12 
         Caption         =   "SUPERBAT -  CREATED BY TYLER ROBBINS EMAILS: TYLER@ROBBINS.TO      SUM1ELSE@OPTONLINE.NET              AIM SCREENNAME: SUM1E1SE"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74520
         TabIndex        =   19
         Top             =   7080
         Width           =   9375
      End
      Begin VB.Label Label11 
         Caption         =   $"Form2.frx":02F9
         Height          =   1095
         Left            =   -69960
         TabIndex        =   18
         Top             =   2880
         Width           =   3375
      End
      Begin VB.Label Label10 
         Caption         =   "Erase - Deletes Files                                                          Usage Erase filename"
         Height          =   1455
         Left            =   -69960
         TabIndex        =   16
         Top             =   2160
         Width           =   4215
      End
      Begin VB.Label Label9 
         Caption         =   $"Form2.frx":0399
         Height          =   1095
         Left            =   -69960
         TabIndex        =   14
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label Label8 
         Caption         =   "DOS Commands:"
         Height          =   495
         Left            =   -69960
         TabIndex        =   13
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label7 
         Caption         =   "ECHO - is to display a text of your choice. It can say anything"
         Height          =   1455
         Left            =   -74640
         TabIndex        =   11
         Top             =   5760
         Width           =   3615
      End
      Begin VB.Label Label6 
         Caption         =   $"Form2.frx":0428
         Height          =   1575
         Left            =   -74640
         TabIndex        =   10
         Top             =   4680
         Width           =   3615
      End
      Begin VB.Label Label5 
         Caption         =   $"Form2.frx":04DF
         Height          =   1215
         Left            =   -74640
         TabIndex        =   9
         Top             =   3480
         Width           =   3735
      End
      Begin VB.Label Label4 
         Caption         =   $"Form2.frx":05D7
         Height          =   1215
         Left            =   -74640
         TabIndex        =   8
         Top             =   2520
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   $"Form2.frx":067D
         Height          =   1095
         Left            =   -74640
         TabIndex        =   7
         Top             =   1440
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "Explanation Of Terms"
         Height          =   495
         Left            =   -74880
         TabIndex        =   6
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Insert Code"
         Height          =   375
         Left            =   7680
         TabIndex        =   5
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnusavproj 
         Caption         =   "Save Project"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuOpenProj 
         Caption         =   "Open Project"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuCompile 
         Caption         =   "Compile Batch File"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuRunmain 
      Caption         =   "Run"
      Begin VB.Menu mnuTest 
         Caption         =   "Test Batch File"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuContactCreator 
         Caption         =   "Contact Creator"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About SuperBatch"
         Shortcut        =   ^B
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sayit
Dim path
Private Sub Command1_Click()
Text1.Text = Text1.Text & "@Echo ON" & Chr(10)


End Sub

Private Sub Command2_Click()
Text1.Text = Text1.Text & "@Echo OFF" & Chr(10)

End Sub

Private Sub Command3_Click()
CommonDialog1.Filter = "DOS Batch File|*.bat|DOS Application|*.com|Executable|*.exe|"
CommonDialog1.FilterIndex = 3
CommonDialog1.ShowOpen
Text1.Text = Text1.Text & (CommonDialog1.Filename) & Chr(10)
Text1.SetFocus


End Sub

Private Sub Command4_Click()
Text1.Text = Text1.Text & "pause" & Chr(10)

End Sub



Private Sub Command5_Click()
Text1.Text = Text1.Text & "CLS" & Chr(10)

End Sub

Private Sub Command6_Click()
Text1.Text = Text1.Text & "Exit" & Chr(10)

End Sub

Private Sub Command7_Click()
Form3.Visible = True
End Sub

Private Sub Command8_Click()

sayit = InputBox("What Should Appear When This Is Run?", "Display Text")

If sayit = "" Then
    MsgBox ("No Text Entered")
 Else
Text1.Text = Text1.Text & "Echo " & sayit & Chr(10)
End If

End Sub

Private Sub Form_Load()
path = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)

If path = "" Then
End ' I put this here becuase it kept starting again when i clicked the x in the corner
Else
Kill path
End ' I put this here becuase it kept starting again when i clicked the x in the corner
End If
End Sub


Private Sub mnuAbout_Click()
MsgBox ("SuperBat VERSION " & App.Major & "." & App.Minor & "." & App.Revision)


End Sub

Private Sub mnuClose_Click()
End
End Sub

Private Sub mnuCompile_Click()
On Error GoTo errohandl:
CommonDialog2.Filter = "DOS Batch File|*.bat|"
CommonDialog2.FilterIndex = 1
CommonDialog2.ShowSave
Open CommonDialog2.Filename For Output As 1
Print #1, Text1.Text
Close #1
Form2.Caption = "Superbat - " + CommonDialog2.Filename
errohandl:


End Sub

Private Sub mnuContactCreator_Click()
MsgBox ("Email Tyler Robbins @  Tyler@Robbins.to")


End Sub

Private Sub mnuCopy_Click()
   Clipboard.Clear                     'Clear Clipboard for new data
    If Text1.SelText = "" Then        'See if the selected text is Blank
    Exit Sub                            'it is so exit the procedure
    Else                                'Otherwise
    Clipboard.SetText Text1.SelText   'Put the selected text into the clipboard
    End If

End Sub

Private Sub mnuCut_Click()
Clipboard.Clear                     'Clear the Clipboard for new data
    Clipboard.SetText Text1.SelText   'Put selected text into clipboard
    Text1.SelText = ""                'Erase Selected text from textbox
    mnuPaste.Enabled = True
End Sub

Private Sub mnuOpenProj_Click()




On Error GoTo ErrorHandler:

CommonDialog2.Filter = "Batch File|*.bat|SuperBat Project|*.sbat|"
CommonDialog2.FilterIndex = 2
CommonDialog2.ShowOpen
Filename = CommonDialog2.Filename
    OpenFile (Filename)
Form2.Caption = "SuperBAT - " + CommonDialog2.Filename

ErrorHandler:
    Exit Sub


Exit Sub

End Sub



Private Sub mnuPaste_Click()
Text1.SelText = Clipboard.GetText()
End Sub

Private Sub mnuPrint_Click()
Printer.NewPage
Printer.Print "" + Text1.Text
Printer.EndDoc

End Sub

Private Sub mnusavproj_Click()
On Error GoTo errohandl:
CommonDialog2.Filter = "SuperBat Project|*.sbat|"
CommonDialog2.FilterIndex = 1
CommonDialog2.ShowSave
Open CommonDialog2.Filename For Output As 1
Print #1, Text1.Text
Close #1
Form2.Caption = "SuperBat - " + CommonDialog2.Filename
errohandl:


End Sub

Private Sub OpenFile(Filename As String)
    Dim F As Integer
    F = FreeFile
    Open Filename For Input As F
    Text1.Text = Input$(LOF(F), F)
    Close F
    Exit Sub
    Exit Sub
End Sub

Private Sub mnuTest_Click()
retval = MsgBox("WARNING: THE PROGRAM YOU JUST WROTE MAY HAVE CONTENTS YOU PUT IN THAT CAN HARM YOUR COMPUTER. ARE YOU SURE YOU STILL WANT TO RUN?", vbYesNo)
path = "c:\windows\temp\supaBat.bat"




Select Case retval
Case vbYes
Open path For Output As 2
Print #2, Text1.Text
Close #2
retval = Shell(path, vbNormalFocus)
MsgBox ("Press OK When done testing")


Case vbNo

End Select

End Sub



Private Sub Timer1_Timer()
If Clipboard.GetText = "" Then
mnuPaste.Enabled = False
Else
mnuPaste.Enabled = True
End If



End Sub
