VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Song Queue"
   ClientHeight    =   4980
   ClientLeft      =   90
   ClientTop       =   6195
   ClientWidth     =   8085
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Song Queue.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8085
   Begin VB.CommandButton Command3 
      Caption         =   "Clear All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2700
      TabIndex        =   9
      Top             =   780
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6240
      TabIndex        =   8
      Top             =   60
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Open List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4020
      TabIndex        =   7
      Top             =   780
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000014&
      Caption         =   "Delete songs from view after playing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      TabIndex        =   6
      Top             =   690
      Width           =   2205
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4680
      TabIndex        =   5
      Top             =   480
      Width           =   3225
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear Selected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5340
      TabIndex        =   4
      Top             =   780
      Width           =   1245
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4860
      TabIndex        =   3
      Text            =   "Select Song List:"
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6660
      TabIndex        =   1
      Top             =   780
      Width           =   1245
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000012&
      ForeColor       =   &H0000FFFF&
      Height          =   3570
      Left            =   60
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   1245
      Width           =   7935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6870
      Left            =   -60
      Picture         =   "Song Queue.frx":164A
      ScaleHeight     =   6870
      ScaleWidth      =   8130
      TabIndex        =   2
      Top             =   -30
      Width           =   8130
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()

Dim Response As Integer
Dim i As Variant

If Form6.Combo1.Text = "" Then
Response = MsgBox("Please select a list filename", , "Name required")
Form6.Combo1.SetFocus
Exit Sub
Else
If Form6.List1.ListCount = 0 Then
Response = MsgBox("There are no songs in the Song Queue!", , "Empty List")
Exit Sub
Else
If dir(App.path & "\" & Form6.Combo1.Text & ".txt") = "" Then
GoTo 10
Else
  Response = MsgBox("A file of that name already exists, Overwrite?", vbInformation + vbYesNo, "Warning!")
If Response = vbYes Then
GoTo 10

Else
If Response = vbNo Then
Exit Sub
10
   Open (App.path & "\" & Form6.Combo1.Text & ".txt") For Output As 1
For i = 0 To Form6.List1.ListCount - 1
    Print #1, Form6.List1.List(i)
   Next
Close #1

End If
End If
End If
End If
End If
End Sub




Private Sub Command2_Click()
If List1.ListCount = 0 Then
MsgBox "There are no songs in the list!"
Else
If List1.SelCount = 0 Then
MsgBox "You have not selected anything"
Else
List1.RemoveItem (List1.ListIndex)
End If
End If
Text3.Text = "There are " & List1.ListCount & " songs in this queue"
End Sub







Private Sub Command3_Click()
If List1.ListCount = 0 Then
MsgBox "There are no songs in the list!"
Else
List1.Clear
End If
Text3.Text = "There are " & List1.ListCount & " songs in this queue"
End Sub

Private Sub Command5_Click()
On Error GoTo Error_Trap

Dim dir
Dim strString As String
If Combo1.Text = "" Then
MsgBox ("Please select a list filename")
Exit Sub
Else
End If
List1.Clear
dir = App.path & "\" & Combo1.Text & ".txt"
Text3.Text = "There are " & List1.ListCount & " songs in this queue"

Open dir For Input As #1
While Not EOF(1)
Line Input #1, strString
List1.AddItem strString
Wend
Close #1
Text3.Text = "There are " & List1.ListCount & " songs in this queue"

Error_Trap:
 Err.Clear
Exit Sub
End Sub

Private Sub Form_Load()
Combo1.Clear
Combo1.AddItem "Regulars"
Combo1.AddItem "Popular"
Combo1.AddItem "Other"
Combo1.AddItem "Blank"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Command7.Caption = "Song Requests"
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Error_Trap
  List1.AddItem Data.Files.Item(1)
  Text3.Text = "There are " & List1.ListCount & " songs in this queue"
  Dim i As Integer, j As Integer
For i = (List1.ListCount - 1) To 0 Step -1 'Reverse Loop through the listbox
For j = (List1.ListCount - 1) To 0 Step -1

If i <> j Then ' Don't compare item to itself
If InStr(List1.List(i), List1.List(j)) <> 0 Then 'If Already exists ..
List1.RemoveItem i ' .. Remove it
End If
End If

Next j
Next i
Text3.Text = "There are " & List1.ListCount & " songs in this Queue"
Exit Sub
Error_Trap:
 Err.Clear
Exit Sub
End Sub
Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If List1.ListCount = 0 Then
  Exit Sub
  Else
  End If
If Button = vbLeftButton Then List1.OLEDrag
  
End Sub


Private Sub Form_Resize()
On Error GoTo Error_Trap

If Me.WindowState <> vbMinimized Then
   
   List1.Width = Me.Width - 300
   List1.Height = Me.Height - List1.Top - 600
   
   Combo1.Left = Me.Width - Combo1.Width - 300
   
   
End If

Exit Sub
Error_Trap:
 Err.Clear
Exit Sub
End Sub

Private Sub List1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error GoTo Error_Trap
AllowedEffects = vbDropEffectCopy
Data.Clear
Data.Files.Add List1.Text
Data.SetData , vbCFFiles
If Check1.Value = 1 Then
List1.RemoveItem (List1.ListIndex)
Else
End If
Exit Sub
Error_Trap:
 Err.Clear
Exit Sub
End Sub





