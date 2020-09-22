VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Boogie Playlist Editor"
   ClientHeight    =   5910
   ClientLeft      =   1110
   ClientTop       =   420
   ClientWidth     =   7080
   Icon            =   "Listeditor.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   7080
   Begin VB.CommandButton Command4 
      Caption         =   "Rem Duplicates"
      Height          =   375
      Left            =   4350
      TabIndex        =   8
      Top             =   750
      Width           =   1245
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2970
      TabIndex        =   7
      Top             =   450
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4470
      TabIndex        =   6
      Top             =   60
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear List"
      Height          =   375
      Left            =   150
      TabIndex        =   5
      Top             =   750
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear Selected"
      Height          =   375
      Left            =   3030
      TabIndex        =   4
      Top             =   750
      Width           =   1245
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3450
      TabIndex        =   3
      Text            =   "Enter Name:"
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Playlist"
      Height          =   375
      Left            =   5700
      TabIndex        =   1
      Top             =   750
      Width           =   1245
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   4545
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   1185
      Width           =   6855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6870
      Left            =   -240
      Picture         =   "Listeditor.frx":164A
      ScaleHeight     =   6870
      ScaleWidth      =   7230
      TabIndex        =   2
      Top             =   -60
      Width           =   7230
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim response As Integer
If Form5.Text1.Text = "" Then
response = MsgBox("Please enter a Name for the Playlist", , "Name required")
Form5.Text1.SetFocus
GoTo 10
Else
If Form5.List1.ListCount = 0 Then
response = MsgBox("There are no songs in the playlist!", , "Empty List")
GoTo 10
Else
If dir(frmOptions.Text3.Text & "\" & Form5.Text1.Text) <> "" Then
 response = MsgBox("You are about to overwrite the file " & Text1.Text & " - Are you Sure?", vbYesNo, "Overwrite file")
   If response = vbNo Then
   Exit Sub
   ElseIf response = vbYes Then
Open frmOptions.Text3.Text & "\" & Form5.Text1.Text For Output As 1
For i = 0 To List1.ListCount - 1
    Print #1, List1.List(i)
   Next
Close #1

Form5.Visible = False
10
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
Text3.Text = "There are " & List1.ListCount & " songs in this playlist"
End Sub

Private Sub Command3_Click()
If List1.ListCount = 0 Then
MsgBox "There are no songs in the list!"
Else
List1.Clear
End If
Text3.Text = "There are " & List1.ListCount & " songs in this playlist"
End Sub





Private Sub Command4_Click()
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
Text3.Text = "There are " & List1.ListCount & " songs in this playlist"
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    List1.AddItem Data.Files.Item(1)
    Text3.Text = "There are " & List1.ListCount & " songs in this playlist"
End Sub

Private Sub Form_Resize()
On Error GoTo Error_Trap

If Me.WindowState <> vbMinimized Then
   
   List1.Width = Me.Width - 300
   List1.Height = Me.Height - List1.Top - 600
   Picture1.Width = Me.Width
   Picture1.Height = Me.Height
   Text1.Left = Me.Width - Text1.Width - 350
   Command1.Left = Me.Width - Command1.Width - 300
   Text2.Left = Me.Width - Text2.Width - 2800
End If

Exit Sub
Error_Trap:
 Err.Clear
Exit Sub
End Sub
