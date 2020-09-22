VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BackColor       =   &H80000009&
   Caption         =   "Playlist"
   ClientHeight    =   4905
   ClientLeft      =   3495
   ClientTop       =   6255
   ClientWidth     =   4665
   Icon            =   "Listviewer.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   4665
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC9529&
      Height          =   285
      Left            =   1950
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   675
      Width           =   2580
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3765
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Double click to Edit the playlist - Or drag onto the Player"
      Top             =   1050
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   6641
      View            =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList2"
      ForeColor       =   16776960
      BackColor       =   -2147483642
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   7673
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   97
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   88
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   -60
      Picture         =   "Listviewer.frx":164A
      ScaleHeight     =   5055
      ScaleWidth      =   4725
      TabIndex        =   2
      Top             =   0
      Width           =   4725
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Dim Playlistpath As String
Dim Filename As String
Dim lx As ListItem
Playlistpath = (frmOptions.Text3.Text & "\")
ListView1.ListItems.Clear

  Filename = dir(fixed(Playlistpath) & "*.m3u", vbArchive Or vbHidden Or vbReadOnly Or vbSystem)
  Do While Filename <> ""
  
            Set lx = ListView1.ListItems.Add(, , Filename)
            lx.SubItems(1) = (Playlistpath & Filename)
            Filename = dir()
            Loop

Text1.Text = "Found " & ListView1.ListItems.Count & " Playlists"
End Sub
Private Function fixed(ByVal path As String) As String
fixed = path & IIf(Right(path, 1) = "\", "", "\")
End Function

Private Sub Form_Unload(Cancel As Integer)
Form1.Command5.Caption = "Show Playlist"
End Sub




Private Sub ListView1_DblClick()
On Error GoTo Error_Trap
Dim response As Integer
response = MsgBox("Do you want to edit the selected playlist?", vbQuestion + vbYesNo, ListView1.SelectedItem.Text)
If response = vbNo Then
GoTo 10
ElseIf response = vbYes Then

End If

Form5.Visible = True
Form5.Text1.Text = ""
Form5.Text1.Text = ListView1.SelectedItem.Text


Dim dir
Dim strString As String
Form5.List1.Clear
dir = frmOptions.Text3.Text & "\" & Form5.Text1.Text

Open dir For Input As #1
While Not EOF(1)
Line Input #1, strString
Form5.List1.AddItem strString
Wend
Close #1
Form5.Text3.Text = "There are " & Form5.List1.ListCount & " songs in this playlist"
Form1.Command5 = True
10
Exit Sub
Error_Trap:
 Err.Clear
Exit Sub
End Sub



Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then ListView1.OLEDrag
End Sub

Private Sub ListView1_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
AllowedEffects = vbDropEffectCopy
Data.Clear
Data.Files.Add ListView1.SelectedItem.SubItems(1)
Data.SetData , vbCFFiles
End Sub
Private Sub Form_Resize()
On Error GoTo Error_Trap

If Me.WindowState <> vbMinimized Then
   
   ListView1.Width = Me.Width - 300
   ListView1.Height = Me.Height - List1.Top - 600
   Picture1.Width = Me.Width
   Picture1.Height = Me.Height
   Text1.Left = Me.Width - Text1.Width - 300
   
End If

Exit Sub
Error_Trap:
 Err.Clear
Exit Sub
End Sub

