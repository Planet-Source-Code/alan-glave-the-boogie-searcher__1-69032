VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "Catalogue Searcher"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   6195
   ClientWidth     =   8100
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "pman"
   ScaleHeight     =   4980
   ScaleWidth      =   8100
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   9
      Text            =   "Enter Search text here :"
      Top             =   1200
      Width           =   2145
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   2235
   End
   Begin VB.OptionButton optSearch 
      BackColor       =   &H80000005&
      Caption         =   "Song"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   5760
      TabIndex        =   7
      Top             =   120
      Width           =   795
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   2340
      TabIndex        =   4
      Top             =   1200
      Width           =   2595
   End
   Begin VB.OptionButton optSearch 
      BackColor       =   &H80000005&
      Caption         =   "Folder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   6930
      TabIndex        =   3
      Top             =   120
      Width           =   945
   End
   Begin VB.OptionButton optSearch 
      BackColor       =   &H80000005&
      Caption         =   "Artist"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   4470
      TabIndex        =   2
      Top             =   120
      Width           =   885
   End
   Begin MSComctlLib.ListView lvresults 
      Height          =   3285
      Left            =   150
      TabIndex        =   1
      Top             =   1560
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   5794
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483642
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6450
      TabIndex        =   0
      Top             =   1020
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6195
      Left            =   -30
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   6195
      ScaleWidth      =   8115
      TabIndex        =   6
      Top             =   0
      Width           =   8115
      Begin VB.Shape shpPct 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         Height          =   195
         Left            =   4920
         Top             =   690
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpPctBG 
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   4890
         Top             =   660
         Visible         =   0   'False
         Width           =   2925
      End
   End
   Begin VB.Label lblDis1Search 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search for :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2400
      TabIndex        =   5
      Top             =   960
      Width           =   960
   End
   Begin VB.Line l1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   480
      X2              =   1560
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line l1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   1080
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
  On Error GoTo myERR
  
  
  
  optSearch(0) = True
  Text2.Text = "Search Artist For:"
  

  
 
  'initialize the listview
  lvresults.View = lvwReport
  lvresults.FullRowSelect = True
  lvresults.LabelEdit = lvwManual
  lvresults.ColumnHeaders.Add , "Artist", "Artist", 2500
  lvresults.ColumnHeaders.Add , "Song", "Song", 3200
  lvresults.ColumnHeaders.Add , "Folder", "Folder", 1100
  lvresults.ColumnHeaders.Add , "Track", "Track", 700
  
  

  
  Exit Sub
myERR:
  MsgBox "Error in sub-routine: Form_Load"
  MsgBox "Error Description: " & Err.Description
  Exit Sub
   
End Sub

Private Sub lvresults_DblClick()
Form1.Option2 = True
Form1.Text6.Text = lvresults.SelectedItem.ListSubItems(2)
Form1.Command1 = True
Form1.Text2.Text = "Exploring " & lvresults.SelectedItem.ListSubItems(2) & " Folder"
Form1.Text6.Text = ""
End Sub

Private Sub optSearch_Click(Index As Integer)
If optSearch(0) = True Then
  Text2.Text = "Search Artist For:"
 
  Else
  If optSearch(1) = True Then
  Text2.Text = "Search Song For:"

  Else
  If optSearch(2) = True Then
  Text2.Text = "Search Folder For:"

  Else
  End If
  End If
  End If
 
End Sub

Private Sub cmdSearch_Click()


  
'  If Not FileExists(App.path & IIf(Right$(App.path, 1) = "\", "", "\") & "data.txt") Then
'    'see if our database exists
'    MsgBox "Data file missing", vbCritical, "Karaoke Catalogue"
'    txtSearch.SetFocus
'    Exit Sub
'  End If

  Dim TmpFile As Integer
  Dim TmpStr As String
  Dim tmpArr() As String
  Dim cntLines As Long
  Dim curOpt As Integer
  Dim tmpLI As ListItem
  lvresults.ListItems.Clear
  cntLines = 0
  Text1.Text = "Searching ........"

If txtSearch.Text = "" Then
optSearch(2) = True
Text2.Text = "Full Listing"

  Else

  End If
  cmdSearch.Enabled = False
  frmMain.MousePointer = vbHourglass
  lvresults.Visible = False
 
  For curOpt = lvresults.ListItems.Count To 1 Step -1
    'destroy any old items in our listview
    DoEvents
    lvresults.ListItems.Remove (curOpt)
  Next curOpt
  
  For curOpt = 0 To 2
    'find out which option button is currently set to true
    If optSearch(curOpt).Value = True Then Exit For
  Next curOpt
  
  shpPctBG.Visible = True
  shpPct.Width = 0
  shpPct.Visible = True
  
  TmpFile = FreeFile
  
  Open App.path & IIf(Right$(App.path, 1) = "\", "", "\") & "Data.txt" For Input As #TmpFile
  'open database

  While Not EOF(TmpFile)
    DoEvents
    
    'this is the progress bar code
    If (cntLines Mod 50) = 0 Then shpPct.Width = cntLines / 2
    cntLines = cntLines + 1
    
    'read the input
    Line Input #TmpFile, TmpStr
    'see if it contains what we're looking for
    If InStr(1, LCase(TmpStr), LCase(txtSearch.Text)) <> 0 Then
      'split it by commas, then see if it matches the right field.
      tmpArr = Split(TmpStr, ",")
      If InStr(1, LCase(tmpArr(curOpt)), LCase(txtSearch.Text)) <> 0 Then
        'add an item to our listview
        Set tmpLI = lvresults.ListItems.Add(, , tmpArr(0))
        tmpLI.ListSubItems.Add , , tmpArr(1)
        tmpLI.ListSubItems.Add , , tmpArr(2)
        tmpLI.ListSubItems.Add , , tmpArr(3)
      
        Set tmpLI = Nothing
      End If
    End If
     
  Wend
  
  Close #TmpFile
  
  shpPct.Visible = False
  shpPctBG.Visible = False
  lvresults.Visible = True

  
  cmdSearch.Enabled = True
  frmMain.MousePointer = vbNormal
  txtSearch.Text = ""
  Text1.Text = lvresults.ListItems.Count & " Songs Found"
  AltLVBackground lvresults, vbWhite, &HF7EBCE
  Exit Sub
myERR:
  MsgBox "Error in sub-routine: cmdSearch_Click"
  MsgBox "Error Description: " & Err.Description
  Exit Sub

End Sub

Private Sub Form_Activate()
  On Error GoTo myERR
  
  txtSearch.SetFocus
  
  Exit Sub
myERR:
  MsgBox "Error in sub-routine: Form_Activate"
  MsgBox "Error Description: " & Err.Description
  Exit Sub
End Sub



Private Sub Form_Resize()
On Error GoTo Error_Trap

If Me.WindowState <> vbMinimized Then
   
   lvresults.Width = Me.Width - 400
   lvresults.Height = Me.Height - lvresults.Top - 600
   Picture1.Width = Me.Width
   Picture1.Height = Me.Height

End If

Exit Sub
Error_Trap: 'Let it error out and continue '''Not very good but it works
 Err.Clear
Exit Sub
End Sub

Private Sub AltLVBackground(lv As ListView, _
    ByVal BackColorOne As OLE_COLOR, _
    ByVal BackColorTwo As OLE_COLOR)

Dim lH      As Long
Dim lSM     As Byte
Dim picAlt  As PictureBox
    With lv
        If .View = lvwReport And .ListItems.Count Then
            Set picAlt = Me.Controls.Add("VB.PictureBox", "picAlt")
            lSM = .Parent.ScaleMode
            .Parent.ScaleMode = vbTwips
            .PictureAlignment = lvwTile
            lH = .ListItems(1).Height
            With picAlt
                .BackColor = BackColorOne
                .AutoRedraw = True
                .Height = lH * 2
                .BorderStyle = 0
                .Width = 10 * Screen.TwipsPerPixelX
                picAlt.Line (0, lH)-(.ScaleWidth, lH * 2), BackColorTwo, BF
                Set lv.Picture = .Image
            End With
            Set picAlt = Nothing
            Me.Controls.Remove "picAlt"
            lv.Parent.ScaleMode = lSM
        End If
    End With
End Sub



Private Sub Form_Unload(Cancel As Integer)
Form1.Command6.Caption = "Catalogue"
End Sub

Private Sub lvresults_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lvresults.Sorted And _
        ColumnHeader.Index - 1 = lvresults.SortKey Then
        lvresults.SortOrder = 1 - lvresults.SortOrder
    Else
        lvresults.SortOrder = lvwAscending
        lvresults.SortKey = ColumnHeader.Index - 1
    End If
    lvresults.Sorted = True
End Sub


