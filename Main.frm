VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "The Boogie Searcher"
   ClientHeight    =   10740
   ClientLeft      =   8310
   ClientTop       =   420
   ClientWidth     =   6990
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleMode       =   0  'User
   ScaleWidth      =   6990
   Begin VB.CheckBox Check2 
      BackColor       =   &H80000003&
      Caption         =   "Enable Delete"
      Height          =   285
      Left            =   2850
      TabIndex        =   22
      Top             =   1470
      Width           =   1305
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H000000FF&
      Caption         =   "Delete File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1500
      MaskColor       =   &H80000003&
      TabIndex        =   21
      Top             =   1560
      Width           =   1245
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Request List"
      Height          =   465
      Left            =   1500
      TabIndex        =   20
      Top             =   1560
      Width           =   1245
   End
   Begin VB.TextBox Text7 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4320
      TabIndex        =   19
      Text            =   "Folder Search:"
      Top             =   1470
      Width           =   1035
   End
   Begin VB.TextBox Text6 
      Height          =   315
      Left            =   5400
      TabIndex        =   18
      Top             =   1440
      Width           =   1515
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Catalogue"
      Height          =   450
      Left            =   1500
      TabIndex        =   17
      Top             =   960
      Width           =   1245
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Open Playlist"
      Height          =   450
      Left            =   2940
      TabIndex        =   16
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Make Playlist"
      Height          =   450
      Left            =   4320
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Set Options"
      Height          =   450
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3420
      TabIndex        =   13
      Text            =   "File search:"
      Top             =   1860
      Width           =   915
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   10365
      Width           =   6735
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      TabIndex        =   11
      Top             =   10185
      Width           =   6975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "View Folders"
      Height          =   450
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   8190
      Left            =   150
      TabIndex        =   9
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2475
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2385
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Karaoke"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Music Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4380
      TabIndex        =   0
      Top             =   1800
      Width           =   2535
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   8145
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   14367
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   16776960
      BackColor       =   0
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   12347
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   88
      EndProperty
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search in Subdirectories"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5700
      MaskColor       =   &H00FFFF80&
      TabIndex        =   1
      Top             =   960
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   285
      Left            =   2700
      TabIndex        =   4
      Top             =   750
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10815
      Left            =   0
      Picture         =   "Main.frx":164A
      ScaleHeight     =   10815
      ScaleWidth      =   6945
      TabIndex        =   8
      Top             =   -120
      Width           =   6945
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Const vbDot = 46
Private Const MAXDWORD As Long = &HFFFFFFFF
Private Const MAX_PATH As Long = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Type FILE_PARAMS
   bRecurse As Boolean
   sFileRoot As String
   sFileNameExt As String
   sResult As String
   sMatches As String
   Count As Long
End Type

Private Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long
   
Private Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function FindNextFile Lib "kernel32" _
   Alias "FindNextFileA" _
  (ByVal hFindFile As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Check2_Click()
Dim Response As String
If Check2.Value = 1 Then
Command8.Visible = True
Response = MsgBox("Use with extreme caution! Files will be Permanently Deleted!!!", vbExclamation, "Warning!")
Else
Command8.Visible = False
End If
End Sub

Private Sub Command1_Click()
 Module1.LoadFormState frmOptions
   Dim FP As FILE_PARAMS
   Dim tstart As Single
   Dim tend As Single
   
   
   
   ListView1.Left = 100
   ListView1.Width = 6500 're-size listview
   ListView1.Width = Me.Width - 300
   Command2.Caption = "Show Folders"
   Dir1.Visible = False
  
   Text3.Text = ""
   ListView1.ListItems.Clear
   ListView1.Visible = False
   
   If Option1.Value = True Then
   Text1.Text = frmOptions.Text1.Text
   Text2.Text = "Exploring Music Directory"
   ListView1.ForeColor = &HFF00&
   Else
   Text1.Text = frmOptions.Text2.Text
   Text2.Text = "Exploring " & Text6.Text & " Karaoke Directory"
   ListView1.ForeColor = &HFFFF00
   End If
  
  'set up search params
   With FP
      .sFileRoot = Text1.Text & "\" & Text6.Text  'start path
      .sFileNameExt = "*" & Text5.Text & "*." & frmOptions.Combo1.Text    'file type of interest
      .bRecurse = Check1.Value = 1  '1 = recursive search
   End With
   
  '
   tstart = GetTickCount()
   Call SearchForFiles(FP)
   tend = GetTickCount()
   
   ListView1.Visible = True
   
  
   
      Text3.Text = "has found " & ListView1.ListItems.Count & " " & frmOptions.Combo1.Text & " files"
   
      Text5.Text = ""
      Text6.Text = ""
End Sub


Private Sub GetFileInformation(FP As FILE_PARAMS)

  'local working variables
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
      
  
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & FP.sFileNameExt
   
 
   hFile = FindFirstFile(sPath, WFD)
   
  
   If hFile <> INVALID_HANDLE_VALUE Then

      Do
         
        
         If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = _
                 FILE_ATTRIBUTE_DIRECTORY Then

           
            FP.Count = FP.Count + 1
            sTmp = TrimNull(WFD.cFileName)
            'List1.AddItem sRoot & sTmp - the subitem carries the
            'path & filename for dragging - the main list is filename only
            Dim lx As ListItem
            Set lx = ListView1.ListItems.Add(, , sTmp)
            lx.SubItems(1) = (sRoot & sTmp)

            
            



         End If
         
      Loop While FindNextFile(hFile, WFD)
      
      
     
      hFile = FindClose(hFile)
   
   End If

End Sub


Private Sub SearchForFiles(FP As FILE_PARAMS)

  
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
      
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & "*.*"
   
  
   hFile = FindFirstFile(sPath, WFD)
   
 
   If hFile <> INVALID_HANDLE_VALUE Then
   
     
      Call GetFileInformation(FP)

      Do
      
        
         If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            
           
            If FP.bRecurse Then
            
             
               If Asc(WFD.cFileName) <> vbDot Then
               
                
                  FP.sFileRoot = sRoot & TrimNull(WFD.cFileName)
                  Call SearchForFiles(FP)
                  
               End If
               
            End If
            
         End If
         
     
      Loop While FindNextFile(hFile, WFD)
      
     
      hFile = FindClose(hFile)
   
   End If
   
End Sub


Private Function QualifyPath(sPath As String) As String

  'assures that a passed path ends in a slash
   If Right$(sPath, 1) <> "\" Then
      QualifyPath = sPath & "\"
  
   Else
      QualifyPath = sPath
   End If
     
End Function


Private Function TrimNull(startstr As String) As String

  'returns the string up to the first
  'null, if present, or the passed string
   Dim pos As Integer
   
   pos = InStr(startstr, Chr$(0))
   
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
  
   TrimNull = startstr
  
End Function

Private Sub Command2_Click()


ListView1.Left = 2300

ListView1.Width = Me.Width - 2400
If Dir1.Visible = True Then
 Command1.Value = True
 Command2.Caption = "Show Folders"
ElseIf Dir1.Visible = False Then
 Dir1.Visible = True
  Command2.Caption = "Hide Folders"
End If
If Option1.Value = True Then
Dir1.path = frmOptions.Text1.Text
End If
If Option2.Value = True Then
Dir1.path = frmOptions.Text2.Text
End If
Dim FP As FILE_PARAMS
   Dim tstart As Single
   Dim tend As Single
   ListView1.ListItems.Clear
 With FP
      .sFileRoot = Text1.Text
      .sFileNameExt = "*" & Text5.Text & "*." & frmOptions.Combo1.Text
      .bRecurse = Check1.Value = 1
   End With
   
  
   tstart = GetTickCount()
   Call SearchForFiles(FP)
   tend = GetTickCount()
Text3.Text = "has found " & ListView1.ListItems.Count & " " & frmOptions.Combo1.Text & " files"
Text2.Text = Dir1.path
End Sub

Private Sub Command3_Click()
frmOptions.Visible = True
End Sub

Private Sub Command4_Click()
If Option2.Value = True Then
MsgBox "You cannot make a playlist from Karaoke Files!"
Command4.Caption = "Make Playlist"
Exit Sub
Else
If Form2.Visible = True Then
 Form2.Visible = False
  Command4.Caption = "Show ListMaker"
ElseIf Form2.Visible = False Then
 Form2.Visible = True
  Command4.Caption = "Hide ListMaker"
End If
End If
Form2.Text1.Text = ""
Form2.List1.Clear

End Sub



Private Sub Command5_Click()
If Form3.Visible = True Then
 Form3.Visible = False
 Command5.Caption = "Show Playlist"
ElseIf Form3.Visible = False Then
 Form3.Visible = True
 Command5.Caption = "Hide Playlist"
End If
End Sub

Private Sub Command6_Click()
Module1.LoadFormState frmOptions
If frmMain.Visible = True Then
 frmMain.Visible = False
 Command6.FontBold = False
 Command6.Caption = "Show Catalogue"
ElseIf frmMain.Visible = False Then
 frmMain.Visible = True
 Command6.FontBold = False
 Command6.Caption = "Hide Catalogue"
End If

End Sub

Private Sub Command7_Click()
If Form6.Visible = True Then
 Form6.Visible = False
 Command7.FontBold = False
 Command7.Caption = "Show Songlist"
ElseIf Form6.Visible = False Then
 Form6.Visible = True
 Command7.FontBold = False
 Command7.Caption = "Hide Songlist"
End If
End Sub

Private Sub Command8_Click()
'Dim response As String
'If ListView1.ListItems.SelCount = 0 Then
'response = MsgBox("you have not selected a file!", vbInformation)
'Exit Sub
'Else
Kill ListView1.SelectedItem.SubItems(1)
Command1.Value = True
'End If
End Sub

'Private Sub Command6_Click()
'
'Module1.LoadFormState frmOptions
''MsgBox "Loading information from Excel file - Please Wait!"
'Command1.Value = True
'frmSplash.Visible = True
'ListView1.ListItems.Clear
'Command6.FontBold = True
'Command6.Caption = "Loading...."
'
'If Form4.Visible = True Then
' Form4.Visible = False
' Command6.FontBold = False
' Command6.Caption = "Show Catalogue"
'ElseIf Form4.Visible = False Then
' Form4.Visible = True
' Command6.FontBold = False
' Command6.Caption = "Hide Catalogue"
'End If
'
'frmSplash.Visible = False
'Command1.Value = True
'Command6.FontBold = False
'End Sub

Private Sub Dir1_Change()

Text2.Text = Dir1.path
Dim FP As FILE_PARAMS
   Dim tstart As Single
   Dim tend As Single
   
  
   Text3.Text = ""
   ListView1.ListItems.Clear
   ListView1.Visible = False
   
   If Option1.Value = True Then
   Text1.Text = Text2.Text
   ListView1.ForeColor = &HFF00&
   Else
   Text1.Text = Text2.Text
   ListView1.ForeColor = &HFFFF00
   End If
  
  
   With FP
      .sFileRoot = Text1.Text       'start path
      .sFileNameExt = "*" & Text5.Text & "*." & frmOptions.Combo1.Text  'file type of interest
      .bRecurse = Check1.Value = 1  '1 = recursive search
   End With
   
  
   tstart = GetTickCount()
   Call SearchForFiles(FP)
   tend = GetTickCount()
   
   ListView1.Visible = True
   
  
      Text3.Text = "has found " & ListView1.ListItems.Count & " " & frmOptions.Combo1.Text & " files"
   
      Text5.Text = ""
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Resize()
On Error GoTo Error_Trap

If Me.WindowState <> vbMinimized Then
   
   ListView1.Width = Me.Width - 300
   Option1.Left = Me.Width - Option1.Width - 350
   Option2.Left = Me.Width - Option2.Width - 350
   ListView1.Height = Me.Height - ListView1.Top - 650
   Picture1.Width = Me.Width
   Picture1.Height = Me.Height
   Text5.Left = Me.Width - Text5.Width - 200
   Text6.Left = Me.Width - Text6.Width - 200
   'Command1.Left = Me.Width - Command1.Width - 300
   'Command4.Left = Me.Width - Command2.Width - 1750
   Dir1.Height = Me.Height - Dir1.Top - 800
   Frame1.Width = Me.Width
   Frame1.Top = Me.Height - Frame1.Height - 400
   Text2.Top = Me.Height - Text2.Height - 500
   
   
End If

Exit Sub
Error_Trap: 'Let it error out and continue '''Not very good but it works
 Err.Clear
Exit Sub
End Sub
Private Sub Form_Load()
Module1.LoadFormState frmOptions
Text5.Text = ""
Text3.Text = ""
Check1.Value = 1
Check2.Value = 0
Option1.Value = True
Command1.Value = True
Command8.Visible = False
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


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   Dim frmWork As Form
   For Each frmWork In Forms
      Unload frmWork
      Set frmWork = Nothing
   Next frmWork

End Sub



