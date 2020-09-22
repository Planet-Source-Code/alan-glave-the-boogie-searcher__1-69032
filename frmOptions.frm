VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00FEE2E3&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Options"
   ClientHeight    =   3795
   ClientLeft      =   4215
   ClientTop       =   4965
   ClientWidth     =   7485
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   -60
      Picture         =   "frmOptions.frx":000C
      ScaleHeight     =   4785
      ScaleWidth      =   1185
      TabIndex        =   19
      Top             =   -1020
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "...."
      Height          =   255
      Left            =   6720
      TabIndex        =   17
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "...."
      Height          =   255
      Left            =   6720
      TabIndex        =   16
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "...."
      Height          =   255
      Left            =   6720
      TabIndex        =   12
      Top             =   480
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5250
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   2370
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2520
      TabIndex        =   10
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2520
      TabIndex        =   9
      Top             =   960
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   8
      Top             =   480
      Width           =   3975
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   6
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   5
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   3180
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4230
      TabIndex        =   0
      Top             =   3180
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FEE2E3&
      Caption         =   "Set directory paths"
      Height          =   1695
      Left            =   2400
      TabIndex        =   18
      Top             =   240
      Width           =   4935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FEE2E3&
      Caption         =   "Set File extension type"
      Height          =   795
      Left            =   2400
      TabIndex        =   20
      Top             =   2070
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FEE2E3&
      Caption         =   "Playlists"
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FEE2E3&
      Caption         =   "Karaoke"
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FEE2E3&
      Caption         =   "Music"
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Option Explicit
Dim SH As New Shell  'reference to shell32.dll class
Dim ShBFF As Folder  'Shell Browse For Folder



Private Sub Command1_Click()
On Error Resume Next
'set object
Set ShBFF = SH.BrowseForFolder(hWnd, "Please choose Music folder and click OK!", 1)
With ShBFF.Items.Item
   'get folder props
   Text1 = .path

End With
End Sub

Private Sub Command2_Click()
On Error Resume Next
'set object
Set ShBFF = SH.BrowseForFolder(hWnd, "Please choose Karaoke folder and click OK!", 1)
With ShBFF.Items.Item
   'get folder props
   Text2 = .path

End With
End Sub
Private Sub Command3_Click()
On Error Resume Next
'set object
Set ShBFF = SH.BrowseForFolder(hWnd, "Please choose Playlist folder and click OK!", 1)
With ShBFF.Items.Item
   'get folder props
   Text3 = .path

End With
End Sub
Private Sub cmdApply_Click()
  Module1.SaveFormState frmOptions
  Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub Form_Load()
Combo1.Clear
Combo1.AddItem "Mp3"
Combo1.AddItem "Cdg"
Combo1.AddItem "Zip"
Combo1.AddItem "Txt"
End Sub
