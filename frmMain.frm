VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Usage of Shell32.dll"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   113
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Label LblCmd 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   11
      Left            =   3120
      TabIndex        =   11
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label LblCmd 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Shutdown"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   10
      Left            =   1680
      TabIndex        =   10
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label LblCmd 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Open File/Folder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label LblCmd 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Unminimize All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label LblCmd 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Minimize All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   1680
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label LblCmd 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Windows Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label LblCmd 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Find Files"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label LblCmd 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Find Computer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   1680
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label LblCmd 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Explore"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label LblCmd 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Control Panel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label LblCmd 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Browse FF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label LblCmd 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Run Dialog"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Tmp1 As String
Dim Tmp2 As String

Private SH32 As New Shell32.Shell

Private Sub Form_Load()
'set backcolor
Me.BackColor = &H8000000D
'Center Form, the easy way
Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = Screen.Height / 2 - Me.Height / 2
End Sub

Public Function Display_Run_Dialog()
SH32.FileRun
End Function

Public Function Display_Browse_for_Folder(BrwTitle As String, Optional RootFolder As String)
'Detect if Rootfolder was given or not
If Not RootFolder = vbNullString Then            'if we have the rootfolder variable...
SH32.BrowseForFolder Me.hWnd, BrwTitle, 1, RootFolder
Else:                                            'if we don't have the rootfolder variable
SH32.BrowseForFolder Me.hWnd, BrwTitle, 1
End If
End Function

Public Function Display_ControlPanel_Item(ItemFileName As String)
'Opens the CPL file of your choice.
'They're all located in the Windows\System Dir
SH32.ControlPanelItem ItemFileName
End Function

Public Function ExploreL(LocationToOpen As String)
'like shell "Explorer.exe Location",1
SH32.Explore LocationToOpen
End Function

Public Function Display_FindComputer_Dialog()
SH32.FindComputer
End Function

Public Function Display_FindFiles_Dialog()
SH32.FindFiles
End Function

Public Function Open_WindowsHelp()
SH32.Help
End Function

Public Function Minimize_All()
SH32.MinimizeAll
End Function

Public Function Un_Minimize_All()
SH32.UndoMinimizeALL
End Function

Public Function Open_File_or_Folder(FileOrFolder As String)
'just like the ShellExecute Api
SH32.Open FileOrFolder
End Function

Public Function ShutDown_Windows()
'Be carefull with this command!
SH32.ShutdownWindows
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
For i = 0 To 11
'make sure every label's backstyle is set to 0 and forecolor is black
If Not LblCmd(i).BackStyle = 0 Then LblCmd(i).BackStyle = 0
If Not LblCmd(i).ForeColor = RGB(0, 0, 0) Then LblCmd(i).ForeColor = RGB(0, 0, 0)
Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Proper way to exit an application
Me.Hide
Set frmMain = Nothing
End
End Sub

Private Sub LblCmd_Click(Index As Integer)
'Look which Labels was pressed
Select Case Index
 Case "0"
  Display_Run_Dialog
 Case "1"
  Tmp1 = InputBox("Enter the text which should be display above the Browse Window", "Window Description", "Select your Folder")
  Tmp2 = InputBox("If you want to define a Root folder enter the folder or just enter nothing!", "Root Folder?")
  Display_Browse_for_Folder Tmp1, Tmp2
 Case "2"
  frmSelect.Show
 Case "3"
  Tmp1 = InputBox("Enter the Folder you want to open!", "Enter Folder")
  ExploreL Tmp1
 Case "4"
  Display_FindComputer_Dialog
 Case "5"
  Display_FindFiles_Dialog
 Case "6"
  Open_WindowsHelp
 Case "7"
  Minimize_All
 Case "8"
  Un_Minimize_All
 Case "9"
  Tmp1 = InputBox("Enter the file or folder you want to open!", "Enter File or Folder")
  Open_File_or_Folder Tmp1
 Case "10"
  If MsgBox("Continue? Your Windows will be shut down if you press yes or press no to cancel!", vbQuestion + vbYesNo) = vbYes Then
  ShutDown_Windows
  Else:
  End If
 Case "11"
  Unload Me
End Select
End Sub


Private Sub LblCmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not LblCmd(Index).BackStyle = 1 Then LblCmd(Index).BackStyle = 1
If Not LblCmd(Index).ForeColor = RGB(255, 255, 255) Then LblCmd(Index).ForeColor = RGB(255, 255, 255)
End Sub
