VERSION 5.00
Begin VB.Form frmSelect 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Select which Control Panel Item"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frmSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Go"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.FileListBox List1 
      BackColor       =   &H8000000D&
      Height          =   2820
      Hidden          =   -1  'True
      Left            =   120
      System          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'Use the function with the selected file
frmMain.Display_ControlPanel_Item List1.Path & "\" & List1.FileName
Me.Hide
Set frmSelect = Nothing
End Sub

Private Sub Command2_Click()
Me.Hide
'Remove from Memory
Set frmSelect = Nothing
End Sub

Private Sub Form_Load()
'set backcolor
Me.BackColor = &H8000000D
List1.BackColor = &H8000000D
'Center Form, the easy way
Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = Screen.Height / 2 - Me.Height / 2
'Change Dir of the ListFile Object
'            Get the Windows System Dir!
'            (This works only for Windows 9x I think)
List1.Path = Environ$("WINDIR") & "\System"
'Define what Filetypes will be displayed
List1.Pattern = "*.CPL"
End Sub
