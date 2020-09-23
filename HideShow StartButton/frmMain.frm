VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hide/Show the Start Menu Button"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Execute"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Show Start Button"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Hide Start Button"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShowWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    
Private Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
    
Private Declare Function FindWindowEx Lib "user32" _
    Alias "FindWindowExA" (ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long

Sub StartButton(blnValue As Boolean)
    Dim lngHandle As Long
    Dim lngStartButton As Long

    lngHandle = FindWindow("Shell_TrayWnd", "")
    lngStartButton = FindWindowEx(lngHandle, 0, "Button", vbNullString)

    If blnValue Then
        ShowWindow lngStartButton, 5
    Else
        ShowWindow lngStartButton, 0
    End If
End Sub



Private Sub Command1_Click()
    StartButton IIf(Option2.Value, True, False)
End Sub

