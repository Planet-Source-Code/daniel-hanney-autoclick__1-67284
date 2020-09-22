VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoClick"
   ClientHeight    =   1995
   ClientLeft      =   1125
   ClientTop       =   1500
   ClientWidth     =   3285
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1995
   ScaleWidth      =   3285
   Visible         =   0   'False
   Begin VB.Timer tmrMain 
      Interval        =   400
      Left            =   720
      Top             =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Const MOUSEEVENTF_LEFTDOWN = &H2

Dim Click As Integer

Private Sub Form_Load()
    If App.PrevInstance Then Unload Me
    App.TaskVisible = False
    Click = 1
End Sub

Private Sub tmrMain_Timer()
    Dim P As POINTAPI, L As Long, H As Long, A$, ClassName As String
    
    L = GetCursorPos(P)
    H = WindowFromPoint(P.X, P.Y)
    A$ = Space$(128)
    L = GetClassName(H, A$, 128)
    A$ = Left$(A$, L)
    
    ClassName = A$
    
    If Click = 1 Then
        If ClassName = "Button" Then mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        Click = 0
    End If
    
    If Click = 0 Then
        If ClassName = "Button" Then
            Click = 0
        Else
            Click = 1
        End If
    End If
End Sub
