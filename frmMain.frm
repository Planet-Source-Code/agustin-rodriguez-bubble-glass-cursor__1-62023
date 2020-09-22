VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   147
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      Caption         =   "You shouldn't make the form too big. If the form is too big, the form will be invisible"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'=              Agustin Rodriguez             =
'=         virtual_guitar_1@hotmail.com       =
'=    http://www.foreverbahia.com.br/agustin  =
'=  http://www.geocities.com/virtual_quality  =
'==============================================

'Based in code written by Apeiron
'http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=55313&lngWId=1
'Press SHIFT + ESC to exit

Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long  ' Declare API

Private Type POINTAPI
     x As Long
     y As Long
 End Type
 
Private PNG As LayeredWindow

Private Sub Form_Activate()
Dim pos As POINTAPI, w As Long, h As Long

w = 1000
h = 1000

Do
    GetCursorPos pos
    Move pos.x * Screen.TwipsPerPixelX - w, pos.y * Screen.TwipsPerPixelY - h
    DoEvents
    If GetAsyncKeyState(27) And GetAsyncKeyState(16) Then
        Unload Me
        End
    End If
Loop

End Sub

Private Sub Form_Load()

Set PNG = New LayeredWindow

PNG.MakeTrans App.Path & "\Test.png", Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

PNG.UnloadPNGForm

End Sub
