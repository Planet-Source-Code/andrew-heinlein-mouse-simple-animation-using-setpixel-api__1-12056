VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "3d"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Sub PaintSquare(PicHDC As Long, SquareSize As Integer, ByVal EX As Long, ByVal WHY As Long, iColor As Long)
    Dim i As Integer, n As Integer
    For n = 1 To SquareSize
        For i = 1 To SquareSize
            If SetPixel(PicHDC, EX + i, WHY, iColor) = -1 Then
                Exit Sub
            End If
        Next i
        WHY = WHY + 1
    Next n
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static Xx As Long, Yy As Long, sSize As Integer, dDown As Boolean
    
    Picture1.Cls
    
    X = X / 15
    Y = Y / 15
    
    If X > Xx Then Xx = Xx + 2 Else: Xx = Xx - 2
    If Y > Yy Then Yy = Yy + 2 Else: Yy = Yy - 2
    
    If dDown = False Then
        sSize = sSize + 1
        If sSize > 15 Then dDown = True
    Else
        sSize = sSize - 1
        If sSize <= 5 Then dDown = False
    End If
    
    If Check1.Value <> 1 Then
        PaintSquare Picture1.hdc, sSize, CLng(Xx), CLng(Yy), vbRed
    Else
        PaintSquare Picture1.hdc, sSize, CLng(Xx + 4), CLng(Yy + 4), vbBlack
        PaintSquare Picture1.hdc, sSize, CLng(Xx), CLng(Yy), vbRed
    End If
    
End Sub


