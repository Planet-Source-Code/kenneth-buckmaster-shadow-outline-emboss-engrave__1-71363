VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7293
   ClientLeft      =   44
   ClientTop       =   440
   ClientWidth     =   9823
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   7.76
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7293
   ScaleWidth      =   9823
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'how shadow outline emboss engrave hollow text effects are created

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private Sub Form_Load()

With Me
.ScaleMode = 3
.AutoRedraw = True
.FontSize = 30
End With

shadow Me, 20, 0, "Shadow", vbRed, RGB(50, 50, 50)
outline Me, 20, 60, "Outline", vbRed, vbBlack
emboss Me, 20, 120, "Emboss", vbBlack
engrave Me, 20, 180, "Engrave", vbBlack
outline Me, 20, 240, "Hollow", Me.BackColor, vbRed 'hollow as outline but uses backcolor for fill

End Sub

Sub emboss(obj As Object, x As Long, y As Long, st As String, col As Long)
With obj
obj.ForeColor = col
TextOut obj.hDC, x + 1, y + 1, st, Len(st)
obj.ForeColor = obj.BackColor
TextOut obj.hDC, x, y, st, Len(st)
End With
End Sub

Sub engrave(obj As Object, x, y, st As String, col As Long)
With obj
obj.ForeColor = col
TextOut obj.hDC, x - 1, y - 1, st, Len(st)
obj.ForeColor = obj.BackColor
TextOut obj.hDC, x, y, st, Len(st)
End With
End Sub


Sub shadow(obj As Object, x As Long, y As Long, st As String, textcol As Long, shadowcol As Long)
With obj
lst = Len(st)
obj.ForeColor = shadowcol
TextOut obj.hDC, x + 2, y + 2, st, lst 'draw shadow with small offset
obj.ForeColor = textcol
TextOut obj.hDC, x, y, st, lst 'draw text
End With
End Sub

Sub outline(obj As Object, x As Long, y As Long, st As String, innercol As Long, outercol As Long)
lst = Len(st)
With obj
obj.ForeColor = outercol
TextOut .hDC, x - 1, y - 1, st, lst 'draw outline
TextOut .hDC, x - 1, y + 1, st, lst
TextOut .hDC, x + 1, y + 1, st, lst
TextOut .hDC, x + 1, y - 1, st, lst
.ForeColor = innercol
TextOut .hDC, x, y, st, lst 'draw text
End With
End Sub

