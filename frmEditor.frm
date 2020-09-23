VERSION 5.00
Begin VB.Form frmEditor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Map Editor"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   353
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   448
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Color() As Byte
Private Sub Form_Load()
ReDim Color(frmEditor.ScaleWidth * frmEditor.ScaleHeight) As Byte
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim intPos As Long
intPos = (Y * frmEditor.ScaleWidth + X)
If Button = 1 Then
    Color(intPos) = 0
Else
    Color(intPos) = 1
End If
Me.Line (X, Y)-(X, Y)
End Sub

Private Sub Form_Resize()
ReDim Color(frmEditor.ScaleWidth * frmEditor.ScaleHeight) As Byte
End Sub
