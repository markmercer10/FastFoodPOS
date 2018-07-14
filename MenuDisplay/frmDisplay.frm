VERSION 5.00
Begin VB.Form frmDisplay 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   14985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Blocks 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Index           =   0
      Left            =   1440
      ScaleHeight     =   3135
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label LayoutID 
      Caption         =   "0"
      Height          =   495
      Left            =   7920
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label DisplayID 
      Caption         =   "1"
      Height          =   495
      Left            =   7920
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   3375
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim moveStartX As Integer
Dim moveStartY As Integer
Dim moveEndX As Integer
Dim moveEndY As Integer

Private Sub Blocks_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Blocks_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Blocks_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseUp Button, Shift, X, Y
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moveStartX = X
    moveStartY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moveEndX = X - moveStartX
    moveEndY = Y - moveStartY
    If Button = 1 Then
        Me.left = Me.left + moveEndX
        Me.Top = Me.Top + moveEndY
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'SNAP!
    If Abs(Me.left) < 1000 Then
        Me.left = 0
    ElseIf Abs(Me.left + Me.Width - Screen.Width) < 1000 Then
        Me.left = Screen.Width - Me.Width
    End If
    
    If Abs(Me.Top) < 1000 Then
        Me.Top = 0
    ElseIf Abs(Me.Top + Me.Height - Screen.Height) < 1000 Then
        Me.Top = Screen.Height - Me.Height
    End If
End Sub

Public Sub DrawPanel(ByVal display As Byte, ByVal layout As Byte)
    Dim v() As Double
    Dim obj As PanelBlock
    Dim i As Long
    Let v = GetLayout(layout)
    BlockCount = UBound(v) + 1
    Set obj = New PanelBlock
    
    For i = 0 To BlockCount - 1
        If i > Blocks.Count - 1 Then Load Blocks(i)
        Blocks(i).left = Me.Width * v(i, 1)
        Blocks(i).Top = Me.Height * v(i, 2)
        Blocks(i).Width = Me.Width * v(i, 3)
        Blocks(i).Height = Me.Height * v(i, 4)
        
        obj.Load display, i
        obj.DrawBlock Blocks(i), 1
        Blocks(i).Visible = True
    Next i

End Sub

