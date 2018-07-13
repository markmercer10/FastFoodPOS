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
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Blocks 
      Appearance      =   0  'Flat
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
    moveStartX = X
    moveStartY = Y
End Sub

Private Sub Blocks_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    moveEndX = X - moveStartX
    moveEndY = Y - moveStartY
    If Button = 1 Then
        Me.left = Me.left + moveEndX
        Me.Top = Me.Top + moveEndY
    End If
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
