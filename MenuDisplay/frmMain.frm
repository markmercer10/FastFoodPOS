VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Control Center"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton butnDisplay 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton butnPanes 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Modify Display Panes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton butnEdit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Edit Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DisplaysShown As Boolean
Dim D1 As New frmDisplay
Dim D2 As New frmDisplay

Private Sub butnDisplay_Click()
    If Not DisplaysShown Then
        Dim q As ADODB.Recordset
        Set q = db.Execute("SELECT * FROM displays ORDER BY id ASC")
        With q
            If Not (.EOF And .BOF) Then
                .MoveFirst
                Load D1
                D1.Width = !resolution_x * Screen.TwipsPerPixelX
                D1.Height = !resolution_y * Screen.TwipsPerPixelY
                D1.DrawPanel !id, !layout
                D1.Show
                
                .MoveNext
                Load D2
                D2.Width = !resolution_x * Screen.TwipsPerPixelX
                D2.Height = !resolution_y * Screen.TwipsPerPixelY
                D2.DrawPanel !id, !layout
                D2.Show
            End If
        End With
        DisplaysShown = True
    Else
        Unload D1
        Unload D2
        DisplaysShown = False
    End If
End Sub

Private Sub butnEdit_Click()
    frmMenuEditor.Show 1
End Sub

Private Sub butnPanes_Click()
    frmDisplayPanels.Show 1
End Sub

Private Sub Form_Load()
    DisplaysShown = False
    ConnectDB
End Sub
