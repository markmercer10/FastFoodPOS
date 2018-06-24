VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenuEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Editor"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   12825
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Ghost 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   11175
   End
   Begin VB.PictureBox HR 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   120
      ScaleHeight     =   45
      ScaleWidth      =   10935
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   10935
   End
   Begin VB.TextBox CellEdit 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin MSComctlLib.ListView MenuList 
      Height          =   8655
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   15266
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Small"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Medium"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Large"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1695
      Left            =   11280
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMenuEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Menu Editor
Dim indexDrag As Long
Dim columnClicked As Byte
Dim ItemHeight As Long
Dim dragging As Boolean

Private Sub CellEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WriteCell indexDrag, columnClicked, CellEdit.Text
        CellEdit.Visible = False
    End If
End Sub

Private Sub CellEdit_LostFocus()
    CellEdit.Visible = False
End Sub

Private Sub Command1_Click()
    Dim i As Long
    Dim li As ListItem
    For i = 1 To 20
        If (i - 1) Mod 5 = 0 Then
            Set li = MenuList.ListItems.Add(, , "Heading")
            li.Tag = "Heading"
        End If
        Set li = MenuList.ListItems.Add(, , "--Menu Item " & i)
        li.SubItems(3) = i
    Next i
    Set li = Nothing
End Sub

Private Sub Form_Load()
    Ghost.FontName = MenuList.Font.Name
    Ghost.FontSize = MenuList.Font.Size
End Sub

Private Sub HR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MenuList_MouseUp Button, Shift, X, HR.Top + 100
    CellEdit = HR.Top
End Sub

Private Sub MenuList_DblClick()
    OpenCellEditor indexDrag, columnClicked
End Sub

Private Sub MenuList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not MenuList.SelectedItem Is Nothing Then
        indexDrag = GetClickedIndex(Y)
        CellEdit = MenuList.ListItems(indexDrag).Tag
        If MenuList.ListItems(indexDrag).Tag <> "Heading" Then dragging = True
    End If
End Sub
 
Private Sub MenuList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 And dragging Then
        HR.Top = MenuList.ListItems(GetClickedIndex(Y)).Top
        Ghost.Text = MenuList.ListItems(indexDrag).Text
        Ghost.Top = Y + Ghost.Height / 2
        HR.Visible = True
        Ghost.Visible = True
    End If
End Sub

Private Sub MenuList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If dragging Then
        columnClicked = GetClickedColumn(X)
        If Not MenuList.SelectedItem Is Nothing Then
            Dim indexDragTo As Long
            indexDragTo = GetClickedIndex(Y)
            If Not indexDragTo = indexDrag Then MoveItem indexDrag, indexDragTo
        End If
        dragging = False
    End If
    HR.Visible = False
    Ghost.Visible = False
End Sub

Private Function MoveItem(ByVal Item1 As Long, ByVal Item2 As Long)
    If Item1 < 1 Or Item1 > MenuList.ListItems.Count Or Item2 < 1 Or Item2 > MenuList.ListItems.Count Then Exit Function
    If Item2 = 1 Then Item2 = 2
    If Item2 > Item1 Then Item2 = Item2 - 1
    Dim oli As ListItem
    Dim nli As ListItem
    Dim si As ListSubItem
    
    Set oli = MenuList.ListItems(Item1)
    MenuList.ListItems.Remove (Item1)
    Set nli = MenuList.ListItems.Add(Item2, , "")
    nli = oli
    For Each si In oli.ListSubItems
        nli.SubItems(si.Index) = ""
        nli.ListSubItems(si.Index) = si
    Next si
    Set oli = Nothing
    Set nli = Nothing
    Set si = Nothing
End Function

Private Function GetClickedIndex(ByVal Y As Long) As Long
    Dim HeaderExtraHeight As Long
    ItemHeight = MenuList.ListItems.Item(1).Height
    HeaderExtraHeight = ItemHeight * 0.15
    GetClickedIndex = Int((MenuList.GetFirstVisible.Index - 1) + (Y - HeaderExtraHeight) / ItemHeight)
    If GetClickedIndex > MenuList.ListItems.Count Then GetClickedIndex = MenuList.ListItems.Count
End Function

Private Function GetClickedColumn(ByVal X As Long) As Long
    Dim i As Long
    GetClickedColumn = 0
    For i = 1 To MenuList.ColumnHeaders.Count
        X = X - MenuList.ColumnHeaders(i).Width
        If X <= 0 Then
            GetClickedColumn = i
            Exit For
        End If
    Next i
End Function

Private Sub OpenCellEditor(ByVal row As Long, ByVal col As Byte)
    If columnClicked = 0 Or row = 0 Or row > MenuList.ListItems.Count Then Exit Sub
    CellEdit.Top = MenuList.ListItems(row).Top
    CellEdit.Left = MenuList.ColumnHeaders(col).Left + 30
    CellEdit.Width = MenuList.ColumnHeaders(col).Width
    If col = 1 Then
        CellEdit.Text = MenuList.ListItems(row).Text
    Else
        CellEdit.Text = MenuList.ListItems(row).SubItems(col - 1)
    End If
    CellEdit.Visible = True
    CellEdit.SetFocus
    CellEdit.SelStart = 0
    CellEdit.SelLength = Len(CellEdit)
End Sub

Private Sub WriteCell(ByVal row As Long, ByVal col As Byte, ByVal value As String)
    If col = 1 Then
        MenuList.ListItems(row).Text = value
    Else
        MenuList.ListItems(row).SubItems(col - 1) = value
    End If
End Sub

