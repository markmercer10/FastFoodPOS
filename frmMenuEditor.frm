VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenuEditor 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Editor"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   13065
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox deletedSections 
      Height          =   840
      Left            =   480
      TabIndex        =   11
      Top             =   5160
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.ListBox deletedItems 
      Height          =   840
      Left            =   480
      TabIndex        =   10
      Top             =   6120
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton ButnView 
      BackColor       =   &H00C0C0C0&
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton ButnSave 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton ButnDeleteItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Delete Item"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton ButnAddItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add Item"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton ButnDeleteSection 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Delete Section"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton ButnAddSection 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add Section"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Ghost 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   0
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   10935
   End
   Begin VB.TextBox CellEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   9360
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComctlLib.ListView MenuList 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
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
      BackColor       =   14737632
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
         Object.Width           =   7938
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
End
Attribute VB_Name = "frmMenuEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ITEM_PREFIX = " -- "
'Menu Editor
Dim indexDrag As Long
Dim columnClicked As Byte
Dim ItemHeight As Long
Dim dragging As Boolean


Private Sub ButnAddItem_Click()
    Dim itemname As String
    itemname = InputBox("Enter the item name", "Add Item", "Item Name")
    If stringValidateSQL(itemname) Then
        AddItem itemname, -1, 0, 0, 0, 0
    Else
        MsgBox "Illegal characters in name"
    End If
End Sub

Private Sub ButnAddSection_Click()
    Dim sectionname As String
    sectionname = InputBox("Enter a section name", "Add Section", "Section")
    If stringValidateSQL(sectionname) Then
        AddSection sectionname, -1 * Int(Rnd * 10000)
    Else
        MsgBox "Illegal characters in name"
    End If
End Sub

Private Sub ButnDeleteItem_Click()
    Dim i As Long
    i = MenuList.SelectedItem.Index
    If LineExists(i) Then
        If IsItem(i) Then
            If GetLineID(i) >= 0 Then
                deletedItems.AddItem GetLineID(i)
            End If
            MenuList.ListItems.Remove i
            If MenuList.ListItems.Count = 0 Then
                ButnAddItem.Enabled = False
                ButnDeleteSection.Enabled = False
            Else
                MenuList.SelectedItem = MenuList.ListItems(1)
                MenuList_ItemClick MenuList.ListItems(1)
            End If
        End If
    End If
End Sub

Private Sub ButnDeleteSection_Click()
    Dim i As Long
    i = MenuList.SelectedItem.Index
    If LineExists(i) Then
        If IsSection(i) And Not SectionHasItems(i) Then
            If GetLineID(i) >= 0 Then
                deletedSections.AddItem GetLineID(i)
            End If
            MenuList.ListItems.Remove i
            If MenuList.ListItems.Count = 0 Then
                ButnAddItem.Enabled = False
                ButnDeleteSection.Enabled = False
            Else
                MenuList.SelectedItem = MenuList.ListItems(1)
                MenuList_ItemClick MenuList.ListItems(1)
            End If
        End If
    End If
End Sub

Private Sub ButnSave_Click()
    Dim i As Long
    Dim sec_id As Long
    Dim item_id As Long
    Dim f As Variant
    Dim v As Variant
    Dim j As Long
    For i = 0 To deletedItems.ListCount - 1
        Delete "menu_items", "id", CLng(val(deletedItems.List(i)))
    Next i
    For i = 0 To deletedSections.ListCount - 1
        Delete "menu_sections", "id", CLng(val(deletedSections.List(i)))
    Next i
    
    For i = 1 To MenuList.ListItems.Count
        If IsSection(i) Then
            sec_id = GetLineID(i)
            If sec_id < 0 Then sec_id = 0
            f = Array("id", "title")
            v = Array(sec_id, MenuList.ListItems(i).text)
            Upsert "menu_sections", f, v
            If sec_id = 0 Then sec_id = QuerySectionID(MenuList.ListItems(i).text)
        ElseIf IsItem(i) Then
            item_id = GetLineID(i)
            If item_id < 0 Then item_id = 0
            f = Array("id", "section_id", "sort_id", "name", "price", "small", "medium", "large")
            v = Array(item_id, sec_id, i, Replace(MenuList.ListItems(i).text, ITEM_PREFIX, ""), val(MenuList.ListItems(i).SubItems(1)), val(MenuList.ListItems(i).SubItems(2)), val(MenuList.ListItems(i).SubItems(3)), val(MenuList.ListItems(i).SubItems(4)))
            Upsert "menu_items", f, v
        End If
    Next i
    LoadMenu
End Sub

Private Sub CellEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WriteCell indexDrag, columnClicked, CellEdit.text
        CellEdit.Visible = False
    End If
End Sub

Private Sub CellEdit_LostFocus()
    CellEdit.Visible = False
End Sub

Private Sub Form_Load()
    LoadMenu
    Ghost.fontname = MenuList.Font.name
    Ghost.FontSize = MenuList.Font.Size
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HR.Visible = False
    Ghost.Visible = False
End Sub

Private Sub HR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MenuList_MouseUp Button, Shift, X, HR.Top + 100
    CellEdit = HR.Top
End Sub

Private Sub MenuList_DblClick()
    'MsgBox GetSectionID(MenuList.SelectedItem.Index)
    OpenCellEditor indexDrag, columnClicked
End Sub

Private Sub MenuList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ButnDeleteSection.Enabled = IsSection(Item.Index) And Not SectionHasItems(Item.Index)
    ButnDeleteItem.Enabled = IsItem(Item.Index)
End Sub

Private Sub MenuList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not MenuList.SelectedItem Is Nothing Then
        indexDrag = GetClickedIndex(Y)
        If indexDrag > 1 Then
            CellEdit = MenuList.ListItems(indexDrag).Tag
            If Not IsSection(indexDrag) Then dragging = True
        End If
    End If
End Sub
 
Private Sub MenuList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 And dragging Then
        Dim i As Long
        i = GetClickedIndex(Y)
        If i <= 0 Then
            HR.Top = MenuList.ListItems(MenuList.ListItems.Count).Top + MenuList.ListItems(MenuList.ListItems.Count).Height
        Else
            HR.Top = MenuList.ListItems(GetClickedIndex(Y)).Top
        End If
        Ghost.text = MenuList.ListItems(indexDrag).text
        Ghost.Top = Y + Ghost.Height / 2
        HR.Visible = True
        Ghost.Visible = True
    Else
        HR.Visible = False
        Ghost.Visible = False
    End If
End Sub

Private Sub MenuList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    columnClicked = GetClickedColumn(X)
    If dragging Then
        If Not MenuList.SelectedItem Is Nothing Then
            Dim indexDragTo As Long
            indexDragTo = GetClickedIndex(Y)
            If indexDragTo = -1 Then
                MoveItem indexDrag, MenuList.ListItems.Count + 1
            Else
                If Not indexDragTo = indexDrag Then MoveItem indexDrag, indexDragTo
            End If
        Else
        End If
        dragging = False
    End If
    HR.Visible = False
    Ghost.Visible = False
End Sub

Private Function MoveItem(ByVal Item1 As Long, ByVal Item2 As Long)
    If Item2 = 1 Then Item2 = 2
    If Item2 > Item1 Then Item2 = Item2 - 1
    If Not LineExists(Item1) Or Not LineExists(Item2) Then Exit Function
    Dim oli As ListItem
    Dim nli As ListItem
    Dim si As ListSubItem
    Dim itmTag As String
    
    Set oli = MenuList.ListItems(Item1)
    itmTag = oli.Tag
    MenuList.ListItems.Remove (Item1)
    Set nli = MenuList.ListItems.Add(Item2, , "")
    nli = oli
    nli.Tag = itmTag
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
    If GetClickedIndex > MenuList.ListItems.Count Then GetClickedIndex = -1 'MenuList.ListItems.Count
End Function

Private Function GetClickedColumn(ByVal X As Long) As Long
    Dim i As Long
    GetClickedColumn = 0
    For i = 1 To MenuList.columnHeaders.Count
        X = X - MenuList.columnHeaders(i).Width
        If X <= 0 Then
            GetClickedColumn = i
            Exit For
        End If
    Next i
End Function

Private Sub OpenCellEditor(ByVal row As Long, ByVal col As Byte)
    If col = 0 Or Not LineExists(row) Then Exit Sub
    CellEdit.Top = MenuList.ListItems(row).Top
    CellEdit.Left = MenuList.columnHeaders(col).Left + 30
    CellEdit.Width = MenuList.columnHeaders(col).Width
    If col = 1 Then
        If IsSection(row) Then
            CellEdit.text = MenuList.ListItems(row).text
        Else
            CellEdit.text = Replace(MenuList.ListItems(row).text, ITEM_PREFIX, "")
        End If
    Else
        CellEdit.text = MenuList.ListItems(row).SubItems(col - 1)
    End If
    CellEdit.Visible = True
    CellEdit.SetFocus
    CellEdit.SelStart = 0
    CellEdit.SelLength = Len(CellEdit)
End Sub

Private Sub WriteCell(ByVal row As Long, ByVal col As Byte, ByVal value As String)
    If col = 1 Then
        If IsSection(row) Then
            MenuList.ListItems(row).text = value
        Else
            MenuList.ListItems(row).text = ITEM_PREFIX & value
        End If
    Else
        MenuList.ListItems(row).SubItems(col - 1) = FormatAmount(val(value))
    End If
End Sub

Private Function AddSection(ByVal name As String, ByVal id As Long)
    Dim li As ListItem
    Set li = MenuList.ListItems.Add(, , name)
    li.Tag = "Section(" & id & ")"
    Set li = Nothing
    ButnAddItem.Enabled = True
End Function

Private Function AddItem(ByVal name As String, ByVal id As Long, ByVal price As Double, ByVal small As Double, ByVal medium As Double, ByVal large As Double)
    Dim li As ListItem
    Set li = MenuList.ListItems.Add(, , ITEM_PREFIX & name)
    li.SubItems(1) = FormatAmount(price)
    li.SubItems(2) = FormatAmount(small)
    li.SubItems(3) = FormatAmount(medium)
    li.SubItems(4) = FormatAmount(large)
    li.Tag = "Item(" & id & ")"
    Set li = Nothing
End Function

Private Function FormatAmount(ByVal val As Double) As String
    If val = 0 Then
        FormatAmount = ""
    Else
        FormatAmount = Format(val, "0.00")
    End If
End Function

Private Function IsSection(ByVal linenumber As Long) As Boolean
    If Not LineExists(linenumber) Then
        IsSection = False
        Exit Function
    End If
    IsSection = InStr(1, MenuList.ListItems(linenumber).Tag, "Section") > 0
End Function

Private Function IsItem(ByVal linenumber As Long) As Boolean
    If Not LineExists(linenumber) Then
        IsItem = False
        Exit Function
    End If
    IsItem = InStr(1, MenuList.ListItems(linenumber).Tag, "Item") > 0
End Function

Private Function GetLineID(ByVal linenumber As Long) As Long
    Dim start As Long
    Dim length As Byte
    GetLineID = -1
    start = InStr(1, MenuList.ListItems(linenumber).Tag, "(") + 1
    If start > 0 Then
        length = InStr(1, MenuList.ListItems(linenumber).Tag, ")") - start
        GetLineID = val(Mid(MenuList.ListItems(linenumber).Tag, start, length))
    End If
End Function

Private Function GetSectionID(ByVal linenumber As Long) As Long
    GetSectionID = -1
    Do Until linenumber < 1 Or IsSection(linenumber)
        linenumber = linenumber - 1
    Loop
    If linenumber >= 1 Then GetSectionID = GetLineID(linenumber)
End Function

Private Function SectionHasItems(ByVal linenumber As Long) As Boolean
    Dim sec As Long
    SectionHasItems = False
    If IsSection(linenumber) Then
        sec = GetSectionID(linenumber)
        If LineExists(linenumber + 1) Then
            SectionHasItems = (sec = GetSectionID(linenumber + 1))
        End If
    End If
End Function

Private Function LineExists(ByVal linenumber As Long) As Boolean
    LineExists = Not (linenumber < 1 Or linenumber > MenuList.ListItems.Count)
End Function

Private Sub LoadMenu()
    Dim s As ADODB.Recordset
    Dim i As ADODB.Recordset
    MenuList.ListItems.Clear
    
    If Not CheckDBConnection Then Exit Sub
    Set s = db.Execute("SELECT * FROM menu_sections")
    With s
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                AddSection !Title, !id
                Set i = db.Execute("SELECT * FROM menu_items WHERE section_id = " & !id & " ORDER BY sort_id ASC")
                With i
                    If Not (.EOF And .BOF) Then
                        .MoveFirst
                        Do Until .EOF
                            AddItem !name, !id, val("" & !price), val("" & !small), val("" & !medium), val("" & !large)
                            .MoveNext
                        Loop
                    End If
                End With
                .MoveNext
            Loop
        End If
    End With
    
    Set s = Nothing
    Set i = Nothing
End Sub

Private Function QuerySectionID(ByVal Section As String) As Long
    Dim q As ADODB.Recordset
    QuerySectionID = -1
    If Not CheckDBConnection Then Exit Sub
    Set q = db.Execute("SELECT * FROM menu_sections WHERE Title = """ & Section & """")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            QuerySectionID = !id
        End If
    End With
    Set q = Nothing
End Function
