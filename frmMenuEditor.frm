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
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1215
      Left            =   11520
      TabIndex        =   2
      Top             =   1920
      Width           =   495
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
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      Height          =   1575
      Left            =   11280
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMenuEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim indexDrag As Long

Private Sub Command1_Click()
    Dim i As Long
    Dim li As ListItem
    For i = 1 To 20
        Set li = MenuList.ListItems.Add(, , "Menu Item " & i)
        li.SubItems(4) = i
        'MenuList.List(
    Next i
    Set li = Nothing
End Sub

Private Sub MenuList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this sets the indexDrag to the current index
    If Not MenuList.SelectedItem Is Nothing Then
        indexDrag = GetClickedIndex(Y)
    End If
End Sub
 
Private Sub MenuList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not MenuList.SelectedItem Is Nothing Then
        Dim indexDragTo As Long
        indexDragTo = GetClickedIndex(Y)
        If Not indexDragTo = indexDrag Then SwapItems indexDrag, indexDragTo
    End If
End Sub

Private Function SwapItems(ByVal Item1 As Long, ByVal Item2 As Long)
    If Item1 < 1 Or Item1 > MenuList.ListItems.Count Or Item2 < 1 Or Item2 > MenuList.ListItems.Count Then Exit Function
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
    Dim ItemHeight As Long
    ItemHeight = MenuList.ListItems.Item(1).Height
    HeaderExtraHeight = ItemHeight * 0.15

    GetClickedIndex = Int((MenuList.GetFirstVisible.Index - 1) + (Y - HeaderExtraHeight) / ItemHeight)
End Function
