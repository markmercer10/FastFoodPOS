VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDisplayPanels 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Display Panels"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   19635
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   19080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer LoadTimer 
      Interval        =   50
      Left            =   240
      Top             =   840
   End
   Begin VB.CommandButton ButnSave 
      BackColor       =   &H0080FFFF&
      Caption         =   "Save Changes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   120
      Width           =   2535
   End
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   720
      Top             =   840
   End
   Begin VB.ComboBox cboRes 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "frmDisplayPanels.frx":0000
      Left            =   5640
      List            =   "frmDisplayPanels.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox BlockBG 
      Height          =   495
      Left            =   16920
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox labColor 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Background Color"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame TabContent 
      BorderStyle     =   0  'None
      Height          =   6855
      Index           =   0
      Left            =   1200
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox Margin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   42
         Text            =   "10"
         Top             =   840
         Width           =   975
      End
      Begin VB.ListBox SectionIDs 
         Height          =   255
         Left            =   3240
         TabIndex        =   38
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox Section 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   240
         Width           =   2775
      End
      Begin VB.PictureBox iFontColor 
         Height          =   495
         Left            =   1920
         ScaleHeight     =   435
         ScaleWidth      =   915
         TabIndex        =   34
         Top             =   5640
         Width           =   975
      End
      Begin VB.CheckBox iFontBold 
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
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   33
         Top             =   5220
         Width           =   1215
      End
      Begin VB.TextBox iFontSize 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   32
         Text            =   "12"
         Top             =   4800
         Width           =   975
      End
      Begin VB.ComboBox iFont 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   4320
         Width           =   2775
      End
      Begin VB.PictureBox hFontColor 
         Height          =   495
         Left            =   1920
         ScaleHeight     =   435
         ScaleWidth      =   915
         TabIndex        =   25
         Top             =   3120
         Width           =   975
      End
      Begin VB.CheckBox hFontBold 
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
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   24
         Top             =   2700
         Width           =   1215
      End
      Begin VB.TextBox hFontSize 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   23
         Text            =   "12"
         Top             =   2280
         Width           =   975
      End
      Begin VB.ComboBox hFont 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label Label15 
         Caption         =   "Heading"
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
         Left            =   240
         TabIndex        =   41
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label13 
         Caption         =   "Menu Section"
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
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Color"
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
         Left            =   840
         TabIndex        =   31
         Top             =   5760
         Width           =   2655
      End
      Begin VB.Label Label11 
         Caption         =   "Bold"
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
         Left            =   840
         TabIndex        =   30
         Top             =   5280
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "Font Size"
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
         Left            =   840
         TabIndex        =   29
         Top             =   4800
         Width           =   2655
      End
      Begin VB.Label Label9 
         Caption         =   "Font"
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
         Left            =   840
         TabIndex        =   28
         Top             =   4320
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Items"
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
         Left            =   240
         TabIndex        =   26
         Top             =   3840
         Width           =   2775
      End
      Begin VB.Label Label7 
         Caption         =   "Color"
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
         Left            =   840
         TabIndex        =   22
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Bold"
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
         Left            =   840
         TabIndex        =   21
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Font Size"
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
         Left            =   840
         TabIndex        =   20
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Font"
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
         Left            =   840
         TabIndex        =   19
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Margin"
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
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   2775
      End
   End
   Begin VB.Frame TabContent 
      BorderStyle     =   0  'None
      Height          =   6855
      Index           =   1
      Left            =   9960
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   4935
      Begin VB.OptionButton opFit 
         Caption         =   "Fit to Box"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   45
         Top             =   5520
         Width           =   1575
      End
      Begin VB.OptionButton opStretch 
         Caption         =   "Stretch && Crop"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   44
         Top             =   5520
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.PictureBox PicturePreview 
         AutoRedraw      =   -1  'True
         Height          =   4575
         Left            =   120
         ScaleHeight     =   4515
         ScaleWidth      =   4515
         TabIndex        =   16
         Top             =   840
         Width           =   4575
      End
      Begin VB.CommandButton butnSelectPicture 
         Caption         =   "Select Picture"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Frame TabContent 
      BorderStyle     =   0  'None
      Height          =   6855
      Index           =   2
      Left            =   14520
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Label Label2 
         Caption         =   "This feature can be added"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   14
         Top             =   120
         Width           =   3975
      End
   End
   Begin MSComctlLib.TabStrip BlockOptions 
      Height          =   8175
      Left            =   14520
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   14420
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Menu Section"
            Key             =   "blockMenu"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Picture"
            Key             =   "blockPicture"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Video"
            Key             =   "blockVideo"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frameLayouts 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Layouts"
      Height          =   8175
      Left            =   7800
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   6735
      Begin VB.PictureBox Layouts 
         AutoRedraw      =   -1  'True
         Height          =   2895
         Index           =   0
         Left            =   2400
         ScaleHeight     =   2835
         ScaleWidth      =   435
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton LayoutOptions 
         Height          =   1095
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton butnSelectLayout 
      Caption         =   "Select Layout"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7800
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.PictureBox Preview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   120
      ScaleHeight     =   8025
      ScaleWidth      =   14295
      TabIndex        =   2
      Top             =   720
      Width           =   14320
      Begin VB.PictureBox Blocks 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1695
         Index           =   0
         Left            =   0
         ScaleHeight     =   1695
         ScaleWidth      =   1935
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.ComboBox DisplaySelect 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "frmDisplayPanels.frx":0047
      Left            =   1200
      List            =   "frmDisplayPanels.frx":0051
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label14 
      Caption         =   "Resolution"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   40
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label labBlockOptions 
      Caption         =   "Block Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14520
      TabIndex        =   37
      Top             =   120
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label1 
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
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmDisplayPanels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SelectedDisplay As Long
Private SelectedLayout As Long
Private SelectedBlock As Long
Private MouseOverBlock As Long
Private BlockSettings() As PanelBlock
Private PreviewScale As Double
Private BlockCount As Byte
Private MadeChanges As Boolean
Private ChangeTracking As Boolean

Private Sub BlockBG_Click()
    BlockBG_DblClick
End Sub

Private Sub BlockBG_DblClick()
    Load frmColorSelect
    frmColorSelect.Tag = BlockBG.BackColor
    frmColorSelect.Show 1
    BlockBG.BackColor = csSelectedColor
    BlockSettings(SelectedBlock).Background = BlockBG.BackColor
    PicturePreview.BackColor = BlockBG.BackColor
    PaintPic BlockSettings(SelectedBlock).PicturePath, PicturePreview, False
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub BlockOptions_Click()
    Dim i As Byte
    For i = 0 To 2
        If i = BlockOptions.SelectedItem.Index - 1 Then
            TabContent(i).Visible = True
            BlockSettings(SelectedBlock).ContentType = i
        Else
            TabContent(i).Visible = False
        End If
    Next i
    BlockSettings(SelectedBlock).DrawBlock Blocks(SelectedBlock), PreviewScale
End Sub

Private Sub BlockOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseOver
End Sub

Private Sub Blocks_Click(Index As Integer)
    If SelectedBlock <> Index Then
        ChangeTracking = False
        SelectedBlock = Index ' this sets the index of the selected block
        Block_BorderChanged
        BlockSelected Index ' this populates the block options for the block that was selected
        ChangeTracking = True
    End If
End Sub

Private Sub Blocks_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MouseOverBlock <> Index Then
        MouseOverBlock = Index
        Block_BorderChanged
    End If
End Sub

Private Sub Block_BorderChanged()
    Dim i As Long
    Dim state As String
    For i = 0 To Blocks.Count - 1
        state = "Default"
        If i = SelectedBlock Then state = "Selected"
        If i = MouseOverBlock Then state = "MouseOver"
        DrawBlockBorder i, state
    Next i
End Sub

Private Sub ButnSave_Click()
    db.Execute "UPDATE displays SET " & _
    "resolution_x = " & left$(cboRes.Text, 4) & _
    ", resolution_y = " & Mid$(cboRes.Text, 8) & _
    ", layout = " & SelectedLayout & _
    " WHERE id = " & SelectedDisplay + 1
    
    Dim i As Long
    For i = 0 To BlockCount - 1
        BlockSettings(i).Save DisplaySelect.ListIndex + 1, i
    Next i
    MadeChanges = False
End Sub

Private Sub butnSelectLayout_Click()
    frameLayouts.Visible = Not frameLayouts.Visible
End Sub

Private Sub butnSelectPicture_Click()
    CommonDialog.Filter = "Jpeg (*.jpg)|*.jpg|PNG (*.png)|*.png|Bitmap (*.bmp)|*.bmp"
    CommonDialog.DefaultExt = "jpg"
    CommonDialog.DialogTitle = "Select Picture"
    CommonDialog.ShowOpen
    
    If CommonDialog.FileName <> "" Then
        BlockSettings(SelectedBlock).PicturePath = CommonDialog.FileName
        PaintPic BlockSettings(SelectedBlock).PicturePath, PicturePreview, False
        DrawTimer.Enabled = True
        If ChangeTracking Then MadeChanges = True
    End If
End Sub

Private Sub cboRes_Click()
    PreviewScale = (Preview.Width / Screen.TwipsPerPixelX) / val(left$(cboRes.Text, 4))
    DrawAllBlocks
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub DisplaySelect_Click()
    If DisplaySelect.ListIndex <> SelectedDisplay Then
        If MadeChanges Then
            If MsgBox("You have made changes to this display, by switching you will loose your changes unless you save them first.  Do you want to switch displays now?", vbYesNo) = vbYes Then
                LoadTimer.Enabled = True
                SelectedDisplay = DisplaySelect.ListIndex
            Else
                DisplaySelect.ListIndex = SelectedDisplay
            End If
        Else
            LoadTimer.Enabled = True
            SelectedDisplay = DisplaySelect.ListIndex
        End If
    End If
End Sub

Private Sub DrawTimer_Timer()
    DrawTimer.Enabled = False
    BlockSettings(SelectedBlock).DrawBlock Blocks(SelectedBlock), PreviewScale
    Block_BorderChanged
End Sub

Private Sub Form_Load()
    Const OFFSET = 250
    
    Dim i As Long
    Dim b As OptionButton
    Dim c As OLE_COLOR
    Dim Margin As Long
    Dim h As Long
    Dim w As Long
    Dim v() As Double
    Dim block As Long
    Dim l As Long
    Dim t As Long
    Dim q As ADODB.Recordset
    Dim li As ListItem
    
    c = &HBBBBBB
    Margin = 40
    SelectedDisplay = 0
    ChangeTracking = False
    
    For i = 0 To 2
        TabContent(i).left = BlockOptions.left
        TabContent(i).Top = 1800
    Next i
    For i = 0 To Screen.FontCount - 1
        If left$(Screen.Fonts(i), 1) <> "@" Then
            hFont.AddItem Screen.Fonts(i)
            iFont.AddItem Screen.Fonts(i)
        End If
    Next
    
    ' have to save these values in the database
    DisplaySelect.ListIndex = 0
    cboRes.ListIndex = 1
    
    Layouts(0).Width = LayoutOptions(0).Width - 200
    Layouts(0).Height = LayoutOptions(0).Height - 200
    
    For i = 1 To 15
        Load LayoutOptions(i)
        Load Layouts(i)
    Next i
    For Each b In LayoutOptions
        LayoutOptions(b.Index).Picture = Layouts(b.Index).Image
        LayoutOptions(b.Index).left = LayoutOptions(b.Index).Width * 1.1 * (b.Index Mod 3) + OFFSET
        LayoutOptions(b.Index).Top = LayoutOptions(b.Index).Height * 1.2 * Int(b.Index / 3) + OFFSET
        LayoutOptions(b.Index).Visible = True
    Next b
    Set b = Nothing
    
    Set q = db.Execute("SELECT * FROM menu_sections")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                Section.AddItem !Title
                SectionIDs.AddItem !id
                .MoveNext
            Loop
        End If
    End With
    Set q = Nothing
    Set li = Nothing
    
    w = Layouts(0).Width
    h = Layouts(0).Height
    
    Rnd -1
    Randomize 314 '314
    For i = 0 To Layouts.Count - 1
        Let v = GetLayout(i)
        For block = 0 To UBound(v)
            l = w * v(block, 1) + Margin
            t = h * v(block, 2) + Margin
            Layouts(i).Line (l, t)-(l + w * v(block, 3) - Margin * 2, t + h * v(block, 4) - Margin * 2), Rnd * 16777216, BF
        Next block
    Next i
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseOver
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MadeChanges Then
        If MsgBox("You have made changes to this display, by closing you will loose your changes unless you save them first.  Are you sure you want to close this window now?", vbYesNo) = vbNo Then
            Cancel = True
        End If
    End If
End Sub

Private Sub hFont_Click()
    BlockSettings(SelectedBlock).HFontName = hFont.Text
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub hFontBold_Click()
    BlockSettings(SelectedBlock).hFontBold = CBool(hFontBold)
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub hFontColor_Click()
    hFontColor_DblClick
End Sub

Private Sub hFontColor_DblClick()
    Load frmColorSelect
    frmColorSelect.Tag = hFontColor.BackColor
    frmColorSelect.Show 1
    hFontColor.BackColor = csSelectedColor
    BlockSettings(SelectedBlock).hFontColor = hFontColor.BackColor
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub hFontSize_Change()
    BlockSettings(SelectedBlock).hFontSize = val(hFontSize)
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub iFont_Click()
    BlockSettings(SelectedBlock).IFontName = iFont.Text
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub iFontBold_Click()
    BlockSettings(SelectedBlock).iFontBold = CBool(iFontBold)
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub iFontColor_Click()
    iFontColor_DblClick
End Sub

Private Sub iFontColor_DblClick()
    Load frmColorSelect
    frmColorSelect.Tag = iFontColor.BackColor
    frmColorSelect.Show 1
    iFontColor.BackColor = csSelectedColor
    BlockSettings(SelectedBlock).iFontColor = iFontColor.BackColor
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub iFontSize_Change()
    BlockSettings(SelectedBlock).iFontSize = val(iFontSize)
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub LayoutOptions_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    LoadLayout Index
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub LoadLayout(ByVal Index As Long)
    Dim v() As Double
    Dim block As Long
    SelectedLayout = Index
    Let v = GetLayout(Index)
    BlockCount = UBound(v) + 1
    ReDim BlockSettings(BlockCount - 1)
    
    For block = 0 To BlockCount - 1
        If block > Blocks.Count - 1 Then Load Blocks(block)
        Blocks(block).left = Preview.Width * v(block, 1)
        Blocks(block).Top = Preview.Height * v(block, 2)
        Blocks(block).Width = Preview.Width * v(block, 3)
        Blocks(block).Height = Preview.Height * v(block, 4)
        
        InitializeBlockSettings block
        
        BlockSettings(block).DrawBlock Blocks(block), PreviewScale
        DrawBlockBorder block, "default"
        
        Blocks(block).Visible = True
    Next block
    frameLayouts.Visible = False
    Blocks_Click -1

End Sub

Private Sub InitializeBlockSettings(ByVal Index As Long)
    Set BlockSettings(Index) = New PanelBlock
    
    BlockSettings(Index).Load DisplaySelect.ListIndex + 1, Index
    If BlockSettings(Index).SectionID = -1 Then
        If SectionIDs.ListCount > 0 Then
            BlockSettings(Index).SectionID = val(SectionIDs.List(0))
        End If
    End If
    
End Sub

Private Sub DrawBlockBorder(ByVal Index As Long, ByVal state As String)
    Dim c As Long
    If state = "MouseOver" Then
        c = vbYellow
    ElseIf state = "Selected" Then
        c = vbRed
    Else '         "Default"
        c = &H666666 'vbBlack
    End If
    
    Blocks(Index).Line (0, 0)-(Blocks(Index).Width - 30, Blocks(Index).Height - 45), c, B
    Blocks(Index).Line (15, 15)-(Blocks(Index).Width - 45, Blocks(Index).Height - 60), c, B
End Sub

Private Sub BlockSelected(ByVal Index As Long)
    If Index = -1 Then
        BlockOptions.Visible = False
        TabContent(0).Visible = False
        TabContent(1).Visible = False
        TabContent(2).Visible = False
        labBlockOptions.Visible = False
        labColor.Visible = False
        BlockBG.Visible = False
    Else
        BlockOptions.Visible = True
        labBlockOptions.Visible = True
        labColor.Visible = True
        BlockBG.Visible = True
        
        BlockOptions.SelectedItem = BlockOptions.Tabs(BlockSettings(Index).ContentType + 1)
        BlockOptions_Click
        
        BlockBG.BackColor = BlockSettings(Index).Background
        PicturePreview.BackColor = BlockBG.BackColor
        
        'load settings from blocksettings
        Section.ListIndex = GetSectionIndex(BlockSettings(Index).SectionID)
        Margin = BlockSettings(Index).Margin
        hFont.ListIndex = GetFontIndex(BlockSettings(Index).HFontName)
        hFontSize = BlockSettings(Index).hFontSize
        hFontBold = -CLng(BlockSettings(Index).hFontBold)
        hFontColor.BackColor = BlockSettings(Index).hFontColor
        iFont.ListIndex = GetFontIndex(BlockSettings(Index).IFontName)
        iFontSize = BlockSettings(Index).iFontSize
        iFontBold = -CLng(BlockSettings(Index).iFontBold)
        iFontColor.BackColor = BlockSettings(Index).iFontColor
        opStretch.value = BlockSettings(Index).Stretch
        opFit.value = Not BlockSettings(Index).Stretch
        PaintPic BlockSettings(SelectedBlock).PicturePath, PicturePreview, False
    End If
End Sub

Private Sub ClearMouseOver()
    MouseOverBlock = -1
    Block_BorderChanged
End Sub

Private Sub LoadTimer_Timer()
    LoadTimer.Enabled = False
    
    Dim i As Byte
    Set q = db.Execute("SELECT * FROM displays WHERE id = " & DisplaySelect.ListIndex + 1)
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            If IsNull(!layout) Then
                LoadLayout 0
            Else
                LoadLayout !layout
            End If
            For i = 0 To BlockCount - 1
                BlockSettings(i).Load DisplaySelect.ListIndex + 1, i
                'BlockSettings(i).DrawBlock Blocks(i), PreviewScale
                'MsgBox BlockSettings(i).PicturePath
            Next i
            For i = 0 To cboRes.ListCount - 1
                If left$(cboRes.List(i), 4) = !resolution_x Then
                    cboRes.ListIndex = i
                    cboRes_Click
                End If
            Next i
        End If
    End With
    
    MadeChanges = False
End Sub

Private Sub Margin_Change()
    BlockSettings(SelectedBlock).Margin = val(Margin)
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub opFit_Click()
    BlockSettings(SelectedBlock).Stretch = False
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub opStretch_Click()
    BlockSettings(SelectedBlock).Stretch = True
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub Section_Click()
    BlockSettings(SelectedBlock).SectionID = val(SectionIDs.List(Section.ListIndex))
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub TabContent_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseOver
End Sub

Private Function GetSectionIndex(ByVal id As Long) As Long
    Dim i As Long
    GetSectionIndex = -1
    For i = 0 To SectionIDs.ListCount - 1
        If SectionIDs.List(i) = id Then GetSectionIndex = i
    Next i
End Function

Private Function GetFontIndex(ByVal fontname As String) As Long
    Dim i As Long
    GetFontIndex = -1
    For i = 0 To hFont.ListCount - 1
        If hFont.List(i) = fontname Then
            GetFontIndex = i
            Exit For
        End If
    Next i
End Function

Private Sub DrawAllBlocks()
    For block = 0 To BlockCount - 1
        If Blocks.Count > block Then
            BlockSettings(block).DrawBlock Blocks(block), PreviewScale
            DrawBlockBorder block, "default"
        End If
    Next block
End Sub
