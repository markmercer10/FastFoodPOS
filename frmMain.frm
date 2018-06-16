VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10785
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   17850
   LinkTopic       =   "Form1"
   ScaleHeight     =   10785
   ScaleWidth      =   17850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   6480
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   8655
      Left            =   12960
      TabIndex        =   2
      Top             =   1920
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   15266
      _Version        =   393216
      Rows            =   30
      Cols            =   3
      FixedCols       =   0
      Appearance      =   0
   End
   Begin VB.TextBox Text1 
      Height          =   4455
      Left            =   6840
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmMain.frx":0000
      Top             =   6120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox Header 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1665
      ScaleWidth      =   16185
      TabIndex        =   0
      Top             =   0
      Width           =   16215
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   16440
      Top             =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditMenu 
         Caption         =   "Menu"
      End
      Begin VB.Menu mnuEditDisplay 
         Caption         =   "Display"
      End
   End
   Begin VB.Menu mnuDisplay 
      Caption         =   "Display"
      Begin VB.Menu mnuDisplayShow 
         Caption         =   "Show Display"
      End
      Begin VB.Menu mnuDisplayHide 
         Caption         =   "Hide Display"
      End
   End
   Begin VB.Menu mnuStats 
      Caption         =   "Statistics"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Const MIM_BACKGROUND As Long = &H2
Private Const MIM_APPLYTOSUBMENUS As Long = &H80000000
 
Private Type MENUINFO
    cbSize As Long
    fMask As Long
    dwStyle As Long
    cyMax As Long
    hbrBack As Long
    dwContextHelpID As Long
    dwMenuData As Long
End Type
 
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, mi As MENUINFO) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const HWND_TOP = 0
Private Const SWP_SHOWWINDOW = &H40

 
Private Sub Command1_Click()
    Dim mi As MENUINFO
   
    With mi
        .cbSize = Len(mi)
        
        .fMask = MIM_BACKGROUND
        .hbrBack = CreateSolidBrush(vbYellow)
        SetMenuInfo GetMenu(Me.hwnd), mi  'main menu bar
        
        .fMask = MIM_BACKGROUND Or MIM_APPLYTOSUBMENUS
        .hbrBack = CreateSolidBrush(vbCyan)
        SetMenuInfo GetSubMenu(GetMenu(Me.hwnd), 0), mi 'this could a File menu perhaps
    End With
    
    DrawMenuBar Me.hwnd
 
End Sub

Private Sub Form_Load()
    Dim cx As Long
    Dim cy As Long
    Dim RetVal As Long
    
    ' Determine if screen is already maximized.
    If Me.WindowState = vbMaximized Then
        ' Set window to normal size
        Me.WindowState = vbNormal
    End If
    ' Get full screen width.
    cx = GetSystemMetrics(SM_CXSCREEN)
    ' Get full screen height.
    cy = GetSystemMetrics(SM_CYSCREEN)
    
    ' Call API to set new size of window.End Sub
    RetVal = SetWindowPos(Me.hwnd, HWND_TOP, 0, 0, cx, cy, SWP_SHOWWINDOW)
    

End Sub

Private Sub GraphicalButton1_Click()
    Text1.Text = Text1.Text & vbCrLf & "Click    "
End Sub

Private Sub Header_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Timer1_Timer()
    Header.SetFocus
End Sub
