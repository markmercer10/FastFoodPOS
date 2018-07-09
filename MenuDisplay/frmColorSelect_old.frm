VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColorSelect 
   Caption         =   "Color Selector"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   231
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   448
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6120
      Top             =   120
   End
   Begin VB.CommandButton okButn 
      Caption         =   "OK"
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox tB 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3480
      TabIndex        =   11
      Text            =   "0"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox tG 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   3480
      TabIndex        =   8
      Text            =   "0"
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox selectedColor 
      Height          =   1215
      Left            =   3000
      ScaleHeight     =   1155
      ScaleWidth      =   1755
      TabIndex        =   5
      Top             =   360
      Width           =   1815
   End
   Begin VB.PictureBox nearestColor 
      Height          =   1215
      Left            =   4800
      ScaleHeight     =   1155
      ScaleWidth      =   1755
      TabIndex        =   4
      ToolTipText     =   "The nearest even hex color"
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox tR 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Text            =   "0"
      Top             =   1680
      Width           =   495
   End
   Begin MSComctlLib.Slider SliderR 
      Height          =   315
      Left            =   3960
      TabIndex        =   1
      Top             =   1680
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      LargeChange     =   25
      Max             =   255
      TickStyle       =   3
      TickFrequency   =   10
      TextPosition    =   1
   End
   Begin VB.PictureBox BigSwatch 
      Height          =   3255
      Left            =   120
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin MSComctlLib.Slider SliderG 
      Height          =   315
      Left            =   3960
      TabIndex        =   6
      Top             =   2040
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      LargeChange     =   25
      Max             =   255
      TickStyle       =   3
      TickFrequency   =   10
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider SliderB 
      Height          =   315
      Left            =   3960
      TabIndex        =   9
      Top             =   2400
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      LargeChange     =   25
      Max             =   255
      TickStyle       =   3
      TickFrequency   =   10
      TextPosition    =   1
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   4680
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label htmlHex 
      Alignment       =   1  'Right Justify
      Caption         =   "FFFFFF"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "HTML Hex : "
      Height          =   255
      Left            =   3000
      TabIndex        =   18
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label vbHex 
      Alignment       =   1  'Right Justify
      Caption         =   "FFFFFF"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3720
      TabIndex        =   16
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "VB Hex : "
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Nearest even hex"
      Height          =   255
      Left            =   4800
      TabIndex        =   14
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Selected Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Blue"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Green"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Red"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   1680
      Width           =   615
   End
End
Attribute VB_Name = "frmColorSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub colorChange()
    Dim col As Long
    Dim nR As Long
    Dim nG As Long
    Dim nB As Long
    col = selectedColor.BackColor
    Shape1.FillColor = col
    tR = Int(col And &HFF)
    tG = Int(Int(col / 256#) And &HFF)
    tB = Int(Int(col / 65536#) And &HFF)
    htmlHex = Hex2D(tR) & Hex2D(tG) & Hex2D(tB)
    vbHex = "&&H" & Hex2D(tB) & Hex2D(tG) & Hex2D(tR)
    
    'MsgBox val("&H33")
    If tR <= 25 Then
        nR = 0
    ElseIf tR <= 76 Then
        nR = 51
    ElseIf tR <= 127 Then
        nR = 102
    ElseIf tR <= 178 Then
        nR = 153
    ElseIf tR <= 229 Then
        nR = 204
    Else
        nR = 255
    End If
    If tG <= 25 Then
        nG = 0
    ElseIf tG <= 76 Then
        nG = 51
    ElseIf tG <= 127 Then
        nG = 102
    ElseIf tG <= 178 Then
        nG = 153
    ElseIf tG <= 229 Then
        nG = 204
    Else
        nG = 255
    End If
    If tB <= 25 Then
        nB = 0
    ElseIf tB <= 76 Then
        nB = 51
    ElseIf tB <= 127 Then
        nB = 102
    ElseIf tB <= 178 Then
        nB = 153
    ElseIf tB <= 229 Then
        nB = 204
    Else
        nB = 255
    End If
    nearestColor.BackColor = "&H" & Hex2D(nB) & Hex2D(nG) & Hex2D(nR)
End Sub

Private Sub Command1_Click()
End Sub

Function convertColorRGB_selector(ByVal R As Double, ByVal g As Double, ByVal b As Double) As Long
    convertColorRGB_selector = (CLng(R * &HFF))
    convertColorRGB_selector = convertColorRGB_selector + (CLng(g * &HFF)) * 256
    convertColorRGB_selector = convertColorRGB_selector + (CLng(b * &HFF)) * 65536
End Function



Function Hex2D(ByVal n As Long) As String
    Hex2D = Hex$(n)
    If Len(Hex2D) < 2 Then Hex2D = "0" & Hex2D
End Function


Private Sub BigSwatch_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    selectedColor.BackColor = BigSwatch.Point(x, y)
    colorChange
End Sub


Private Sub nearestColor_Click()
    selectedColor.BackColor = nearestColor.BackColor
    colorChange
End Sub


Private Sub okButn_Click()
    'Place code here that sets the returned color to a global variable
    Unload Me
End Sub


Private Sub SliderB_Click()
    tB = SliderB.Value
End Sub

Private Sub SliderG_Click()
    tG = SliderG.Value
End Sub


Private Sub SliderR_Click()
    tR = SliderR.Value
End Sub


Private Sub tB_Change()
    If tB >= 0 And tB <= 255 Then SliderB = tB
    selectedColor.BackColor = "&H" & Hex2D(tB) & Hex2D(tG) & Hex2D(tR)
    colorChange
End Sub

Private Sub tG_Change()
    If tG >= 0 And tG <= 255 Then SliderG = tG
    selectedColor.BackColor = "&H" & Hex2D(tB) & Hex2D(tG) & Hex2D(tR)
    colorChange
End Sub

Private Sub Timer1_Timer()
    Dim i As Long
    Dim j As Long
    Dim R As Double
    Dim g As Double
    Dim b As Double
    Dim s As Double
    Dim bright As Double
    Timer1.Enabled = False
    BigSwatch.Visible = False
    BigSwatch.AutoRedraw = True
    DoEvents
    
    For i = 0 To BigSwatch.Width
        For j = 0 To BigSwatch.Height
            R = Sin(((i / BigSwatch.Width) + (3 / 12#)) * 2# * 3.14159) + 0.5
            If R < 0 Then R = 0
            If R > 1 Then R = 1
            g = Sin(((i / BigSwatch.Width) - (1 / 12#)) * 2# * 3.14159) + 0.5
            If g < 0 Then g = 0
            If g > 1 Then g = 1
            b = Sin(((i / BigSwatch.Width) - (5 / 12#)) * 2# * 3.14159) + 0.5
            If b < 0 Then b = 0
            If b > 1 Then b = 1
            
            s = Sin(j / BigSwatch.Height * 3.14159)
            's = s ^ 2
            bright = 1 - j / BigSwatch.Height
            'If j > BigSwatch.Height / 2 Then
                R = (R * s + bright * (1 - s)) '/ s 'CDbl(BigSwatch.Height)
                g = (g * s + bright * (1 - s)) '/ s 'CDbl(BigSwatch.Height)
                b = (b * s + bright * (1 - s)) '/ s 'CDbl(BigSwatch.Height)
                'g = g * bright
                'b = b * bright
            'Else
            '    R = R * (1 - bright)
            '    g = g * (1 - bright)
            '    b = b * (1 - bright)
            '    If R > 1 Then R = 1
            '    If g > 1 Then g = 1
            '    If b > 1 Then b = 1
            'End If
            BigSwatch.PSet (i, j), convertColorRGB_selector(R, g, b)
        Next j
    Next i
    BigSwatch.Visible = True
    BigSwatch.AutoRedraw = False
End Sub


Private Sub tR_Change()
    If tR >= 0 And tR <= 255 Then SliderR = tR
    selectedColor.BackColor = "&H" & Hex2D(tB) & Hex2D(tG) & Hex2D(tR)
    colorChange
End Sub


