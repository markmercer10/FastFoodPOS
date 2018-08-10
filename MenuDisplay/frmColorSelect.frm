VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColorSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Color Selector"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   304
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   231
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   240
      Top             =   240
   End
   Begin VB.PictureBox BigSwatch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   119
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   215
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.PictureBox inverseColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2280
      ScaleHeight     =   705
      ScaleWidth      =   1065
      TabIndex        =   19
      ToolTipText     =   "The Inverse of the selected color"
      Top             =   1920
      Width           =   1095
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Inverse"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.TextBox vbHex 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "&HFFFFFF"
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox htmlHex 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "FFFFFF"
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton okButn 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox tB 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   600
      TabIndex        =   11
      Text            =   "0"
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox tG 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   600
      TabIndex        =   8
      Text            =   "0"
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox selectedColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   1065
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox nearestColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1200
      ScaleHeight     =   705
      ScaleWidth      =   1065
      TabIndex        =   4
      ToolTipText     =   "The nearest web safe color"
      Top             =   1920
      Width           =   1095
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Web Safe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.TextBox tR 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Text            =   "0"
      Top             =   2760
      Width           =   495
   End
   Begin MSComctlLib.Slider SliderR 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   2760
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      LargeChange     =   25
      Max             =   255
      TickStyle       =   3
      TickFrequency   =   10
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider SliderG 
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      Top             =   3120
      Width           =   2415
      _ExtentX        =   4260
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
      Left            =   1080
      TabIndex        =   9
      Top             =   3480
      Width           =   2415
      _ExtentX        =   4260
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
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "HTML Hex : "
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "VB Hex"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Blue"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Green"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Red"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   615
   End
End
Attribute VB_Name = "frmColorSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MouseIsDown As Boolean

Sub colorChange()
    Dim col As Long
    Dim nR As Long
    Dim nG As Long
    Dim nB As Long
    Dim avg As Long
    Dim similar1 As Boolean
    Dim similar2 As Boolean
    col = selectedColor.BackColor
    Shape1.FillColor = col
    tR = Int(col And &HFF)
    tG = Int(Int(col / 256#) And &HFF)
    tB = Int(Int(col / 65536#) And &HFF)
    htmlHex = Hex2D(tR) & Hex2D(tG) & Hex2D(tB)
    vbHex = "&H" & Hex2D(tB) & Hex2D(tG) & Hex2D(tR)
    
    inverseColor.BackColor = "&H" & Hex2D(255 - tB) & Hex2D(255 - tG) & Hex2D(255 - tR)
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
    
    avg = (val(tR) + val(tG) + val(tB)) / 3
    similar1 = (((tR > 180 And tB > 200) Or (val(tR) + val(tB) > 360)) And tG < 140)
    similar2 = ((((255 - tR) > 180 And (255 - tB) > 200) Or ((255 - val(tR)) + (255 - val(tR)) > 360)) And (255 - tG) < 140)
    'MsgBox avg
    If (avg >= 100 And avg <= 155) Or similar1 Or similar2 Then
        Label4.ForeColor = vbBlack
        Label5.ForeColor = vbBlack
        Label7.ForeColor = vbBlack
    Else
        Label4.ForeColor = inverseColor.BackColor
        Label5.ForeColor = inverseColor.BackColor
        Label7.ForeColor = selectedColor.BackColor
    End If
End Sub

Private Sub Command1_Click()
End Sub

Function convertColorRGB_selector(ByVal r As Double, ByVal g As Double, ByVal b As Double) As Long
    convertColorRGB_selector = (CLng(r * &HFF))
    convertColorRGB_selector = convertColorRGB_selector + (CLng(g * &HFF)) * 256
    convertColorRGB_selector = convertColorRGB_selector + (CLng(b * &HFF)) * 65536
End Function



Function Hex2D(ByVal n As Long) As String
    Hex2D = Hex$(n)
    If Len(Hex2D) < 2 Then Hex2D = "0" & Hex2D
End Function


Private Sub BigSwatch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseIsDown = True
    selectedColor.BackColor = BigSwatch.Point(X, Y)
    colorChange
End Sub

Private Sub BigSwatch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MouseIsDown Then
        If X >= 0 And Y >= 0 And X < BigSwatch.ScaleWidth And Y < BigSwatch.ScaleHeight Then
            selectedColor.BackColor = BigSwatch.Point(X, Y)
            colorChange
        End If
    End If
End Sub

Private Sub BigSwatch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseIsDown = False
End Sub

Private Sub inverseColor_Click()
    selectedColor.BackColor = inverseColor.BackColor
    colorChange
End Sub

Private Sub Label5_Click()
    nearestColor_Click
End Sub

Private Sub Label7_Click()
    inverseColor_Click
End Sub

Private Sub nearestColor_Click()
    selectedColor.BackColor = nearestColor.BackColor
    colorChange
End Sub


Private Sub okButn_Click()
    csSelectedColor = selectedColor.BackColor
    Unload Me
End Sub


Private Sub SliderB_Change()
    SliderB_Click
End Sub

Private Sub SliderB_Click()
    tB = SliderB.value
End Sub

Private Sub SliderG_Change()
    SliderG_Click
End Sub

Private Sub SliderG_Click()
    tG = SliderG.value
End Sub


Private Sub SliderR_Change()
    SliderR_Click
End Sub

Private Sub SliderR_Click()
    tR = SliderR.value
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
    Dim r0 As Double
    Dim g0 As Double
    Dim b0 As Double
    Dim r As Double
    Dim g As Double
    Dim b As Double
    Dim s As Double
    Dim bigSwatchWidth As Long
    Dim brightenDarken As Double
    Timer1.Enabled = False
    BigSwatch.Visible = False
    BigSwatch.AutoRedraw = True
    bigSwatchWidth = BigSwatch.Width - 14
    DoEvents
    
    selectedColor.BackColor = val(Me.Tag)
    colorChange
    
    For i = 0 To bigSwatchWidth
        r0 = Sin(((i / bigSwatchWidth) + (3 / 12#)) * 2# * 3.14159) + 0.5
        If r0 < 0 Then r0 = 0
        If r0 > 1 Then r0 = 1
        g0 = Sin(((i / bigSwatchWidth) - (1 / 12#)) * 2# * 3.14159) + 0.5
        If g0 < 0 Then g0 = 0
        If g0 > 1 Then g0 = 1
        b0 = Sin(((i / bigSwatchWidth) - (5 / 12#)) * 2# * 3.14159) + 0.5
        If b0 < 0 Then b0 = 0
        If b0 > 1 Then b0 = 1
        For j = 0 To BigSwatch.Height
            r = r0
            g = g0
            b = b0
            
            'old method (trigonometric)
            s = (Sin(j / BigSwatch.Height * 3.14159) ^ 0.5) * (Sin(j / BigSwatch.Height * 3.14159) ^ 2)
            brightenDarken = 1 - j / BigSwatch.Height
            r = (r * s + brightenDarken * (1 - s))
            g = (g * s + brightenDarken * (1 - s))
            b = (b * s + brightenDarken * (1 - s))
            
            'combination of old and new method (linear + weighted average)
            brightenDarken = 1 - j / BigSwatch.Height * 2
            r = (r * 3 + Clamp(r0 + brightenDarken, 0, 1)) / 4
            g = (g * 3 + Clamp(g0 + brightenDarken, 0, 1)) / 4
            b = (b * 3 + Clamp(b0 + brightenDarken, 0, 1)) / 4
            
            BigSwatch.PSet (i, j), convertColorRGB_selector(r, g, b)
        Next j
    Next i
    For i = bigSwatchWidth + 1 To BigSwatch.Width
        For j = 0 To BigSwatch.Height
            r = 1 - j / BigSwatch.Height
            g = r
            b = r
            BigSwatch.PSet (i, j), convertColorRGB_selector(r, g, b)
        Next j
    Next i
    BigSwatch.Visible = True
    BigSwatch.AutoRedraw = False
End Sub

Private Function Clamp(ByVal number As Double, ByVal min As Double, ByVal max As Double)
    If number < min Then
        Clamp = min
    ElseIf number > max Then
        Clamp = max
    Else
        Clamp = number
    End If
End Function

Private Sub tR_Change()
    If tR >= 0 And tR <= 255 Then SliderR = tR
    selectedColor.BackColor = "&H" & Hex2D(tB) & Hex2D(tG) & Hex2D(tR)
    colorChange
End Sub


