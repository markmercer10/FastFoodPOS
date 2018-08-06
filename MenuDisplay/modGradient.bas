Attribute VB_Name = "modGradient"
'The following code snippet paints a gradient between 2 colors on the entire form.
'It uses Win32 API in order to get the best performances.
'You can change the colors of the gradient by changing the RGB values
'passed to PaintGradient function.
'
'Written by Nir Sofer
'Web site: http://nirsoft.mirrorz.com

Private Declare Function GetClientRect Lib "user32" _
(ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" _
(ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" _
(ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" _
(ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" _
(ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" _
(lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Private Function SafeDiv(X1 As Double, X2 As Double) As Double
    If X2 = 0 Then SafeDiv = 0 Else SafeDiv = X1 / X2
End Function

Public Sub PaintGradientH(obj As Object, Color1 As OLE_COLOR, Color2 As OLE_COLOR, Optional Top As Long = -1, Optional Bottom As Long = -1)
    PaintGradient obj, Color1 And 255, Color1 \ 256 And 255, Color1 \ 65536 And 255, Color2 And 255, Color2 \ 256 And 255, Color2 \ 65536 And 255, Top, Bottom
End Sub

Public Sub PaintGradient(obj As Object, Red1 As Integer, Green1 As Integer, Blue1 As Integer, _
Red2 As Integer, Green2 As Integer, Blue2 As Integer, Optional Top As Long = -1, Optional Bottom As Long = -1)
    Dim WinRect     As RECT
    Dim ColorRect   As RECT
    Dim Y           As Long
    Dim hBrush      As Long
    Dim hPrevBrush  As Long
    Dim DivValue    As Double
    Dim CurrRed     As Integer
    Dim CurrGreen   As Integer
    Dim CurrBlue    As Integer
    
    GetClientRect obj.hwnd, WinRect
    If obj.ScaleMode = vbPixels Then
        If Top > -1 Then WinRect.Top = Top
        If Bottom > -1 Then WinRect.Bottom = Bottom
    ElseIf obj.ScaleMode = vbTwips Then
        If Top > -1 Then WinRect.Top = Top / Screen.TwipsPerPixelY
        If Bottom > -1 Then WinRect.Bottom = Bottom / Screen.TwipsPerPixelY
    End If
    
    For Y = WinRect.Top To WinRect.Bottom
        DivValue = SafeDiv((WinRect.Bottom - WinRect.Top), (Y - WinRect.Top))
        CurrRed = Red1 + SafeDiv((Red2 - Red1), DivValue)
        CurrGreen = Green1 + SafeDiv((Green2 - Green1), DivValue)
        CurrBlue = Blue1 + SafeDiv((Blue2 - Blue1), DivValue)
        SetRect ColorRect, WinRect.Left, Y, WinRect.Right, Y + 1
        hBrush = CreateSolidBrush(RGB(CurrRed, CurrGreen, CurrBlue))
        hPrevBrush = SelectObject(obj.hDC, hBrush)
        FillRect obj.hDC, ColorRect, hBrush
        SelectObject obj.hDC, hPrevBrush
        DeleteObject hBrush
    Next
End Sub

Public Function Lighten(ByVal Color1 As OLE_COLOR) As OLE_COLOR
    Dim r
    Dim b
    Dim g
    r = Color1 And 255
    g = Color1 \ 256 And 255
    b = Color1 \ 65536 And 255
    Lighten = RGB(Int((r + 255) / 2), Int((g + 255) / 2), Int((b + 255) / 2))
End Function

Public Function Darken(ByVal Color1 As OLE_COLOR) As OLE_COLOR
    Dim r
    Dim b
    Dim g
    r = Color1 And 255
    g = Color1 \ 256 And 255
    b = Color1 \ 65536 And 255
    Darken = RGB(Int(r / 2), Int(g / 2), Int(b / 2))
End Function

