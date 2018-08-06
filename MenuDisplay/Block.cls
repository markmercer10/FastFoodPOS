VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PanelBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'backing fields
Private mvarBackground As Long
Private mvarBackground2 As Long
Private mvarBackgroundGradient As Boolean
Private mvarContentType As Byte
Private mvarMargin As Long
Private mvarSectionID As Long
Private mvarHFontName As String
Private mvarHFontSize As Long
Private mvarHFontBold As Boolean
Private mvarHFontColor As Long
Private mvarHStroke As Boolean
Private mvarHStrokeColor As Long
Private mvarHBack As Boolean
Private mvarHBackColor1 As Long
Private mvarHBackColor2 As Long
Private mvarIFontName As String
Private mvarIFontSize As Long
Private mvarIFontBold As Boolean
Private mvarIFontColor As Long
Private mvarPicturePath As String
Private mvarStretch As Boolean
Private mvarTText As String
Private mvarTFontName As String
Private mvarTFontSize As Long
Private mvarTFontBold As Boolean
Private mvarTFontColor As Long

Private Declare Function SetRect Lib "user32" _
(lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Private Sub Class_Initialize()
    Let mvarBackground = vbBlack
    Let mvarBackground2 = vbBlack
    Let mvarBackgroundGradient = False
    Let mvarContentType = 0
    Let mvarMargin = 20
    Let mvarHFontName = "Arial"
    Let mvarHFontSize = 50
    Let mvarHFontBold = False
    Let mvarHFontColor = vbWhite
    Let mvarHStroke = False
    Let mvarHStrokeColor = vbBlack
    Let mvarHBack = False
    Let mvarHBackColor1 = vbCyan
    Let mvarHBackColor2 = vbBlue
    Let mvarIFontName = "Arial"
    Let mvarIFontSize = 30
    Let mvarIFontBold = False
    Let mvarIFontColor = vbWhite
    Let mvarPicturePath = ""
    Let mvarStretch = True
    Let mvarTText = "Text"
    Let mvarTFontName = "Arial"
    Let mvarTFontSize = 60
    Let mvarTFontBold = True
    Let mvarTFontColor = vbYellow
End Sub

Public Property Let Background(ByVal vData As Long)
    mvarBackground = vData
End Property

Public Property Get Background() As Long
    Background = mvarBackground
End Property

Public Property Let Background2(ByVal vData As Long)
    mvarBackground2 = vData
End Property

Public Property Get Background2() As Long
    Background2 = mvarBackground2
End Property

Public Property Let BackgroundGradient(ByVal vData As Boolean)
    mvarBackgroundGradient = vData
End Property

Public Property Get BackgroundGradient() As Boolean
    BackgroundGradient = mvarBackgroundGradient
End Property

Public Property Let ContentType(ByVal vData As Byte)
    mvarContentType = vData
End Property

Public Property Get ContentType() As Byte
    ContentType = mvarContentType
End Property

Public Property Let Margin(ByVal vData As Long)
    mvarMargin = vData
End Property

Public Property Get Margin() As Long
    Margin = mvarMargin
End Property

Public Property Let SectionID(ByVal vData As Long)
    mvarSectionID = vData
End Property

Public Property Get SectionID() As Long
    SectionID = mvarSectionID
End Property

Public Property Let HFontName(ByVal vData As String)
    mvarHFontName = vData
End Property

Public Property Get HFontName() As String
    HFontName = mvarHFontName
End Property

Public Property Let hFontSize(ByVal vData As Long)
    mvarHFontSize = vData
End Property

Public Property Get hFontSize() As Long
    hFontSize = mvarHFontSize
End Property

Public Property Let hFontBold(ByVal vData As Boolean)
    mvarHFontBold = vData
End Property

Public Property Get hFontBold() As Boolean
    hFontBold = mvarHFontBold
End Property

Public Property Let hFontColor(ByVal vData As Long)
    mvarHFontColor = vData
End Property

Public Property Get hFontColor() As Long
    hFontColor = mvarHFontColor
End Property

Public Property Let hStroke(ByVal vData As Boolean)
    mvarHStroke = vData
End Property

Public Property Get hStroke() As Boolean
    hStroke = mvarHStroke
End Property

Public Property Let hStrokeColor(ByVal vData As Long)
    mvarHStrokeColor = vData
End Property

Public Property Get hStrokeColor() As Long
    hStrokeColor = mvarHStrokeColor
End Property

Public Property Let hBack(ByVal vData As Boolean)
    mvarHBack = vData
End Property

Public Property Get hBack() As Boolean
    hBack = mvarHBack
End Property

Public Property Let hBackColor1(ByVal vData As Long)
    mvarHBackColor1 = vData
End Property

Public Property Get hBackColor1() As Long
    hBackColor1 = mvarHBackColor1
End Property

Public Property Let hBackColor2(ByVal vData As Long)
    mvarHBackColor2 = vData
End Property

Public Property Get hBackColor2() As Long
    hBackColor2 = mvarHBackColor2
End Property

Public Property Let IFontName(ByVal vData As String)
    mvarIFontName = vData
End Property

Public Property Get IFontName() As String
    IFontName = mvarIFontName
End Property

Public Property Let iFontSize(ByVal vData As Long)
    mvarIFontSize = vData
End Property

Public Property Get iFontSize() As Long
    iFontSize = mvarIFontSize
End Property

Public Property Let iFontBold(ByVal vData As Boolean)
    mvarIFontBold = vData
End Property

Public Property Get iFontBold() As Boolean
    iFontBold = mvarIFontBold
End Property

Public Property Let iFontColor(ByVal vData As Long)
    mvarIFontColor = vData
End Property

Public Property Get iFontColor() As Long
    iFontColor = mvarIFontColor
End Property

Public Property Let PicturePath(ByVal vData As String)
    mvarPicturePath = vData
End Property

Public Property Get PicturePath() As String
    PicturePath = mvarPicturePath
End Property

Public Property Let Stretch(ByVal vData As Boolean)
    mvarStretch = vData
End Property

Public Property Get Stretch() As Boolean
    Stretch = mvarStretch
End Property

Public Property Let tText(ByVal vData As String)
    mvarTText = vData
End Property

Public Property Get tText() As String
    tText = mvarTText
End Property

Public Property Let tFontName(ByVal vData As String)
    mvarTFontName = vData
End Property

Public Property Get tFontName() As String
    tFontName = mvarTFontName
End Property

Public Property Let tFontSize(ByVal vData As Long)
    mvarTFontSize = vData
End Property

Public Property Get tFontSize() As Long
    tFontSize = mvarTFontSize
End Property

Public Property Let tFontBold(ByVal vData As Boolean)
    mvarTFontBold = vData
End Property

Public Property Get tFontBold() As Boolean
    tFontBold = mvarTFontBold
End Property

Public Property Let tFontColor(ByVal vData As Long)
    mvarTFontColor = vData
End Property

Public Property Get tFontColor() As Long
    tFontColor = mvarTFontColor
End Property


Public Sub DrawBlock(ByRef pb As PictureBox, ByVal d_scale As Double)
    Dim s As Long
    Dim h As ADODB.Recordset
    Dim i As ADODB.Recordset
    Dim lineHeight As Long
    Dim CurrentY As Long
    Dim Left As Long
    Dim hasPrice As Boolean
    Dim hasSmall As Boolean
    Dim hasMedium As Boolean
    Dim hasLarge As Boolean
    Dim columnCount As Byte
    Dim columnHeaders() As String
    Dim columnBoundaries() As Long
    Dim pricesMaxWidth As Long
    Dim longestText As Long
    Dim columnWidth As Long
    Dim col As Long
    Dim showColumnHeaders As Boolean
    Dim price As String
    Dim gradBottom As Long
    Dim tlColor As OLE_COLOR
    Dim brColor As OLE_COLOR
    
    columnCount = 0
    Left = Margin * d_scale * Screen.TwipsPerPixelX
    lineHeight = iFontSize * 2 * d_scale * Screen.TwipsPerPixelY
    CurrentY = Margin * 0.8 * d_scale * Screen.TwipsPerPixelY
    
    If IsNull(pb) Then Exit Sub
    If pb Is Nothing Then Exit Sub
    pb.Cls
    
    If BackgroundGradient Then
        PaintGradientH pb, Background, Background2
    Else
        pb.BackColor = Background
    End If
    
    If Not CheckDBConnection Then Exit Sub
    If ContentType = 0 Then
        s = SectionID
        Set h = db.Execute("SELECT * FROM menu_sections WHERE id = " & s)
        If Not (h.EOF And h.BOF) Then
            h.MoveFirst
            
            pb.Font = HFontName
            pb.FontSize = hFontSize * d_scale
            pb.FontBold = -CLng(hFontBold)
            
            If hBack Then
                tlColor = Lighten(hBackColor1)
                brColor = Darken(hBackColor2)
                gradBottom = CurrentY * 2 + pb.TextHeight("A")
                PaintGradientH pb, hBackColor1, hBackColor2, 0, gradBottom
                If BackgroundGradient Then PaintGradientH pb, Background, Background2, gradBottom
                pb.Line (0, 0)-(pb.Width, gradBottom), brColor, B
                pb.Line (0, 0)-(pb.Width - 15, gradBottom), tlColor, B
                pb.Line (15, 15)-(pb.Width - 15, gradBottom), brColor, B
                pb.Line (15, 15)-(pb.Width - 30, gradBottom - 15), tlColor, B
                pb.Line (30, 30)-(pb.Width - 30, gradBottom - 15), brColor, B
                pb.Line (30, 30)-(pb.Width - 45, gradBottom - 30), tlColor, B
                pb.Line (45, gradBottom - 30)-(pb.Width - 45, gradBottom - 30), brColor
                pb.Line -(pb.Width - 45, 45), brColor
            End If
            
            pb.CurrentX = Left
            pb.CurrentY = CurrentY
            pb.ForeColor = hFontColor
            
            If hStroke Then
                PrintTextStroke pb, h!Title, pb.CurrentX, pb.CurrentY, hFontColor, hStrokeColor, 15
            Else
                pb.Print h!Title
            End If
            
            CurrentY = CurrentY + (hFontSize + Margin * 0.2) * 2 * d_scale * Screen.TwipsPerPixelY
            
            Set i = db.Execute("SELECT * FROM menu_items WHERE section_id = " & s & " ORDER BY sort_id ASC")
            With i
                If Not (.EOF And .BOF) Then
                    longestText = 0
                    Do Until .EOF
                        pb.Font = IFontName
                        pb.FontSize = iFontSize * d_scale
                        pb.FontBold = -CLng(iFontBold)
                        If pb.TextWidth(i!name) > longestText Then longestText = pb.TextWidth(i!name)
                        If !price <> 0 Then hasPrice = True
                        If !small <> 0 Then hasSmall = True
                        If !medium <> 0 Then hasMedium = True
                        If !large <> 0 Then hasLarge = True
                        .MoveNext
                    Loop
                    
                    If hasPrice Then columnCount = columnCount + 1
                    If hasSmall Then columnCount = columnCount + 1
                    If hasMedium Then columnCount = columnCount + 1
                    If hasLarge Then columnCount = columnCount + 1
                    pricesMaxWidth = pb.Width - Margin * Screen.TwipsPerPixelX * 2 - longestText
                    columnWidth = pricesMaxWidth / columnCount
                    
                    ReDim columnHeaders(columnCount - 1)
                    ReDim columnBoundaries(columnCount - 1)
                    
                    pb.ForeColor = iFontColor
                    columnBoundaries(columnCount - 1) = pb.ScaleWidth - Margin * d_scale * Screen.TwipsPerPixelX
                    For col = columnCount - 2 To 0 Step -1
                        columnBoundaries(col) = columnBoundaries(col + 1) - columnWidth
                    Next col
                    col = columnCount - 1
                    If hasLarge Then
                        columnHeaders(col) = "Large"
                        col = col - 1
                    End If
                    If hasMedium Then
                        columnHeaders(col) = "Medium"
                        col = col - 1
                    End If
                    If hasSmall Then
                        columnHeaders(col) = "Small"
                        col = col - 1
                    End If
                    If hasPrice Then
                        columnHeaders(col) = "Price"
                        col = col - 1
                    End If
                    
                    If Not (columnCount = 1 And columnHeaders(0) = "Price") Then
                        For col = 0 To columnCount - 1
                            pb.CurrentX = columnBoundaries(col) - pb.TextWidth(columnHeaders(col))
                            pb.CurrentY = CurrentY
                            pb.Print columnHeaders(col)
                        Next col
                        CurrentY = CurrentY + lineHeight
                    End If
                    
                    .MoveFirst
                    longestText = 0
                    Do Until .EOF
                        pb.CurrentX = Left
                        pb.CurrentY = CurrentY
                        pb.Font = IFontName
                        pb.FontSize = iFontSize * d_scale
                        pb.FontBold = -CLng(iFontBold)
                        pb.ForeColor = iFontColor
                        pb.Print i!name
                        
                        For col = columnCount - 1 To 0 Step -1
                            price = Trim(Str(i.fields(LCase$(columnHeaders(col)))))
                            pb.CurrentX = columnBoundaries(col) - pb.TextWidth(price)
                            pb.CurrentY = CurrentY
                            If i.fields(LCase$(columnHeaders(col))) <> 0 Then
                                pb.Print price
                            End If
                        Next col
                        
                        CurrentY = CurrentY + lineHeight
                        .MoveNext
                    Loop
                End If
            End With
        End If
    ElseIf ContentType = 1 Then
        PaintPic PicturePath, pb, Stretch
    ElseIf ContentType = 2 Then
        'render video ... or not
    ElseIf ContentType = 3 Then
        Dim text As String
        Dim rec As RECT
        
        pb.Font = tFont
        pb.FontSize = tFontSize * d_scale
        pb.FontBold = tFontBold
        pb.ForeColor = tFontColor
        
        If pb.TextWidth(tText) > pb.Width Then
            SetRect rec, Left / Screen.TwipsPerPixelX, Top / Screen.TwipsPerPixelY, (pb.Width - Left) / Screen.TwipsPerPixelX, (pb.Height - Top) / Screen.TwipsPerPixelY
            DrawTextEx pb.hDC, tText, Len(tText), rec, DT_WORDBREAK, ByVal 0&
        Else
            SetRect rec, 0, 0, pb.Width / 15, pb.Height / 15
            DrawTextEx pb.hDC, tText, Len(tText), rec, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER, ByVal 0&
        End If
    End If


End Sub

Public Sub Save(ByVal display_id As Long, ByVal block_number As Long)
    Dim q As ADODB.Recordset
    Dim id As Long
    Dim fields As Variant
    Dim values As Variant
    If Not CheckDBConnection Then Exit Sub
    Set q = db.Execute("SELECT * FROM blocks WHERE display_id = " & display_id & " AND block_number = " & block_number)
    If Not (q.EOF And q.BOF) Then
        q.MoveFirst
        id = q!id
    End If
    
    fields = Array("id", "display_id", "block_number", "content_type", "margin", "background", "background2", "gradient", "section_id", "header_font", "header_size", "header_bold", "header_color", "header_stroke", "header_stroke_color", "header_back", "header_back_color1", "header_back_color2", "item_font", "item_size", "item_bold", "item_color", "picture", "stretch", "text", "text_font", "text_size", "text_bold", "text_color")
    values = Array(id, display_id, block_number, ContentType, Margin, Background, Background2, -CLng(BackgroundGradient), SectionID, HFontName, hFontSize, -CLng(hFontBold), hFontColor, -CLng(hStroke), hStrokeColor, -CLng(hBack), hBackColor1, hBackColor2, IFontName, iFontSize, -CLng(iFontBold), iFontColor, EscapeBackslashes(PicturePath), -CLng(Stretch), tText, tFontName, tFontSize, -CLng(tFontBold), tFontColor)
    Upsert "blocks", fields, values
End Sub

Public Sub Load(ByVal display_id As Long, ByVal block_number As Long)
    Dim q As ADODB.Recordset
    Dim s As ADODB.Recordset
    If Not CheckDBConnection Then Exit Sub
    Set q = db.Execute("SELECT * FROM blocks WHERE display_id = " & display_id & " AND block_number = " & block_number)
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Let mvarContentType = !content_type
            Let mvarBackground = !Background
            Let mvarBackground2 = val("" & !Background2)
            Let mvarBackgroundGradient = CBool(!Gradient)
            Let mvarPicturePath = !Picture
            Let mvarStretch = CBool(!Stretch)
            
            Set s = db.Execute("SELECT * FROM menu_sections WHERE id = " & !section_id)
            If Not (s.EOF And s.BOF) Then
                Let mvarSectionID = !section_id
                Let mvarMargin = !Margin
                Let mvarHFontName = !header_font
                Let mvarHFontSize = !header_size
                Let mvarHFontBold = CBool(!header_bold)
                Let mvarHFontColor = !header_color
                Let mvarHStroke = CBool(!header_stroke)
                Let mvarHStrokeColor = val("" & !header_stroke_color)
                Let mvarHBack = CBool(!header_back)
                Let mvarHBackColor1 = val("" & !header_back_color1)
                Let mvarHBackColor2 = val("" & !header_back_color2)
                Let mvarIFontName = !item_font
                Let mvarIFontSize = !item_size
                Let mvarIFontBold = CBool(!item_bold)
                Let mvarIFontColor = !item_color
                Let mvarTText = "" & !text
                Let mvarTFontName = "" & !text_font
                Let mvarTFontSize = val("" & !text_size)
                Let mvarTFontBold = CBool(!text_bold)
                Let mvarTFontColor = val("" & !text_color)
            Else
                Let mvarSectionID = -1
            End If
        End If
    End With
End Sub
