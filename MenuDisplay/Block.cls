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
Private mvarContentType As Byte
Private mvarMargin As Long
Private mvarSectionID As Long
Private mvarHFontName As String
Private mvarHFontSize As Long
Private mvarHFontBold As Boolean
Private mvarHFontColor As Long
Private mvarIFontName As String
Private mvarIFontSize As Long
Private mvarIFontBold As Boolean
Private mvarIFontColor As Long
Private mvarPicturePath As String
Private mvarStretch As Boolean

Private Sub Class_Initialize()
    Let mvarBackground = vbBlack
    Let mvarContentType = 0
    Let mvarMargin = 20
    Let mvarHFontName = "Arial"
    Let mvarHFontSize = 50
    Let mvarHFontBold = False
    Let mvarHFontColor = vbWhite
    Let mvarIFontName = "Arial"
    Let mvarIFontSize = 30
    Let mvarIFontBold = False
    Let mvarIFontColor = vbWhite
    Let mvarPicturePath = ""
    Let mvarStretch = True
End Sub

Public Property Let Background(ByVal vData As Long)
    mvarBackground = vData
End Property

Public Property Get Background() As Long
    Background = mvarBackground
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

Public Sub DrawBlock(ByRef pb As PictureBox, ByVal d_scale As Double)
    Dim s As Long
    Dim h As ADODB.Recordset
    Dim i As ADODB.Recordset
    Dim lineHeight As Long
    Dim CurrentY As Long
    Dim left As Long
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
    
    columnCount = 0
    left = Margin * d_scale * Screen.TwipsPerPixelX
    lineHeight = iFontSize * 2 * d_scale * Screen.TwipsPerPixelY
    CurrentY = Margin * d_scale * Screen.TwipsPerPixelY
    
    If IsNull(pb) Then Exit Sub
    If pb Is Nothing Then Exit Sub
    pb.Cls
    pb.BackColor = Background
    If ContentType = 0 Then
        s = SectionID
        Set h = db.Execute("SELECT * FROM menu_sections WHERE id = " & s)
        If Not (h.EOF And h.BOF) Then
            h.MoveFirst
            
            pb.CurrentX = left
            pb.CurrentY = CurrentY
            pb.Font = HFontName
            pb.FontSize = hFontSize * d_scale
            pb.FontBold = -CLng(hFontBold)
            pb.ForeColor = hFontColor
            pb.Print h!Title
            CurrentY = CurrentY + hFontSize * 2 * d_scale * Screen.TwipsPerPixelY
            
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
                        pb.CurrentX = left
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
    End If


End Sub

Public Sub Save(ByVal display_id As Long, ByVal block_number As Long)
    Dim q As ADODB.Recordset
    Dim id As Long
    Dim fields As Variant
    Dim values As Variant
    Set q = db.Execute("SELECT * FROM blocks WHERE display_id = " & display_id & " AND block_number = " & block_number)
    If Not (q.EOF And q.BOF) Then
        q.MoveFirst
        id = q!id
    End If
    
    fields = Array("id", "display_id", "block_number", "content_type", "margin", "background", "section_id", "header_font", "header_size", "header_bold", "header_color", "item_font", "item_size", "item_bold", "item_color", "picture", "stretch")
    values = Array(id, display_id, block_number, ContentType, Margin, Background, SectionID, HFontName, hFontSize, -CLng(hFontBold), hFontColor, IFontName, iFontSize, -CLng(iFontBold), iFontColor, EscapeBackslashes(PicturePath), -CLng(Stretch))
    Upsert "blocks", fields, values
End Sub

Public Sub Load(ByVal display_id As Long, ByVal block_number As Long)
    Dim q As ADODB.Recordset
    Dim s As ADODB.Recordset
    Set q = db.Execute("SELECT * FROM blocks WHERE display_id = " & display_id & " AND block_number = " & block_number)
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Let mvarContentType = !content_type
            Let mvarBackground = !Background
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
                Let mvarIFontName = !item_font
                Let mvarIFontSize = !item_size
                Let mvarIFontBold = CBool(!item_bold)
                Let mvarIFontColor = !item_color
            Else
                Let mvarSectionID = -1
            End If
        End If
    End With
End Sub
