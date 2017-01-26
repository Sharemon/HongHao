VERSION 5.00
Begin VB.UserControl Curve 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "Curve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EnumCurveStyle
    ecsCurveMid = 0 '¥”÷–º‰ø™ ºÀ„ 0 ”√”⁄”– +- ÷µ
    ecsCurveDown   '¥”µ◊≤øø™ ºÀ„ 0 ”√”⁄÷ª”– + ÷µ
    ecsCurveUp   '¥”µ◊≤øø™ ºÀ„ 0 ”√”⁄÷ª”– - ÷µ
End Enum

Private m_PictureBox As PictureBox
Private m_CurveData As Collection
Private m_CurveColor As OLE_COLOR
Private m_BorderSize As Integer
Private m_BorderColor As Long
Private m_GridVisible As Boolean
Private m_GridColor As Long
Private m_BorderStyle As Boolean
Private m_GridRange As String
Private m_colValue As Collection

Private m_MaxValue As Single
Private m_MinValue As Single
Private m_MidValue As Single
Private m_ValueRange As Single
Private m_Value As Single

Private m_CurvCount As Long
Private m_ForeColor As OLE_COLOR
Private m_ShowMidLine As Boolean
Private m_MidLineColor As OLE_COLOR
Private m_ShowScale As Boolean
Private m_AutoRange As Boolean
Private m_CurveStyle As EnumCurveStyle

'«˙œﬂ¿‡–Õ
Public Property Get CurveStyle() As EnumCurveStyle
    CurveStyle = m_CurveStyle
End Property

Public Property Let CurveStyle(ByVal bValue As EnumCurveStyle)
    m_CurveStyle = bValue
    Call DrawCurve
End Property

'◊‘∂Øµ˜’˚∑∂Œß
Public Property Get AutoRange() As Boolean
    AutoRange = m_AutoRange
End Property

Public Property Let AutoRange(ByVal bValue As Boolean)
    m_AutoRange = bValue
    Call DrawCurve
End Property

'«˙œﬂøÃ∂»∑∂Œß
Public Property Get ValueRange() As Single
    ValueRange = m_ValueRange
End Property

Public Property Let ValueRange(ByVal LonValue As Single)
    m_ValueRange = Abs(LonValue)
    Call DrawCurve
End Property

'÷–º‰÷µ
Public Property Get MidValue() As Single
    MidValue = m_MidValue
End Property

Public Property Let MidValue(ByVal LonValue As Single)
    m_MidValue = LonValue
    Call DrawCurve
End Property

'œ‘ æøÃ∂»
Public Property Get ShowScale() As Boolean
    ShowScale = m_ShowScale
End Property

Public Property Let ShowScale(ByVal bValue As Boolean)
    m_ShowScale = bValue
    Call DrawCurve
End Property

'÷–º‰œﬂ—’…´
Public Property Get MidLineColor() As OLE_COLOR
    MidLineColor = m_MidLineColor
End Property

Public Property Let MidLineColor(ByVal OleValue As OLE_COLOR)
    m_MidLineColor = OleValue
    Call DrawCurve
End Property

'œ‘ æ÷–º‰œﬂ
Public Property Get ShowMidLine() As Boolean
    ShowMidLine = m_ShowMidLine
End Property

Public Property Let ShowMidLine(ByVal bValue As Boolean)
    m_ShowMidLine = bValue
    Call DrawCurve
End Property

'◊÷ÃÂ—’…´
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_PictureBox.ForeColor
End Property

Public Property Let ForeColor(ByVal OleValue As OLE_COLOR)
    m_PictureBox.ForeColor = OleValue
    Call DrawCurve
End Property

'“ª√Êœ‘ æµƒ◊Ó ˝¡ø÷µ
Public Property Get CurvCount() As Long
    CurvCount = m_CurvCount
End Property

Public Property Let CurvCount(ByVal LonValue As Long)
    m_CurvCount = LonValue
    Call DrawCurve
End Property

'◊Ó¥Û÷µ
Public Property Get MaxValue() As Single
    MaxValue = m_MaxValue
End Property

'Public Property Let MaxValue(ByVal LonValue As Single)
'    m_MaxValue = LonValue
'End Property

'◊Ó–°÷µ
Public Property Get MinValue() As Single
    MinValue = m_MinValue
End Property

'Public Property Let MinValue(ByVal LonValue As Single)
'    m_MinValue = LonValue
'End Propertyˇ

'Õ¯∏Ò∑∂Œß
Public Property Get GridRange() As String
    GridRange = m_GridRange
End Property

Public Property Let GridRange(ByVal StrValue As String)
    Dim p() As String
    
    p = Split(StrValue, ",")
    If Not (UBound(p) = 1 And IsNumeric(p(0)) And IsNumeric(p(1))) Then Exit Property
    
    StrValue = Val(p(0)) & "," & Val(p(1))
    m_GridRange = StrValue
    Call DrawCurve
End Property

'Õ¯¬Á—’…´
Public Property Get GridColor() As OLE_COLOR
    GridColor = m_GridColor
End Property

Public Property Let GridColor(ByVal vData As OLE_COLOR)
    m_GridColor = vData
    Call DrawCurve
End Property

' «∑Òœ‘ æœﬂÃı
Public Property Get GridVisible() As Boolean
    GridVisible = m_GridVisible
End Property

Public Property Let GridVisible(ByVal vData As Boolean)
    m_GridVisible = vData
    Call DrawCurve
End Property

'±ﬂøÚ—’…´
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal vData As OLE_COLOR)
    m_BorderColor = vData
    Call DrawCurve
End Property

'±ﬂøÚ¥Û–°
Public Property Get BorderSize() As Long
    BorderSize = m_BorderSize
End Property

Public Property Let BorderSize(ByVal vData As Long)
    m_BorderSize = vData
    Call DrawCurve
End Property

'«˙œﬂ—’…´
Public Property Get CurveColor() As OLE_COLOR
    CurveColor = m_CurveColor
End Property

Public Property Let CurveColor(ByVal vData As OLE_COLOR)
    Let m_CurveColor = vData
End Property

'±≥æ∞—’…´
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_PictureBox.BackColor
End Property

Public Property Let BackColor(ByVal vData As OLE_COLOR)
    m_PictureBox.BackColor = vData
    Call DrawCurve
End Property

'«˙œﬂ ˝æ›
Public Property Get CurveData() As Collection
    Set CurveData = m_CurveData
End Property

Public Property Set CurveData(ByVal vData As Collection)
    Set m_CurveData = vData
End Property

'ª≠«˙œﬂ
Public Sub DrawCurve()
    Dim BDR As Integer, X As Integer
    Dim NewX As Double, NewY As Double
    Dim OldX As Double, OldY As Double
    Dim GridHeight As Double, GridWidth As Double
    Dim p() As String

    With m_PictureBox
        .Cls
        BDR = .BorderStyle

        '========== ª≠±ﬂøÚ
        If m_BorderSize > 0 Then
            For X = 0 To m_BorderSize
                m_PictureBox.Line (X, X)-(.ScaleWidth - (BDR + X), .ScaleHeight - (BDR + X)), m_BorderColor, B
            Next X
        End If

        '========== ª≠Õ¯∏Ò
        If m_GridVisible = True Then
            p = Split(GridRange, ",")

            For X = 1 To Val(p(0))
                m_PictureBox.Line (m_BorderSize, m_BorderSize)-((((.ScaleWidth - m_BorderSize) / Val(p(0))) * X), (.ScaleHeight - m_BorderSize)), m_GridColor, B
            Next X
            For X = 1 To Val(p(1))
                m_PictureBox.Line (m_BorderSize, m_BorderSize)-((.ScaleWidth - m_BorderSize), (((.ScaleHeight - m_BorderSize) / Val(p(1))) * X)), m_GridColor, B
            Next X

            '÷–º‰œﬂ
            'If m_ShowMidLine = True And m_CurveStyle = ecsCurveMid Then
                'm_PictureBox.Line (.ScaleWidth - m_BorderSize, .ScaleHeight / 2)-(m_BorderSize, .ScaleHeight / 2), m_MidLineColor, B
            'End If
        End If

        '========== »Áπ˚¥Ê‘⁄ ˝æ›‘Úª≠«˙œﬂ
        If m_CurveData.Count > 0 Then
            GridHeight = ((.ScaleHeight - (m_BorderSize * 2)) / m_ValueRange) + 0    ' 0-100%
            GridWidth = ((.ScaleWidth - (m_BorderSize * 2)) / m_CurvCount) + 0       ' 1-100 Items
            Do
                If m_CurveData.Count > m_CurvCount Then m_CurveData.Remove 1
            Loop While m_CurveData.Count > m_CurvCount

            OldX = m_BorderSize + 2
            OldY = ((.ScaleHeight - (m_BorderSize * 2)) - (m_CurveData(1) * IIf(m_CurveStyle = ecsCurveUp, -GridHeight, GridHeight)))
            If m_CurveStyle = ecsCurveMid Then OldY = OldY / 2
            
            For X = 1 To m_CurveData.Count
                NewX = (.ScaleWidth - (m_BorderSize * 2)) - ((m_CurvCount - (X - 1)) * GridWidth)
                NewY = ((.ScaleHeight - (m_BorderSize * 2)) - (m_CurveData(X) * IIf(m_CurveStyle = ecsCurveUp, -GridHeight, GridHeight)))
                If m_CurveStyle = ecsCurveMid Then NewY = NewY / 2
                
                NewX = NewX + 2
                If NewX < m_BorderSize Then NewX = m_BorderSize
                If NewY < m_BorderSize Then NewY = m_BorderSize

                m_PictureBox.Line (OldX, OldY)-(NewX, NewY), m_CurveColor
                OldX = NewX
                OldY = NewY
                If OldX < m_BorderSize Then OldX = m_BorderSize
                If OldY < m_BorderSize Then OldY = m_BorderSize
            Next
        End If

        '========= ª≠øÃ∂»
        If m_ShowScale = True Then DrawText
    End With
End Sub

'ª≠Œƒ◊÷
Private Sub DrawText()
    Dim k As Long, k1 As Long
    'm_ValueRange = 100
    With m_PictureBox
        k = .ScaleHeight / 2
        If m_CurveStyle = ecsCurveMid Then
            '===== ◊Ó¥Û÷µ
            .CurrentX = m_BorderSize + 3
            .CurrentY = m_BorderSize
            m_PictureBox.Print FormatNumber(m_MidValue + m_ValueRange, 2, vbTrue)

            '===== ÷–…œ÷µ
            .CurrentX = m_BorderSize + 3
            .CurrentY = k / 2
            m_PictureBox.Print FormatNumber(m_MidValue + m_ValueRange / 2, 2, vbTrue)

            '===== ÷–÷µ
            .CurrentX = m_BorderSize + 3
            .CurrentY = k - 6
            m_PictureBox.Print FormatNumber(m_MidValue, 2, vbTrue)

            '===== ÷–œ¬÷µ
            .CurrentX = m_BorderSize + 3
            .CurrentY = k + k / 2
            m_PictureBox.Print FormatNumber(m_MidValue - m_ValueRange / 2, 2, vbTrue)

            '===== ◊Ó–°÷µ
            .CurrentX = m_BorderSize + 3
            .CurrentY = .ScaleHeight - 13 - m_BorderSize
            m_PictureBox.Print FormatNumber(m_MidValue - m_ValueRange, 2, vbTrue)

        ElseIf m_CurveStyle = ecsCurveDown Or m_CurveStyle = ecsCurveUp Then
            k1 = m_ValueRange / 4
            If m_CurveStyle = ecsCurveUp Then k1 = -k1
            
            '===== …œ÷µ
            .CurrentX = m_BorderSize + 3
            .CurrentY = m_BorderSize
            m_PictureBox.Print m_MidValue + k1 * 4

            '===== …œ…œ÷µ
            .CurrentX = m_BorderSize + 3
            .CurrentY = k / 2
            m_PictureBox.Print m_MidValue + k1 * 3

            '===== ÷–…œ÷µ
            .CurrentX = m_BorderSize + 3
            .CurrentY = k - 6
            m_PictureBox.Print m_MidValue + k1 * 2

            '===== ÷–÷–÷µ
            .CurrentX = m_BorderSize + 3
            .CurrentY = k + k / 2
            m_PictureBox.Print m_MidValue + k1

            '===== ÷–÷µ
            .CurrentX = m_BorderSize + 3
            .CurrentY = .ScaleHeight - 13 - m_BorderSize
            m_PictureBox.Print m_MidValue
        End If

        '===== µ±«∞÷µ
        .CurrentX = .ScaleWidth - 125
        .CurrentY = .ScaleHeight - 18
        'm_PictureBox.Print "µ±«∞÷µ£∫" & Format(m_Value / 100, "0.00%")
        m_PictureBox.Print "µ±«∞÷µ£∫" & Format(m_Value, "0.0000000")
    End With
End Sub

Public Sub Clear()
    Set m_CurveData = New Collection
    m_MaxValue = 0
    m_MinValue = 0
    m_Value = 0
    Call DrawCurve
End Sub

'µ±«∞÷µ
Public Property Let Value(ByVal vData As Single)
    If Ambient.UserMode = False Then Exit Property
    
    m_Value = vData
    
    If m_Value > m_MaxValue Then
        m_MaxValue = m_Value
        If m_Value > m_ValueRange And m_AutoRange = True Then
            m_ValueRange = m_Value
            If m_Value Mod 2 <> 0 Then m_ValueRange = m_ValueRange + 1
        End If
        
    ElseIf m_Value < m_MinValue Then
        m_MinValue = m_Value
        
        If m_Value < m_ValueRange And m_AutoRange = True Then
            m_ValueRange = Abs(m_Value)
            If m_Value Mod 2 <> 0 Then m_ValueRange = m_ValueRange - 1
        End If
    End If
    
    m_CurveData.Add m_Value
    Call DrawCurve
End Property

Public Property Get Value() As Single
    Value = m_Value
End Property

'≥ı ºªØ
Private Sub UserControl_Initialize()
    Set m_PictureBox = Controls.Add("VB.PictureBox", "PictureBox1")
    With m_PictureBox
        .Visible = True
        .AutoRedraw = True
        .ScaleMode = 3
        .BorderStyle = 0
        .ForeColor = &HE5E5E5
        .BackColor = vbBlack
        .FontBold = True
    End With
    
    Set CurveData = New Collection
    
    m_GridRange = "20,20"
    m_GridColor = &H808000
    m_GridVisible = True
    m_BorderSize = 0
    m_BorderColor = vbRed
    m_CurveColor = vbGreen
    m_MidLineColor = m_CurveColor
    m_MidValue = 0
    m_CurvCount = 100
    m_ValueRange = 100
    m_ShowMidLine = True
    m_ShowScale = True
    m_AutoRange = True
    m_CurveStyle = ecsCurveMid
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_GridRange = .ReadProperty("GridRange", m_GridRange)
        m_GridColor = .ReadProperty("GridColor", m_GridColor)
        m_GridVisible = .ReadProperty("GridVisible", m_GridVisible)
        m_BorderColor = .ReadProperty("BorderColor", m_BorderColor)
        m_BorderSize = .ReadProperty("BorderSize", m_BorderSize)
        m_CurveColor = .ReadProperty("CurveColor", m_CurveColor)
        m_MidLineColor = .ReadProperty("MidLineColor", m_MidLineColor)
        m_MaxValue = .ReadProperty("MaxValue", m_MaxValue)
        m_MinValue = .ReadProperty("MinValue", m_MinValue)
        m_CurvCount = .ReadProperty("CurvCount", m_CurvCount)
        m_PictureBox.BackColor = .ReadProperty("BackColor", m_PictureBox.BackColor)
        m_PictureBox.ForeColor = .ReadProperty("ForeColor", m_PictureBox.ForeColor)
        m_ShowMidLine = .ReadProperty("ShowMidLine", m_ShowMidLine)
        m_ShowScale = .ReadProperty("ShowScale", m_ShowScale)
        m_MidValue = .ReadProperty("MidValue", m_MidValue)
        m_ValueRange = .ReadProperty("ValueRange", m_ValueRange)
        m_AutoRange = .ReadProperty("AutoRange", m_AutoRange)
        m_CurveStyle = .ReadProperty("CurveStyle", m_CurveStyle)
    End With
    Call DrawCurve
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "GridRange", m_GridRange
        .WriteProperty "GridColor", m_GridColor
        .WriteProperty "GridVisible", m_GridVisible
        .WriteProperty "BorderColor", m_BorderColor
        .WriteProperty "BorderSize", m_BorderSize
        .WriteProperty "CurveColor", m_CurveColor
        .WriteProperty "BackColor", m_PictureBox.BackColor
        .WriteProperty "MaxValue", m_MaxValue
        .WriteProperty "MinValue", m_MinValue
        .WriteProperty "CurvCount", m_CurvCount
        .WriteProperty "ForeColor", m_PictureBox.ForeColor
        .WriteProperty "ShowMidLine", m_ShowMidLine
        .WriteProperty "MidLineColor", m_MidLineColor
        .WriteProperty "ShowScale", m_ShowScale
        .WriteProperty "MidValue", m_MidValue
        .WriteProperty "ValueRange", m_ValueRange
        .WriteProperty "AutoRange", m_AutoRange
        .WriteProperty "CurveStyle", m_CurveStyle
    End With
End Sub

'–∂‘ÿ
Private Sub UserControl_Terminate()
    Set CurveData = Nothing
    Set m_PictureBox = Nothing
End Sub

Private Sub UserControl_Resize()
    m_PictureBox.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    Call DrawCurve
End Sub

