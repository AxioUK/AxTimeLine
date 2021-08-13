VERSION 5.00
Begin VB.UserControl AxTimeLine 
   AutoRedraw      =   -1  'True
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
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "AxTimeLine.ctx":0000
   Begin Proyecto1.ucScrollbar ucScroll 
      Height          =   3375
      Left            =   4455
      TabIndex        =   0
      Top             =   135
      Width           =   165
      _ExtentX        =   291
      _ExtentY        =   5953
      Style           =   4
      SmoothScrollFactor=   0,15
      BeginProperty ThumbTooltipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "AxTimeLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'-UC-VB6-----------------------------
'UC Name  : AxDashTimeLine
'Version  : 0.03
'Editor   : David Rojas [AxioUK]
'Date     : 22/07/2021
'------------------------------------
Option Explicit

Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function TlsGetValue Lib "kernel32.dll" (ByVal dwTlsIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SetRECL Lib "user32" Alias "SetRect" (lpRect As RECTL, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
'-
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long

'Private Declare Function ReleaseCapture Lib "User32" () As Long
'Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal Token As Long)
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As Long, ByRef mLineGradient As Long) As Long
Private Declare Function GdipDrawRectangle Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GdipFillRectangle Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal Brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDrawRectangleI Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal Brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal mhDC As Long, ByRef mGraphics As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal Brush As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mpath As Long) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mpath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mpath As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal RGBA As Long, ByRef Brush As Long) As Long
Private Declare Function GdipAddPathString Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFamily As Long, ByVal mStyle As Long, ByVal mEmSize As Single, ByRef mLayoutRect As RECTS, ByVal mFormat As Long) As Long
Private Declare Function GdipMeasureString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTS, ByVal mStringFormat As Long, ByRef mBoundingBox As RECTS, ByRef mCodepointsFitted As Long, ByRef mLinesFilled As Long) As Long
Private Declare Function GdipCreateFont Lib "GdiPlus.dll" (ByVal mFontFamily As Long, ByVal mEmSize As Single, ByVal mStyle As Long, ByVal mUnit As Long, ByRef mFont As Long) As Long
Private Declare Function GdipDeleteFont Lib "GdiPlus.dll" (ByVal mFont As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "GdiPlus.dll" (ByRef mNativeFamily As Long) As Long
Private Declare Function GdipDrawString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTS, ByVal mStringFormat As Long, ByVal mBrush As Long) As Long
Private Declare Function GdipSetStringFormatTrimming Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mTrimming As eStringTrimming) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatFlags Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mFlags As eStringFormatFlags) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As eStringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mAlign As eStringAlignment) As Long
Private Declare Function GdipDeleteStringFormat Lib "GdiPlus.dll" (ByVal mFormat As Long) As Long
Private Declare Function GdipDrawLineI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByRef pPoints As Any, ByVal count As Long) As Long
Private Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, ByRef pPoints As Any, ByVal count As Long, ByVal FillMode As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dx As Single, ByVal dy As Single, ByVal Order As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipResetWorldTransform Lib "GdiPlus.dll" (ByVal graphics As Long) As Long
Private Declare Function GdipResetPath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipAddPathCurve Lib "gdiplus" (ByVal path As Long, pPoints As Any, ByVal count As Long) As Long
Private Declare Function GdipDrawEllipse Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipFillEllipse Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipFillEllipseI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDrawClosedCurve Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, POINTS As POINTS, ByVal count As Long) As Long
Private Declare Function GdipFillClosedCurve Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, POINTS As POINTS, ByVal count As Long) As Long
Private Declare Function GdipDrawCurve Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, POINTS As POINTS, ByVal count As Long) As Long
Private Declare Function GdipDrawPie Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipFillPie Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipAddPathEllipse Lib "gdiplus" (ByVal path As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipAddPathEllipseI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipSetPathGradientCenterColor Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mColors As Long) As Long
Private Declare Function GdipSetPathGradientSurroundColorsWithCount Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mColor As Long, ByRef mCount As Long) As Long
Private Declare Function GdipCreatePathGradientFromPath Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mPolyGradient As Long) As Long
Private Declare Function GdipSetLinePresetBlend Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mBlend As Long, ByRef mPositions As Single, ByVal mCount As Long) As Long
'---
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As Rect, ByVal wFormat As Long) As Long
'---
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTL) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type RECTL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type RECTS
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Private Type POINTS
   X As Single
   Y As Single
End Type

Private Type POINTL
    X As Long
    Y As Long
End Type

Private Enum GDIPLUS_FONTSTYLE
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum

Private Type GDIPlusStartupInput
    GdiPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type PicBmp
  Size As Long
  type As Long
  hBmp As Long
  hpal As Long
  Reserved As Long
End Type

Public Enum eStringAlignment
    StringAlignmentNear = &H0
    StringAlignmentCenter = &H1
    StringAlignmentFar = &H2
End Enum
  
Public Enum eStringTrimming
    StringTrimmingNone = &H0
    StringTrimmingCharacter = &H1
    StringTrimmingWord = &H2
    StringTrimmingEllipsisCharacter = &H3
    StringTrimmingEllipsisWord = &H4
    StringTrimmingEllipsisPath = &H5
End Enum

Public Enum eStringFormatFlags
    StringFormatFlagsNone = &H0
    StringFormatFlagsDirectionRightToLeft = &H1
    StringFormatFlagsDirectionVertical = &H2
    StringFormatFlagsNoFitBlackBox = &H4
    StringFormatFlagsDisplayFormatControl = &H20
    StringFormatFlagsNoFontFallback = &H400
    StringFormatFlagsMeasureTrailingSpaces = &H800
    StringFormatFlagsNoWrap = &H1000
    StringFormatFlagsLineLimit = &H2000
    StringFormatFlagsNoClip = &H4000
End Enum

Public Enum eTextAlignH
    eLeft
    eCenter
    eRight
End Enum

Public Enum eTextAlignV
    eTop
    eMiddle
    eBottom
End Enum

Public Enum eTypeLine
  eLine
  eDots
  eBoxs
End Enum

Public Enum pStyle
  pVertical
  pHorizontal
End Enum

Private Type eTimeSection
  Caption1 As String
  Caption2 As String
  Iconchar As Long
  eTime As String
  eDate As String
  Visible As Boolean
End Type

'Constants
Private Const CombineModeExclude As Long = &H4
Private Const WrapModeTileFlipXY = &H3
Private Const SmoothingModeHighQuality As Long = &H2
Private Const SmoothingModeAntiAlias As Long = &H4
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const TLS_MINIMUM_AVAILABLE As Long = 64
Private Const IDC_HAND As Long = 32649
Private Const UnitPixel As Long = &H2&

'Define EVENTS-------------------
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Property Variables:
Private hFontCollection As Long
Private GdipToken As Long
Private nScale    As Single
Private hGraphics As Long
Private hCur      As Long

Private m_Enabled       As Boolean
Private m_BorderColor   As OLE_COLOR
Private m_BackColor     As OLE_COLOR
Private m_BorderWidth   As Long
Private m_CornerCurve   As Long
Private m_ForeColor1    As OLE_COLOR
Private m_ForeColor2    As OLE_COLOR
Private m_Font1         As StdFont
Private m_Font2         As StdFont
Private m_IconFont      As StdFont
Private m_CaptionAlignV As eTextAlignV
Private m_CaptionAlignH As eTextAlignH
Private m_Caption2AlignV As eTextAlignV
Private m_Caption2AlignH As eTextAlignH
Private m_IconForeColor As OLE_COLOR
Private m_IconAlignV    As eTextAlignV
Private m_IconAlignH    As eTextAlignH
Private m_TimeVisible   As Boolean
Private m_DateVisible   As Boolean
Private m_PointBackColor As OLE_COLOR

Private m_ColorActive   As OLE_COLOR
Private m_LineDistance As Long
Private m_LineWidth    As Long
Private m_LineColor     As OLE_COLOR
Private m_TimeLine      As eTypeLine

Private m_ActiveSection As Long
Private m_SectionHW      As Long
Private m_Section()     As eTimeSection
Private SectionCount    As Long
Private m_Style         As pStyle

Public Function AddTimePoint(eCaption1 As String, eCaption2 As String, eIconchar As String, _
                             eTime As String, eDate As String, eVisible As Boolean) As Boolean
ReDim Preserve m_Section(SectionCount)
    
With m_Section(SectionCount)
  .Caption1 = eCaption1
  .Caption2 = eCaption2
  .Iconchar = IconCharCode(eIconchar)
  .eDate = eDate
  .eTime = eTime
  .Visible = eVisible
  
  SectionCount = SectionCount + 1
End With
Refresh
End Function

Public Function ChrW2(ByVal CharCode As Long) As String
  Const POW10 As Long = 2 ^ 10
  If CharCode <= &HFFFF& Then ChrW2 = ChrW$(CharCode) Else _
                              ChrW2 = ChrW$(&HD800& + (CharCode And &HFFFF&) \ POW10) & _
                                      ChrW$(&HDC00& + (CharCode And (POW10 - 1)))
End Function

Public Sub Refresh()
UserControl.Cls
With ucScroll
  .Max = (m_SectionHW * SectionCount) - IIf(m_Style = pVertical, UserControl.ScaleHeight, UserControl.ScaleWidth)
  If ucScroll.Max > 0 Then
    .Visible = True
    .TrackMouseWheelOnHwnd UserControl.hwnd
  Else
    .Visible = False
    .TrackMouseWheelOnHwndStop
  End If
End With
Draw
End Sub

Public Function UpdateTimePoint(ByVal eTimePoint As Long, eVisible As Boolean, Optional eCaption1 As String = vbNullString, _
                                Optional eCaption2 As String = vbNullString, Optional eIconchar As String = vbNullString, _
                                Optional eTime As String = vbNullString, Optional eDate As String = vbNullString) As Boolean
With m_Section(eTimePoint)
   If eCaption1 <> vbNullString Then .Caption1 = eCaption1
   If eCaption2 <> vbNullString Then .Caption2 = eCaption2
   If eIconchar <> vbNullString Then .Iconchar = IconCharCode(eIconchar)
   If eDate <> vbNullString Then .eDate = eDate
   If eTime <> vbNullString Then .eTime = eTime
  .Visible = eVisible
End With
Refresh
End Function

Private Sub Draw()
Dim i As Long, lY As Long
Dim TopH As Long, wlDate As Long
Dim REC As RECTL, IcoBox As RECTS
Dim PTime  As RECTL, PTimeA As RECTL
Dim cp1REC As RECTS, cp2REC As RECTS
Dim rTime As RECTS, rDate As RECTS
Dim cpH As Long, mBorder As Long, lBorder As Long

  GdipCreateFromHDC hdc, hGraphics
  GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias

  'Valores Bordes
  lBorder = m_BorderWidth * 2 * nScale
  mBorder = m_BorderWidth * nScale
  'Control BackColor
  UserControl.BackColor = m_BackColor
  'Scroll Value
  lY = -ucScroll.Value
  
If m_DateVisible Or m_TimeVisible Then
  wlDate = IIf(m_Style = pVertical, UserControl.TextWidth("00-00-0000"), UserControl.TextHeight("000") * 2) 'Ájh
Else
  wlDate = 0
End If

'Altura Captions V
cpH = m_SectionHW / 2 * nScale

'TimeLine
If m_Style = pVertical Then
  SetRECL REC, wlDate, lY + cpH, m_SectionHW, (m_SectionHW * SectionCount) - (cpH * 2)
Else
  SetRECL REC, lY + cpH, wlDate, (m_SectionHW * SectionCount) - (cpH * 2), (m_SectionHW / 4) * 3
End If

DrawLine hGraphics, m_TimeLine, REC, RGBA(m_LineColor, 100), m_LineWidth, m_LineDistance, m_ActiveSection, RGBA(m_ColorActive, 100), m_Style

Do While i <= SectionCount - 1 And lY < UserControl.ScaleHeight
  With m_Section(i)
    If .Visible Then
      If TopH + m_SectionHW > 0 Then
        Select Case m_Style
          Case pHorizontal
          'Point
          SetRECL PTime, lY + (m_SectionHW * i) + ((m_SectionHW / 2) / 2), wlDate + (REC.Height - (m_SectionHW / 2)) / 2, m_SectionHW / 2, m_SectionHW / 2
          'Reset cpH
          cpH = (UserControl.ScaleHeight - (PTime.Top + PTime.Height + 12)) / 2 * nScale
          'IconBox
          SetRECS IcoBox, PTime.Left, PTime.Top, PTime.Width, PTime.Height + 3
          'Caption1
          SetRECS cp1REC, lY + (m_SectionHW * i) + 3, PTime.Top + PTime.Height + 8, m_SectionHW - 6, cpH - 5
          'Caption2
          SetRECS cp2REC, cp1REC.Left, cp1REC.Top + (cpH - 5), cp1REC.Width, cpH
          'Date
          SetRECS rDate, cp1REC.Left, 1, cp1REC.Width, wlDate / 2
          'Time
          SetRECS rTime, cp1REC.Left, wlDate / 2, cp1REC.Width, wlDate / 2
          
          If i = m_ActiveSection Then
            'Point
            DrawRoundRect hGraphics, PTime, RGBA(m_PointBackColor, 100), RGBA(m_BorderColor, 0), m_BorderWidth, m_CornerCurve, True
            'IconChar
            DrawCaption hGraphics, .Iconchar, IconFont, IcoBox, RGBA(m_IconForeColor, 100), 0, m_IconAlignH, m_IconAlignV, True
            'Activepoint
            SetRECL PTimeA, PTime.Left - (mBorder + 2), PTime.Top - (mBorder + 2), PTime.Width + lBorder + 4, PTime.Height + lBorder + 4
            DrawRoundRect hGraphics, PTimeA, RGBA(m_PointBackColor, 10), RGBA(m_ColorActive, 100), m_BorderWidth, m_CornerCurve, False
            'Caption1
            DrawCaption hGraphics, .Caption1, m_Font1, cp1REC, RGBA(m_ForeColor1, 100), 0, m_CaptionAlignH, m_CaptionAlignV, False
            'Caption2
            DrawCaption hGraphics, .Caption2, m_Font2, cp2REC, RGBA(m_ForeColor2, 100), 0, m_Caption2AlignH, m_Caption2AlignV, False
            'Date
            If m_DateVisible Then DrawCaption hGraphics, .eDate, m_Font2, rDate, RGBA(m_ForeColor1, 50), 0, eCenter, eMiddle, False
            'Time
            If m_TimeVisible Then DrawCaption hGraphics, .eTime, m_Font2, rTime, RGBA(m_ForeColor1, 50), 0, eCenter, eTop, False
          Else
            'Point
            DrawRoundRect hGraphics, PTime, RGBA(m_PointBackColor, 100), RGBA(m_BorderColor, 100), m_BorderWidth, m_CornerCurve, True
            'Caption1
            DrawCaption hGraphics, .Caption1, m_Font1, cp1REC, RGBA(m_ForeColor1, 50), 0, m_CaptionAlignH, m_CaptionAlignV, False
            'Date
            If m_DateVisible Then DrawCaption hGraphics, .eDate, m_Font2, rDate, RGBA(m_ForeColor1, 50), 0, eCenter, eMiddle, False
          End If
          
        Case pVertical
          'Point
          SetRECL PTime, wlDate + (REC.Width - (m_SectionHW / 2)) / 2, lY + (m_SectionHW * i) + ((m_SectionHW / 2) / 2), m_SectionHW / 2, m_SectionHW / 2
          'IconBox
          SetRECS IcoBox, PTime.Left, PTime.Top, PTime.Width, PTime.Height + 3
          'Caption1
          SetRECS cp1REC, PTime.Left + PTime.Width + (mBorder + 5), lY + (m_SectionHW * i) + 2, (UserControl.ScaleWidth - (PTime.Left + PTime.Width + (mBorder + 5))) * nScale, cpH
          'Caption2
          SetRECS cp2REC, cp1REC.Left, lY + (m_SectionHW * i) + cpH + 2, cp1REC.Width, cpH
          'Date
          SetRECS rDate, 1, cp1REC.Top, PTime.Left, cp1REC.Height
          'Time
          SetRECS rTime, 1, cp2REC.Top, PTime.Left, cp2REC.Height
          
          If i = m_ActiveSection Then
            'Point
            DrawRoundRect hGraphics, PTime, RGBA(m_PointBackColor, 100), RGBA(m_BorderColor, 0), m_BorderWidth, m_CornerCurve, True
            'IconChar
            DrawCaption hGraphics, .Iconchar, IconFont, IcoBox, RGBA(m_IconForeColor, 100), 0, m_IconAlignH, m_IconAlignV, True
            'Activepoint
            SetRECL PTimeA, PTime.Left - (mBorder + 2), PTime.Top - (mBorder + 2), PTime.Width + lBorder + 4, PTime.Height + lBorder + 4
            DrawRoundRect hGraphics, PTimeA, RGBA(m_PointBackColor, 10), RGBA(m_ColorActive, 100), m_BorderWidth, m_CornerCurve, False
            'Caption1
            DrawCaption hGraphics, .Caption1, m_Font1, cp1REC, RGBA(m_ForeColor1, 100), 0, m_CaptionAlignH, m_CaptionAlignV, False
            'Caption2
            DrawCaption hGraphics, .Caption2, m_Font2, cp2REC, RGBA(m_ForeColor2, 100), 0, m_Caption2AlignH, m_Caption2AlignV, False
            'Date
            If m_DateVisible Then DrawCaption hGraphics, .eDate, m_Font2, rDate, RGBA(m_ForeColor1, 50), 0, eLeft, eBottom, False
            'Time
            If m_TimeVisible Then DrawCaption hGraphics, .eTime, m_Font2, rTime, RGBA(m_ForeColor1, 50), 0, eLeft, eTop, False
          Else
            'Point
            DrawRoundRect hGraphics, PTime, RGBA(m_PointBackColor, 100), RGBA(m_BorderColor, 100), m_BorderWidth, m_CornerCurve, True
            'Caption1
            DrawCaption hGraphics, .Caption1, m_Font1, cp1REC, RGBA(m_ForeColor1, 50), 0, m_CaptionAlignH, m_CaptionAlignV, False
            'Date
            If m_DateVisible Then DrawCaption hGraphics, .eDate, m_Font2, rDate, RGBA(m_ForeColor1, 50), 0, eLeft, eBottom, False
          End If
        End Select
        TopH = TopH + m_SectionHW
      End If
    End If
  End With
  i = i + 1
  Loop

 GdipDeleteGraphics hGraphics
End Sub

Private Function DrawCaption(ByVal hGraphics As Long, sString As Variant, oFont As StdFont, layoutRect As RECTS, _
                             TextColor As Long, mAngle As Single, HAlign As eTextAlignH, VAlign As eTextAlignV, _
                             Icon As Boolean) As Long
Dim hPath As Long
Dim hBrush As Long
Dim hFontFamily As Long
Dim hFormat As Long
Dim lFontSize As Long
Dim lFontStyle As GDIPLUS_FONTSTYLE
Dim newY As Long, newX As Long

On Error Resume Next

    If GdipCreatePath(&H0, hPath) = 0 Then

        If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
            GdipSetStringFormatTrimming hFormat, StringTrimmingEllipsisWord
            GdipSetStringFormatAlign hFormat, HAlign
            GdipSetStringFormatLineAlign hFormat, VAlign
        End If

        GetFontStyleAndSize oFont, lFontStyle, lFontSize

        If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) Then
            If hFontCollection Then
                If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), hFontCollection, hFontFamily) Then
                    If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
                End If
            Else
                If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
            End If
        End If
'------------------------------------------------------------------------
        If mAngle <> 0 Then
            newY = (layoutRect.Height / 2)
            newX = (layoutRect.Width / 2)
            Call GdipTranslateWorldTransform(hGraphics, newX, newY, 0)
            Call GdipRotateWorldTransform(hGraphics, mAngle, 0)
            Call GdipTranslateWorldTransform(hGraphics, -newX, -newY, 0)
        End If
'------------------------------------------------------------------------
      If Icon Then
        GdipAddPathString hPath, StrPtr(ChrW2(sString)), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
      Else
        GdipAddPathString hPath, StrPtr(sString), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
      End If
'------------------------------------------------------------------------
        GdipDeleteStringFormat hFormat
        GdipCreateSolidFill TextColor, hBrush
        GdipFillPath hGraphics, hBrush, hPath
        GdipDeleteBrush hBrush
        If mAngle <> 0 Then GdipResetWorldTransform hGraphics
        GdipDeleteFontFamily hFontFamily
        GdipDeletePath hPath
    End If

End Function

Private Sub DrawLine(ByVal iGraphics As Long, mShape As eTypeLine, Rct As RECTL, LineColor As Long, _
                      BorderW As Long, Distant As Long, ActivePoint As Long, ActiveColor As Long, lStyle As pStyle)
Dim i As Integer
Dim hBrush As Long
Dim hPen As Long
Dim PenActive As Long
Dim BrushActive As Long
Dim X As Long, Y As Long
Dim W As Long, H As Long
Dim p As Long
Dim s As Long

X = Rct.Left:  Y = Rct.Top
W = Rct.Width + 1: H = Rct.Height

GdipCreatePen1 LineColor, IIf(mShape = eBoxs, BorderW / 2, BorderW), UnitPixel, hPen
GdipCreatePen1 ActiveColor, IIf(mShape = eBoxs, BorderW / 2, BorderW), UnitPixel, PenActive

On Error GoTo zEnd
If lStyle = pVertical Then
  p = H / Distant
  'S = ((m_SectionHW * ActivePoint)) / Distant  '- (m_SectionHW)
Else
  p = W / Distant
End If

s = ((m_SectionHW * ActivePoint)) / Distant

ReDim iPts(p) As POINTS
  
  For i = 0 To p
    If lStyle = pVertical Then
      iPts(i).X = X + (W / 2)
      iPts(i).Y = Y + (H / p) * i
    Else
      iPts(i).X = X + (W / p) * i
      iPts(i).Y = Y + (H / 2)
    End If
  Next i
  
Select Case mShape
  Case eLine
      For i = 0 To p - 1
        If i <= s Then
          GdipDrawLineI iGraphics, PenActive, iPts(i).X, iPts(i).Y, iPts(i + 1).X, iPts(i + 1).Y
        Else
          GdipDrawLineI iGraphics, hPen, iPts(i).X, iPts(i).Y, iPts(i + 1).X, iPts(i + 1).Y
        End If
      Next i
      
  Case eDots
      GdipCreateSolidFill LineColor, hBrush
      GdipCreateSolidFill ActiveColor, BrushActive
      For i = 0 To UBound(iPts)
        If i <= s Then
          GdipFillEllipse iGraphics, BrushActive, iPts(i).X - (BorderW / 2), iPts(i).Y - (BorderW / 2), BorderW, BorderW
        Else
          GdipFillEllipse iGraphics, hBrush, iPts(i).X - (BorderW / 2), iPts(i).Y - (BorderW / 2), BorderW, BorderW
        End If
      Next i
      Call GdipDeleteBrush(hBrush)
      Call GdipDeleteBrush(BrushActive)
  
  Case eBoxs
      For i = 0 To p
        If i <= s Then
          GdipDrawRectangleI iGraphics, PenActive, iPts(i).X - BorderW, iPts(i).Y - BorderW, BorderW * 2, BorderW * 2
        Else
          GdipDrawRectangleI iGraphics, hPen, iPts(i).X - BorderW, iPts(i).Y - BorderW, BorderW * 2, BorderW * 2
        End If
      Next i
  
End Select

zEnd:
  Call GdipDeletePen(hPen)
  Call GdipDeletePen(PenActive)

End Sub

Private Function DrawRoundRect(ByVal hGraphics As Long, Rect As RECTL, ByVal BackColor As Long, _
                               ByVal BorderColor As Long, ByVal BorderWidth As Long, _
                               ByVal Round As Long, Filled As Boolean) As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mpath As Long
    Dim mRound As Long
    
    If m_BorderWidth > 0 Then GdipCreatePen1 BorderColor, BorderWidth * nScale, &H2, hPen
    If Filled Then GdipCreateSolidFill BackColor, hBrush
    'GdipCreateLineBrushFromRectWithAngleI Rect, BackColor, BackColor, 90, 0, WrapModeTileFlipXY, hBrush
    
    GdipCreatePath &H0, mpath   '&H0
    
    With Rect
        mRound = GetSafeRound((Round * nScale), .Width * 2, .Height * 2)
        If mRound = 0 Then mRound = 1
            GdipAddPathArcI mpath, .Left, .Top, mRound, mRound, 180, 90
            GdipAddPathArcI mpath, (.Left + .Width) - mRound, .Top, mRound, mRound, 270, 90
            GdipAddPathArcI mpath, (.Left + .Width) - mRound, (.Top + .Height) - mRound, mRound, mRound, 0, 90
            GdipAddPathArcI mpath, .Left, (.Top + .Height) - mRound, mRound, mRound, 90, 90
    End With
    
    GdipClosePathFigures mpath
    GdipFillPath hGraphics, hBrush, mpath
    GdipDrawPath hGraphics, hPen, mpath
    
    Call GdipDeletePath(mpath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)
End Function

Private Function GetFontStyleAndSize(oFont As StdFont, lFontStyle As Long, lFontSize As Long)
On Error GoTo ErrO
    Dim hdc As Long
    lFontStyle = 0
    If oFont.Bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
    
    hdc = GetDC(0&)
    lFontSize = MulDiv(oFont.Size, GetDeviceCaps(hdc, LOGPIXELSY), 72)
    ReleaseDC 0&, hdc
ErrO:
End Function

Private Function GetSafeRound(Angle As Integer, Width As Long, Height As Long) As Integer
    Dim lRet As Integer
    lRet = Angle
    If lRet * 2 > Height Then lRet = Height \ 2
    If lRet * 2 > Width Then lRet = Width \ 2
    GetSafeRound = lRet
End Function

Private Function GetSection(ByVal Z As Single) As Long
    Z = Z + ucScroll.Value
    GetSection = Z \ m_SectionHW
    If GetSection >= SectionCount Then GetSection = -1
End Function

Private Function GetWindowsDPI() As Double
    Dim hdc As Long, LPX  As Double, LPY As Double
    hdc = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hdc, LOGPIXELSX))
    LPY = CDbl(GetDeviceCaps(hdc, LOGPIXELSY))
    ReleaseDC 0, hdc

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

Private Function IconCharCode(ByVal New_IconCharCode As String) As Long
    New_IconCharCode = UCase(Replace(New_IconCharCode, Space(1), vbNullString))
    New_IconCharCode = UCase(Replace(New_IconCharCode, "U+", "&H"))
    If Not VBA.Left$(New_IconCharCode, 2) = "&H" And Not IsNumeric(New_IconCharCode) Then
        IconCharCode = "&H" & New_IconCharCode
    Else
        IconCharCode = New_IconCharCode
    End If
End Function

Private Sub InitGDI()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Sub

Private Function ReadValue(ByVal lProp As Long, Optional Default As Long) As Long
    Dim i       As Long
    For i = 0 To TLS_MINIMUM_AVAILABLE - 1
        If TlsGetValue(i) = lProp Then
            ReadValue = TlsGetValue(i + 1)
            Exit Function
        End If
    Next
    ReadValue = Default
End Function

Private Function RGBA(ByVal RGBColor As Long, ByVal Opacity As Long) As Long
  If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
  RGBA = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
  Opacity = CByte((Abs(Opacity) / 100) * 255)
  If Opacity < 128 Then
      If Opacity < 0& Then Opacity = 0&
      RGBA = RGBA Or Opacity * &H1000000
  Else
      If Opacity > 255& Then Opacity = 255&
      RGBA = RGBA Or (Opacity - 128&) * &H1000000 Or &H80000000
  End If
End Function

Private Sub SafeRange(Value, Min, Max)
    If Value < Min Then Value = Min
    If Value > Max Then Value = Max
End Sub


Private Function SetRECS(lpRect As RECTS, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long) As Long
  lpRect.Left = X
  lpRect.Top = Y
  lpRect.Width = W
  lpRect.Height = H
End Function

Private Sub TerminateGDI()
    Call GdiplusShutdown(GdipToken)
End Sub

Private Sub ucScroll_Change()
Refresh
End Sub

Private Sub ucScroll_Scroll()
Refresh
End Sub

Private Sub UserControl_Initialize()
InitGDI
nScale = GetWindowsDPI
End Sub

Private Sub UserControl_InitProperties()
hFontCollection = ReadValue(&HFC)
m_Enabled = True
  m_Style = pVertical
  
  m_BorderColor = &HFF8080
  m_BackColor = Extender.Container.BackColor
  m_PointBackColor = &HFFC0C0
  m_BorderWidth = 1
  m_CornerCurve = 10
  m_ForeColor1 = &H404040
  m_ForeColor2 = &H404040
  m_IconForeColor = &HFFFFFF
  Set m_Font1 = UserControl.Font
  Set m_Font2 = UserControl.Font
  Set m_IconFont = UserControl.Font
  m_CaptionAlignV = eMiddle
  m_CaptionAlignH = eLeft
  m_Caption2AlignV = eMiddle
  m_Caption2AlignH = eLeft
  m_IconAlignV = eMiddle
  m_IconAlignH = eCenter
  
  m_ActiveSection = 1
  m_SectionHW = 50
  m_ColorActive = vbRed
  
  m_LineDistance = 10
  m_LineWidth = 2
  m_LineColor = &HFFC0C0
  m_TimeLine = eLine

  m_TimeVisible = False
  m_DateVisible = False
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ActiveSection = IIf(m_Style = pVertical, GetSection(Y), GetSection(X))
RaiseEvent MouseDown(Button, Shift, X, Y)
RaiseEvent Click
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
  m_Enabled = .ReadProperty("Enabled", True)
  m_Style = .ReadProperty("Style", 0)

  m_BorderColor = .ReadProperty("BorderColor", &HFF8080)
  m_BackColor = .ReadProperty("BackColor", &H8000000F)
  
  m_BorderWidth = .ReadProperty("BorderWidth", 1)
  m_CornerCurve = .ReadProperty("CornerCurve", 10)
  m_ForeColor1 = .ReadProperty("Caption1Color", &HFFFFFF)
  m_ForeColor2 = .ReadProperty("Caption2Color", &HFFFFFF)
  m_IconForeColor = .ReadProperty("IconForeColor", &HFFFFFF)
  Set m_Font1 = .ReadProperty("Caption1Font", UserControl.Font)
  Set m_Font2 = .ReadProperty("Caption2Font", UserControl.Font)
  Set m_IconFont = .ReadProperty("IconFont", UserControl.Font)
  m_CaptionAlignV = .ReadProperty("Caption1AlignV", 1)
  m_CaptionAlignH = .ReadProperty("Caption1AlignH", 1)
  m_Caption2AlignV = .ReadProperty("Caption2AlignV", 1)
  m_Caption2AlignH = .ReadProperty("Caption2AlignH", 1)
  m_IconAlignV = .ReadProperty("IconAlignV", 1)
  m_IconAlignH = .ReadProperty("IconAlignH", 1)
  
  m_ActiveSection = .ReadProperty("ActiveSection", 1)
  m_SectionHW = .ReadProperty("SectionSpace", 50)
  m_ColorActive = .ReadProperty("BorderColorActive", vbRed)
  m_PointBackColor = .ReadProperty("PointBackColor", &HFFC0C0)
  m_LineDistance = .ReadProperty("LineDistance", 10)
  m_LineWidth = .ReadProperty("LineWidth", 2)
  m_LineColor = .ReadProperty("LineColor", &HFFC0C0)
  m_TimeLine = .ReadProperty("LineStyle", 0)

  m_TimeVisible = .ReadProperty("TimeVisible", False)
  m_DateVisible = .ReadProperty("DateVisible", False)
End With
End Sub

Private Sub UserControl_Resize()
If m_Style = pVertical Then
  ucScroll.Move UserControl.ScaleWidth - 10, 0, 10, UserControl.ScaleHeight
Else
  ucScroll.Move 0, UserControl.ScaleHeight - 10, UserControl.ScaleWidth, 10
End If
Refresh
End Sub

Private Sub UserControl_Terminate()
TerminateGDI
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
  Call .WriteProperty("Enabled", m_Enabled)
  Call .WriteProperty("Style", m_Style)
  
  Call .WriteProperty("BorderColor", m_BorderColor)
  Call .WriteProperty("BackColor", m_BackColor)
  Call .WriteProperty("BorderWidth", m_BorderWidth)
  Call .WriteProperty("CornerCurve", m_CornerCurve)
  Call .WriteProperty("Caption1Color", m_ForeColor1)
  Call .WriteProperty("Caption2Color", m_ForeColor2)
  Call .WriteProperty("IconForeColor", m_IconForeColor)
  Call .WriteProperty("Caption1Font", m_Font1)
  Call .WriteProperty("Caption2Font", m_Font2)
  Call .WriteProperty("IconFont", m_IconFont)
  Call .WriteProperty("Caption1AlignV", m_CaptionAlignV)
  Call .WriteProperty("Caption1AlignH", m_CaptionAlignH)
  Call .WriteProperty("Caption2AlignV", m_Caption2AlignV)
  Call .WriteProperty("Caption2AlignH", m_Caption2AlignH)
  Call .WriteProperty("IconAlignV", m_IconAlignV)
  Call .WriteProperty("IconAlignH", m_IconAlignH)
  
  Call .WriteProperty("ActiveSection", m_ActiveSection)
  Call .WriteProperty("SectionSpace", m_SectionHW)
  Call .WriteProperty("BorderColorActive", m_ColorActive)
  Call .WriteProperty("PointBackColor", m_PointBackColor)
  Call .WriteProperty("LineDistance", m_LineDistance)
  Call .WriteProperty("LineWidth", m_LineWidth)
  Call .WriteProperty("LineColor", m_LineColor)
  Call .WriteProperty("LineStyle", m_TimeLine)
  
  Call .WriteProperty("TimeVisible", m_TimeVisible, False)
  Call .WriteProperty("DateVisible", m_DateVisible, False)
    
End With
  
End Sub

Public Property Get ActiveSection() As Long
  ActiveSection = m_ActiveSection
End Property

Public Property Let ActiveSection(ByVal NewActiveSection As Long)
  m_ActiveSection = NewActiveSection
  PropertyChanged "ActiveSection"
  Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)
  m_BackColor = New_Color
  PropertyChanged "BackColor"
  Refresh
End Property

Public Property Get BorderColorActive() As OLE_COLOR
  BorderColorActive = m_ColorActive
End Property

Public Property Let BorderColorActive(ByVal NewColorActive As OLE_COLOR)
  m_ColorActive = NewColorActive
  PropertyChanged "BorderColorActive"
  Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
  BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)
  m_BorderColor = NewBorderColor
  PropertyChanged "BorderColor"
  Refresh
End Property

Public Property Get BorderWidth() As Long
  BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal NewBorderWidth As Long)
  m_BorderWidth = NewBorderWidth
  PropertyChanged "BorderWidth"
  Refresh
End Property

Public Property Get Caption1AlignH() As eTextAlignH
  Caption1AlignH = m_CaptionAlignH
End Property

Public Property Let Caption1AlignH(ByVal NewCaptionAlignH As eTextAlignH)
  m_CaptionAlignH = NewCaptionAlignH
  PropertyChanged "Caption1AlignH"
  Refresh
End Property

Public Property Get Caption1AlignV() As eTextAlignV
  Caption1AlignV = m_CaptionAlignV
End Property

Public Property Let Caption1AlignV(ByVal NewCaptionAlignV As eTextAlignV)
  m_CaptionAlignV = NewCaptionAlignV
  PropertyChanged "Caption1AlignV"
  Refresh
End Property

Public Property Get Caption1Color() As OLE_COLOR
  Caption1Color = m_ForeColor1
End Property

Public Property Let Caption1Color(ByVal NewForeColor As OLE_COLOR)
  m_ForeColor1 = NewForeColor
  PropertyChanged "Caption1Color"
  Refresh
End Property

Public Property Get Caption1Font() As StdFont
  Set Caption1Font = m_Font1
End Property

Public Property Set Caption1Font(ByVal New_Font As StdFont)
  Set m_Font1 = New_Font
  PropertyChanged "Caption1Font"
  Refresh
End Property

Public Property Get Caption2AlignH() As eTextAlignH
  Caption2AlignH = m_Caption2AlignH
End Property

Public Property Let Caption2AlignH(ByVal NewCaptionAlignH As eTextAlignH)
  m_Caption2AlignH = NewCaptionAlignH
  PropertyChanged "Caption2AlignH"
  Refresh
End Property

Public Property Get Caption2AlignV() As eTextAlignV
  Caption2AlignV = m_Caption2AlignV
End Property

Public Property Let Caption2AlignV(ByVal NewCaptionAlignV As eTextAlignV)
  m_Caption2AlignV = NewCaptionAlignV
  PropertyChanged "Caption2AlignV"
  Refresh
End Property

Public Property Get Caption2Color() As OLE_COLOR
  Caption2Color = m_ForeColor2
End Property

Public Property Let Caption2Color(ByVal NewForeColor As OLE_COLOR)
  m_ForeColor2 = NewForeColor
  PropertyChanged "Caption2Color"
  Refresh
End Property

Public Property Get Caption2Font() As StdFont
  Set Caption2Font = m_Font2
End Property

Public Property Set Caption2Font(ByVal New_Font As StdFont)
  Set m_Font2 = New_Font
  PropertyChanged "Caption2Font"
  Refresh
End Property

Public Property Get CornerCurve() As Long
  CornerCurve = m_CornerCurve
End Property

Public Property Let CornerCurve(ByVal NewCornerCurve As Long)
  m_CornerCurve = NewCornerCurve
  PropertyChanged "CornerCurve"
  Refresh
End Property

Public Property Get DateVisible() As Boolean
  DateVisible = m_DateVisible
End Property

Public Property Let DateVisible(ByVal NewDateVisible As Boolean)
  m_DateVisible = NewDateVisible
  PropertyChanged "DateVisible"
  Refresh
End Property

Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  m_Enabled = New_Enabled
  PropertyChanged "Enabled"
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get IconAlignH() As eTextAlignH
  IconAlignH = m_IconAlignH
End Property

Public Property Let IconAlignH(ByVal NewIconAlignH As eTextAlignH)
  m_IconAlignH = NewIconAlignH
  PropertyChanged "IconAlignH"
  Refresh
End Property

Public Property Get IconAlignV() As eTextAlignV
  IconAlignV = m_IconAlignV
End Property

Public Property Let IconAlignV(ByVal NewIconAlignV As eTextAlignV)
  m_IconAlignV = NewIconAlignV
  PropertyChanged "IconAlignV"
  Refresh
End Property

Public Property Get IconFont() As StdFont
    Set IconFont = m_IconFont
End Property

Public Property Set IconFont(New_Font As StdFont)
  Set m_IconFont = New_Font
    PropertyChanged "IconFont"
  Refresh
End Property

Public Property Get IconForeColor() As OLE_COLOR
    IconForeColor = m_IconForeColor
End Property

Public Property Let IconForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_IconForeColor = New_ForeColor
    PropertyChanged "IconForeColor"
    Refresh
End Property

Public Property Get LineColor() As OLE_COLOR
  LineColor = m_LineColor
End Property

Public Property Let LineColor(ByVal NewLineColor As OLE_COLOR)
  m_LineColor = NewLineColor
  PropertyChanged "LineColor"
  Refresh
End Property

Public Property Get LineDistance() As Long
  LineDistance = m_LineDistance
End Property

Public Property Let LineDistance(ByVal NewLineDistance As Long)
  m_LineDistance = NewLineDistance
  PropertyChanged "LineDistance"
  Refresh
End Property

Public Property Get LineStyle() As eTypeLine
  LineStyle = m_TimeLine
End Property

Public Property Let LineStyle(ByVal NewTimeLine As eTypeLine)
  m_TimeLine = NewTimeLine
  PropertyChanged "LineStyle"
  Refresh
End Property

Public Property Get LineWidth() As Long
  LineWidth = m_LineWidth
End Property

Public Property Let LineWidth(ByVal NewLineWidth As Long)
  m_LineWidth = NewLineWidth
  PropertyChanged "LineWidth"
  Refresh
End Property

Public Property Get PointBackColor() As OLE_COLOR
  PointBackColor = m_PointBackColor
End Property

Public Property Let PointBackColor(ByVal New_Color As OLE_COLOR)
  m_PointBackColor = New_Color
  PropertyChanged "PointBackColor"
  Refresh
End Property

Public Property Get SectionSpace() As Long
  SectionSpace = m_SectionHW
End Property

Public Property Let SectionSpace(ByVal NewSectionSpace As Long)
  m_SectionHW = NewSectionSpace
  PropertyChanged "SectionSpace"
  Refresh
End Property

Public Property Get Style() As pStyle
  Style = m_Style
End Property

Public Property Let Style(ByVal NewStyle As pStyle)
  m_Style = NewStyle
  PropertyChanged "Style"
  ucScroll.Orientation = m_Style
  UserControl_Resize
End Property

Public Property Get TimeVisible() As Boolean
  TimeVisible = m_TimeVisible
End Property

Public Property Let TimeVisible(ByVal NewTimeVisible As Boolean)
  m_TimeVisible = NewTimeVisible
  PropertyChanged "TimeVisible"
  Refresh
End Property

Public Property Get Version() As String
Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Property Get Visible() As Boolean
  Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal NewVisible As Boolean)
  Extender.Visible = NewVisible
End Property



