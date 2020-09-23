VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGlassText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const DTT_COMPOSITED    As Long = &H2000
Private Const DTT_GLOWSIZE      As Long = &H800
Private Const DT_SINGLELINE     As Long = &H20
Private Const DT_CENTER         As Long = &H1
Private Const DT_VCENTER        As Long = &H4
Private Const DT_NOPREFIX       As Long = &H800
Private Const DT_TEXTFORMAT     As Long = DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX
Private Const SRCCOPY           As Long = &HCC0020

Private Type RECT
    Left                    As Long
    Top                     As Long
    Right                   As Long
    Bottom                  As Long
End Type

Private Type BITMAPINFOHEADER
    biSize                  As Long
    biWidth                 As Long
    biHeight                As Long
    biPlanes                As Integer
    biBitCount              As Integer
    biCompression           As Long
    biSizeImage             As Long
    biXPelsPerMeter         As Long
    biYPelsPerMeter         As Long
    biClrUsed               As Long
    biClrImportant          As Long
End Type

Private Type RGBQUAD
    rgbBlue                 As Byte
    rgbGreen                As Byte
    rgbRed                  As Byte
    rgbReserved             As Byte
End Type

Private Type BITMAPINFO
    bmiHeader               As BITMAPINFOHEADER
    bmiColors               As RGBQUAD
End Type

Private Type POINTAPI
    x                       As Long
    y                       As Long
End Type

Private Type DTTOPTS
    dwSize                  As Long
    dwFlags                 As Long
    crText                  As Long
    crBorder                As Long
    crShadow                As Long
    iTextShadowType         As Long
    ptShadowOffset          As POINTAPI
    iBorderSize             As Long
    iFontPropId             As Long
    iColorPropId            As Long
    iStateId                As Long
    fApplyOverlay           As Long
    iGlowSize               As Long
    pfnDrawTextCallback     As Long
    lParam                  As Long
End Type


Private m_Wnd               As Long
Private m_GlowSize          As Long
Private m_Text              As String
Private m_Font              As StdFont
Private m_AreaRect          As RECT

Private n_bDraw             As Boolean
Private n_hTheme            As Long
Private n_DC                As Long
Private n_MemDC             As Long
Private n_Bitmap            As Long
Private n_BitmapOld         As Long
Private n_hFont             As Long
Private n_hFontOld          As Long
Private n_dttOpts           As DTTOPTS
Private n_RC1               As RECT
Private n_RC2               As RECT
Private n_DIB               As BITMAPINFO



Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function OpenThemeData Lib "UxTheme" (ByVal hWnd As Long, ByVal szClasses As Long) As Long
Private Declare Function CloseThemeData Lib "UxTheme" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeTextEx Lib "UxTheme" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal text As Long, ByVal iCharCount As Long, ByVal dwFlags As Long, pRect As RECT, pOptions As DTTOPTS) As Long
Public Sub Refresh()

    If (n_bDraw) Then
    
        BitBlt n_DC, _
               n_RC1.Left, _
               n_RC1.Top, _
               n_RC1.Right - n_RC1.Left, _
               n_RC1.Bottom - n_RC1.Top, _
               n_MemDC, _
               0, 0, _
               SRCCOPY
               
    End If
            
End Sub

Public Sub SetData( _
        ByVal hWnd As Long, _
        ByVal nText As String, _
        ByVal nLeft As Long, _
        ByVal nTop As Long)
    
    m_Wnd = hWnd
    
    With m_AreaRect
        .Left = 10
        .Top = 10
        .Right = .Left + Form1.TextWidth(nText) + 5
        .Bottom = .Top + Form1.TextHeight(nText)
    End With
    
    m_GlowSize = 10
    m_Text = nText
    
    pSetData

End Sub

Private Sub Class_Initialize()
    
    n_hTheme = OpenThemeData(m_Wnd, StrPtr("Window"))
    
End Sub


Private Sub Class_Terminate()
    
    CloseThemeData n_hTheme
    pFree
    
End Sub



Private Sub pSetData()
    
    n_bDraw = False

    n_DC = GetDC(m_Wnd)
    
    With n_RC1
        .Top = m_AreaRect.Top
        .Left = m_AreaRect.Left
        .Right = m_AreaRect.Right + 2 * m_GlowSize
        .Bottom = m_AreaRect.Bottom + 2 * m_GlowSize
    End With
    
    With n_RC2
        .Right = n_RC1.Right - n_RC1.Left
        .Bottom = n_RC1.Bottom - n_RC1.Top
    End With
    
    n_MemDC = CreateCompatibleDC(n_DC)

    With n_DIB.bmiHeader
    
        .biSize = Len(n_DIB)
        
        .biHeight = n_RC1.Top - n_RC1.Bottom
        .biWidth = n_RC1.Right - n_RC1.Left
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = 0      ' BI_RGB
        
    End With
    
    If (SaveDC(n_MemDC) <> 0) Then
        
        n_Bitmap = CreateDIBSection(n_MemDC, _
                                    n_DIB, _
                                    0, _
                                    0, _
                                    0, _
                                    0)
        
        If (n_Bitmap <> 0) Then
            
            n_BitmapOld = SelectObject(n_MemDC, n_Bitmap)
            ' n_hFont = m_Font
            n_hFontOld = SelectObject(n_MemDC, n_hFont)

            With n_dttOpts
            
                .dwSize = Len(n_dttOpts)
                .dwFlags = DTT_COMPOSITED Or DTT_GLOWSIZE
                .iGlowSize = m_GlowSize
            
            End With
            
                DrawThemeTextEx n_hTheme, _
                                n_MemDC, _
                                0, 0, _
                                StrPtr(m_Text), _
                                -1, _
                                DT_TEXTFORMAT, _
                                n_RC2, _
                                n_dttOpts
                                
                n_bDraw = True
        End If
    End If
                 
End Sub

Private Sub pFree()

    SelectObject n_MemDC, n_BitmapOld
    SelectObject n_MemDC, n_hFontOld
            
    DeleteObject n_Bitmap
    DeleteObject n_hFont
            
    ReleaseDC n_MemDC, -1
    DeleteDC n_MemDC
    
End Sub
