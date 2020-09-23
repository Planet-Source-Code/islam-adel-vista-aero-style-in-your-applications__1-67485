VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVistaGlass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type Margins
    m_Left          As Long
    m_Right         As Long
    m_Top           As Long
    m_Buttom        As Long
End Type


Private Declare Sub DwmIsCompositionEnabled Lib "dwmapi.dll" (cEnabled As Long)
Private Declare Sub DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hwnd As Long, p_Margins As Margins)

Public Function IsGlassEnabled() As Boolean

    Dim isEnabled       As Long
    
    Call DwmIsCompositionEnabled(isEnabled)
    IsGlassEnabled = CBool(isEnabled)
    
End Function

Public Function Glass( _
            ByVal m_hWnd As Long, _
            Optional ByVal m_Left As Long = -1, _
            Optional ByVal m_Right As Long = -1, _
            Optional ByVal m_Top As Long = -1, _
            Optional ByVal m_Buttom As Long = -1) As Boolean
    
    
    Dim n_Margins       As Margins
    
    
    If (Not IsGlassEnabled()) Then
        Exit Function
    End If
    
    With n_Margins
        .m_Left = m_Left
        .m_Right = m_Right
        .m_Top = m_Top
        .m_Buttom = m_Buttom
    End With
    
    DwmExtendFrameIntoClientArea m_hWnd, n_Margins
    Glass = True
    
End Function

