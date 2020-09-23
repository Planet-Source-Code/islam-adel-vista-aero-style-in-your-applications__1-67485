VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTaskDialogModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Enum TASKDIALOG_FLAGS
    TDF_ENABLE_HYPERLINKS = 1
    TDF_USE_HICON_MAIN = 2
    TDF_USE_HICON_FOOTER = 4
    TDF_ALLOW_DIALOG_CANCELLATION = 8
    TDF_USE_COMMAND_LINKS = &H10
    TDF_USE_COMMAND_LINKS_NO_ICON = &H20
    TDF_EXPAND_FOOTER_AREA = &H40
    TDF_EXPANDED_BY_DEFAULT = &H80
    TDF_VERIFICATION_FLAG_CHECKED = &H100
    TDF_SHOW_PROGRESS_BAR = &H200
    TDF_SHOW_MARQUEE_PROGRESS_BAR = &H400
    TDF_CALLBACK_TIMER = &H800
    TDF_POSITION_RELATIVE_TO_WINDOW = &H1000
    TDF_RTL_LAYOUT = &H2000
    TDF_NO_DEFAULT_RADIO_BUTTON = &H4000
End Enum


Private Type TASKDIALOGCONFIG
    cbSize                      As Long
    hwndParent                  As Long
    hInstance                   As Long
    dwFlags                     As TASKDIALOG_FLAGS
    dwCommonButtons             As TASKDIALOG_COMMON_BUTTON_FLAGS
    pszWindowTitle              As Long
    hMainIcon                   As Long
    pszMainInstruction          As Long
    pszContent                  As Long
    cButtons                    As Long
    pButtons                    As Long
    nDefaultButton              As Long
    cRadioButtons               As Long
    pRadioButtons               As Long
    nDefaultRadioButton         As Long
    pszVerificationText         As Long
    pszExpandedInformation      As Long
    pszExpandedControlText      As Long
    pszCollapsedControlText     As Long
    hFooterIcon                 As Long
    pszFooter                   As Long
    pfCallback                  As Long
    lpCallbackData              As Long
    cxWidth                     As Long
End Type

Private Type TASKDIALOG_BUTTON
    nButtonID                   As Long
    pszButtonText               As Long
End Type

Private m_TaskConfig            As TASKDIALOGCONFIG
Private m_DialogButtons()       As TASKDIALOG_BUTTON
Private m_DialogButtonStr()     As String
Private m_DialogButtonsC        As Long

Private m_WindowTitle           As String
Private m_MainInstruction       As String
Private m_Content               As String
Private m_VerificationText      As String
Private m_ExpandedInformation   As String
Private m_ExpandedControlText   As String
Private m_CollapsedControlText  As String


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
                    Destination As Any, Source As Any, ByVal Length As Long)
Private Function MAKEINTRESOURCE(ByVal iVal As Integer) As Long
    CopyMemory MAKEINTRESOURCE, iVal, 2
End Function
Public Function ConfigPtr() As Long
    ConfigPtr = VarPtr(m_TaskConfig)
End Function

Public Property Get OwnerhWnd() As Long
    OwnerhWnd = m_TaskConfig.hwndParent
End Property

Public Property Let OwnerhWnd(ByVal vNewValue As Long)
    m_TaskConfig.hwndParent = vNewValue
End Property

Public Property Get WindowTitle() As String
    WindowTitle = m_WindowTitle
End Property

Public Property Let WindowTitle(ByVal vNewValue As String)
    m_WindowTitle = vNewValue
End Property

Public Property Get MainIcon() As TASKDIALOG_ICON
    MainIcon = m_TaskConfig.hMainIcon
End Property

Public Property Let MainIcon(ByVal vNewValue As TASKDIALOG_ICON)
    m_TaskConfig.hMainIcon = MAKEINTRESOURCE(vNewValue)
End Property

Public Property Get MainInstruction() As String
    MainInstruction = m_MainInstruction
End Property

Public Property Let MainInstruction(ByVal vNewValue As String)
    m_MainInstruction = vNewValue
End Property

Public Property Get Content() As String
    Content = m_Content
End Property

Public Property Let Content(ByVal vNewValue As String)
    m_Content = vNewValue
End Property

Public Property Get VerificationText() As String
    VerificationText = m_VerificationText
End Property

Public Property Let VerificationText(ByVal vNewValue As String)
    m_VerificationText = vNewValue
End Property

Public Property Get ExpandedInformation() As String
    ExpandedInformation = m_ExpandedInformation
End Property

Public Property Let ExpandedInformation(ByVal vNewValue As String)
    m_ExpandedInformation = vNewValue
End Property

Public Property Get ExpandedControlText() As String
    ExpandedControlText = m_ExpandedControlText
End Property

Public Property Let ExpandedControlText(ByVal vNewValue As String)
    m_ExpandedControlText = vNewValue
End Property

Public Property Get CollapsedControlText() As String
    CollapsedControlText = m_CollapsedControlText
End Property

Public Property Let CollapsedControlText(ByVal vNewValue As String)
    m_CollapsedControlText = vNewValue
End Property

Public Property Get FooterIcon() As TASKDIALOG_ICON
    FooterIcon = m_TaskConfig.hFooterIcon
End Property

Public Property Let FooterIcon(ByVal vNewValue As TASKDIALOG_ICON)
    m_TaskConfig.hFooterIcon = MAKEINTRESOURCE(vNewValue)
End Property

Public Function CreateButton( _
            ByVal m_Text As String, _
            ByVal m_ID As TASKDIALOG_RESULT) As Boolean
    
    
    On Error GoTo er_CreateButton:
    
    
    ReDim Preserve m_DialogButtons(m_DialogButtonsC) As TASKDIALOG_BUTTON
    ReDim Preserve m_DialogButtonStr(m_DialogButtonsC) As String
    
    m_DialogButtonStr(m_DialogButtonsC) = m_Text
    
    With m_DialogButtons(m_DialogButtonsC)
        .nButtonID = m_ID
        .pszButtonText = StrPtr(m_DialogButtonStr(m_DialogButtonsC))
    End With
    
    m_DialogButtonsC = m_DialogButtonsC + 1
    CreateButton = True
    
er_CreateButton:
End Function

Private Sub pPrepareButtons()
    
    With m_TaskConfig
    
        If (m_DialogButtonsC) Then
            .cButtons = m_DialogButtonsC
            .cRadioButtons = m_DialogButtonsC
            
            .pButtons = VarPtr(m_DialogButtons(0).nButtonID)
            .pRadioButtons = VarPtr(m_DialogButtons(0).nButtonID)
        End If
        
    End With
    
End Sub

Public Property Get DefaultButton() As TASKDIALOG_RESULT
    DefaultButton = m_TaskConfig.nDefaultButton
End Property

Public Property Let DefaultButton(ByVal vNewValue As TASKDIALOG_RESULT)
    m_TaskConfig.nDefaultButton = vNewValue
End Property

Private Sub FreeModel()
    
    Erase m_DialogButtons()
    Erase m_DialogButtonStr()
    
End Sub

Private Sub Class_Terminate()
    Call FreeModel
End Sub



Public Sub PrepareModel()
    
    Call pPrepareButtons
    
    With m_TaskConfig
        
        .pszWindowTitle = StrPtr(m_WindowTitle)
        .pszCollapsedControlText = StrPtr(m_CollapsedControlText)
        .pszContent = StrPtr(m_Content)
        .pszExpandedControlText = StrPtr(m_ExpandedControlText)
        .pszExpandedInformation = StrPtr(m_ExpandedInformation)
        .pszVerificationText = StrPtr(m_VerificationText)
        .pszMainInstruction = StrPtr(m_MainInstruction)
        
        .cbSize = Len(m_TaskConfig)
        
    End With
    
End Sub
