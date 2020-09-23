VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTaskDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const WM_USER           As Long = &H400


Public Enum TASKDIALOG_ICON
    TD_NONE_ICON = 0
    TD_WARNING_ICON = -1
    TD_ERROR_ICON = -2
    TD_INFORMATION_ICON = -3
    TD_SHIELD_ICON = -4
End Enum

Public Enum TASKDIALOG_RESULT
    IDOK = 1
    IDCANCEL = 2
    IDABORT = 3
    IDRETRY = 4
    IDIGNORE = 5
    IDYES = 6
    IDNO = 7
End Enum

Public Enum TASKDIALOG_MESSAGES
    TDM_NAVIGATE_PAGE = WM_USER + 101
    TDM_CLICK_BUTTON = WM_USER + 102
    TDM_SET_MARQUEE_PROGRESS_BAR = WM_USER + 103
    TDM_SET_PROGRESS_BAR_STATE = WM_USER + 104
    TDM_SET_PROGRESS_BAR_RANGE = WM_USER + 105
    TDM_SET_PROGRESS_BAR_POS = WM_USER + 106
    TDM_SET_PROGRESS_BAR_MARQUEE = WM_USER + 107
    TDM_SET_ELEMENT_TEXT = WM_USER + 108
    TDM_CLICK_RADIO_BUTTON = WM_USER + 110
    TDM_ENABLE_BUTTON = WM_USER + 111
    TDM_ENABLE_RADIO_BUTTON = WM_USER + 112
    TDM_CLICK_VERIFICATION = WM_USER + 113
    TDM_UPDATE_ELEMENT_TEXT = WM_USER + 114
    TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE = WM_USER + 115
    TDM_UPDATE_ICON = WM_USER + 116
End Enum

Public Enum TASKDIALOG_NOTIFICATIONS
    TDN_CREATED = 0
    TDN_NAVIGATED = 1
    TDN_BUTTON_CLICKED = 2
    TDN_HYPERLINK_CLICKED = 3
    TDN_TIMER = 4
    TDN_DESTROYED = 5
    TDN_RADIO_BUTTON_CLICKED = 6
    TDN_DIALOG_CONSTRUCTED = 7
    TDN_VERIFICATION_CLICKED = 8
    TDN_HELP = 9
    TDN_EXPANDO_BUTTON_CLICKED = 10
End Enum

Public Enum TASKDIALOG_ELEMENTS
    TDE_CONTENT = 0
    TDE_EXPANDED_INFORMATION = 1
    TDE_FOOTER = 2
    TDE_MAIN_INSTRUCTION = 3
End Enum

Public Enum TASKDIALOG_ICON_ELEMENTS
    TDIE_ICON_MAIN = 0
    TDIE_ICON_FOOTER = 1
End Enum

Public Enum TASKDIALOG_COMMON_BUTTON_FLAGS
    TDCBF_OK_BUTTON = 1
    TDCBF_YES_BUTTON = 2
    TDCBF_NO_BUTTON = 4
    TDCBF_CANCEL_BUTTON = 8
    TDCBF_RETRY_BUTTON = &H10
    TDCBF_CLOSE_BUTTON = &H20
End Enum


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
                    Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function TaskDialogIndirect Lib "comctl32.dll" ( _
                    ByVal pTaskConfig As Long, pnButton As Long, _
                    pnRadioButton As Long, pfVerificationFlagChecked As Long) As Long

Private Declare Function TaskDialog Lib "comctl32.dll" ( _
                    ByVal hwndParent As Long, ByVal hInstance As Long, _
                    ByVal pszWindowTitle As Long, ByVal pszMainInstruction As Long, _
                    ByVal pszContent As Long, _
                    ByVal dwCommonButtons As TASKDIALOG_COMMON_BUTTON_FLAGS, _
                    ByVal pszIcon As Long, pnButton As Long) As Long


Private m_OwnerhWnd         As Long
Public Function Dialog( _
            ByVal m_Prompt As String, _
            Optional ByVal m_Content As String = "", _
            Optional ByVal m_Title As String = "", _
            Optional ByVal m_Buttons As TASKDIALOG_COMMON_BUTTON_FLAGS = TDCBF_OK_BUTTON, _
            Optional ByVal m_PromptIcon As TASKDIALOG_ICON = TD_INFORMATION_ICON) As TASKDIALOG_RESULT
    
    
    Dim l_Ret       As Long
    
    
    l_Ret = TaskDialog(OwnerHandle, _
                       0, _
                       StrPtr(m_Title), _
                       StrPtr(m_Prompt), _
                       StrPtr(m_Content), _
                       m_Buttons, _
                       MAKEINTRESOURCE(m_PromptIcon), _
                       Dialog)
    
End Function

Private Function MAKEINTRESOURCE(ByVal iVal As Integer) As Long
    CopyMemory MAKEINTRESOURCE, iVal, 2
End Function
Public Property Get OwnerHandle() As Long
    OwnerHandle = m_OwnerhWnd
End Property

Public Property Let OwnerHandle(ByVal vNewValue As Long)
    m_OwnerhWnd = vNewValue
End Property

Public Function DialogEx( _
            ByRef p_TaskModel As cTaskDialogModel, _
            ByRef r_RadioButton As Long, _
            ByRef r_VerificationFlagChecked As Long) As TASKDIALOG_RESULT
    
    
    Dim l_Ret       As Long, _
        r_Button    As Long
    
    
    If (p_TaskModel.OwnerhWnd = 0) Then
        If (OwnerHandle) Then
            p_TaskModel.OwnerhWnd = OwnerHandle
        End If
    End If
    
    l_Ret = TaskDialogIndirect(p_TaskModel.ConfigPtr(), _
                               r_Button, _
                               r_RadioButton, _
                               r_VerificationFlagChecked)
    
    DialogEx = r_Button
    
End Function
