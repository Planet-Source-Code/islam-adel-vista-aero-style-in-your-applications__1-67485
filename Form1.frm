VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   204
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private GText           As cGlassText


Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long


Private Sub Command1_Click()
    
    GText.SetData Me.hwnd, InputBox(""), 10, 10
    GText.Refresh
    
End Sub

Private Sub Command2_Click()
    
    Dim pm      As New cTaskDialogModel
    Dim td      As New cTaskDialog
    Dim a       As Long, _
        b       As Long
    
    
    With pm
    
        .CollapsedControlText = "Where the heck did this come from?"
        .Content = "Here it is."
        .ExpandedControlText = "Ah, got it!"
        .ExpandedInformation = "An example provided by Islam Adel, http://www.planet-source-code.com"
        .MainInstruction = "Do you want to view the larger TaskDialog?"
        .VerificationText = "Don't show this message again"
        .WindowTitle = "A Larger Task Dialog"
        
        .OwnerhWnd = Me.hwnd
        .DefaultButton = IDNO
        
        .FooterIcon = TD_NONE_ICON
        .MainIcon = TD_INFORMATION_ICON
        
        .CreateButton "Yeah", IDYES
        .CreateButton "Nah", IDNO
        .CreateButton "Repeat the question?", IDCANCEL
        
        .PrepareModel
        
    End With
    
    td.DialogEx pm, a, b
    
    Set td = Nothing
    Set pm = Nothing
    
End Sub

Private Sub Form_Initialize()
    Call InitCommonControls
End Sub

Private Sub Form_Load()
   
    Dim i As New cVistaGlass
    
    If (i.Glass(Me.hwnd)) Then
        Me.BackColor = vbBlack
    End If
    
    Set i = Nothing
    
    Set GText = New cGlassText
    
End Sub

Private Sub Form_Paint()
    
    GText.Refresh
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set GText = Nothing

End Sub


