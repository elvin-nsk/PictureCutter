VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PreprocessorView 
   Caption         =   "Настройки"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6975
   OleObjectBlob   =   "PreprocessorView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PreprocessorView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' # State

Public IsOk As Boolean
Public IsCancel As Boolean

Public RawTemplatesFolder As FolderBrowserHandler
Public PreparedTemplatesFolder As FolderBrowserHandler

'===============================================================================
' # Constructor

Private Sub UserForm_Initialize()
    Caption = _
        APP_DISPLAYNAME & " - Подготовка шаблонов" & " (v" & APP_VERSION & ")"
    
    Set RawTemplatesFolder = _
        FolderBrowserHandler.New_( _
            RawTemplatesFolderBox, _
            RawTemplatesFolderBrowse _
        )
    Set PreparedTemplatesFolder = _
        FolderBrowserHandler.New_( _
            PreparedTemplatesFolderBox, _
            PreparedTemplatesFolderBrowse _
        )
End Sub

Private Sub UserForm_Activate()
    '
End Sub

Private Sub btnOk_Click()
    FormОК
End Sub

Private Sub btnCancel_Click()
    FormCancel
End Sub

'===============================================================================
' # Logic

Private Sub FormОК()
    Hide
    IsOk = True
End Sub

Private Sub FormCancel()
    Hide
    IsCancel = True
End Sub

'===============================================================================
' # Boilerplate

Private Sub UserForm_QueryClose(Сancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Сancel = True
        FormCancel
    End If
End Sub
