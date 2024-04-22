VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PictureCutterView 
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6975
   OleObjectBlob   =   "PictureCutterView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PictureCutterView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' # State

Private Const MIN_SIZE As Double = 1

Public IsOk As Boolean
Public IsCancel As Boolean

Public SourceFolder As FolderBrowserHandler
Public OutputFolder As FolderBrowserHandler
Public DivWidth As TextBoxHandler
Public DivHeight As TextBoxHandler
Public ShortestSide As TextBoxHandler
Public SizeDelta As TextBoxHandler

Public ImagesQuantity As TextBoxHandler
Public HTemplatesFolder As FolderBrowserHandler
Public VTemplatesFolder As FolderBrowserHandler
Public ETemplatesFolder As FolderBrowserHandler

'===============================================================================
' # Constructor

Private Sub UserForm_Initialize()
    Caption = APP_DISPLAYNAME & " (v" & APP_VERSION & ")"
    
    With ComboBoxResolution
        .AddItem 72
        .AddItem 96
        .AddItem 150
        .AddItem 300
    End With
    
    Set SourceFolder = _
        FolderBrowserHandler.New_(SourceFolderBox, SourceFolderBrowse)
    Set OutputFolder = _
        FolderBrowserHandler.New_(OutputFolderBox, OutputFolderBrowse)
    Set DivWidth = _
        TextBoxHandler.New_(DivWidthBox, TextBoxTypeLong, 1)
    Set DivHeight = _
        TextBoxHandler.New_(DivHeightBox, TextBoxTypeLong, 1)
    Set SizeDelta = _
        TextBoxHandler.New_(SizeDeltaBox, TextBoxTypeDouble, 0)
    Set ShortestSide = _
        TextBoxHandler.New_(ShortestSideBox, TextBoxTypeDouble, MIN_SIZE)
        
    Set ImagesQuantity = _
        TextBoxHandler.New_(ImagesQuantityBox, TextBoxTypeLong, 1)
    Set HTemplatesFolder = _
        FolderBrowserHandler.New_(HTemplatesFolderBox, HTemplatesFolderBrowse)
    Set VTemplatesFolder = _
        FolderBrowserHandler.New_(VTemplatesFolderBox, VTemplatesFolderBrowse)
    Set ETemplatesFolder = _
        FolderBrowserHandler.New_(ETemplatesFolderBox, ETemplatesFolderBrowse)
End Sub

'===============================================================================
' # Handlers

Private Sub UserForm_Activate()
    '
End Sub

Private Sub btnOk_Click()
    FormŒ 
End Sub

Private Sub btnCancel_Click()
    FormCancel
End Sub

'===============================================================================
' # Logic

Private Sub FormŒ ()
    Hide
    IsOk = True
End Sub

Private Sub FormCancel()
    Hide
    IsCancel = True
End Sub

'===============================================================================
' # Helpers


'===============================================================================
' # Boilerplate

Private Sub UserForm_QueryClose(—ancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        —ancel = True
        FormCancel
    End If
End Sub
