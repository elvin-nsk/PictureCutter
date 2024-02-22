Attribute VB_Name = "PictureCutter"
'===============================================================================
'   Макрос          : PictureCutter
'   Версия          : 2024.02.22
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "PictureCutter"
Public Const APP_DISPLAYNAME As String = APP_NAME
Public Const APP_VERSION As String = "2024.02.22"

'===============================================================================
' # Entry points

Sub Start()

    #If DebugMode = 0 Then
    On Error GoTo Catch
    Optimization = True
    #End If
    
    VBA.Randomize
    
    Dim Cfg As New Config
    If Not ShowViewAndGetConfig(Cfg) Then GoTo Finally
    
    ProcessImages Cfg
    
Finally:
    #If DebugMode = 0 Then
    Optimization = False
    #End If
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================
' # Helpers

Private Sub ProcessImages(ByVal Cfg As Config)
    Dim File As Scripting.File
    Dim Files As Scripting.Files: Set Files = _
        FSO.GetFolder(Cfg.SourceFolder).Files
    Dim PBar As ProgressBar: Set PBar = _
        ProgressBar.New_(ProgressBarNumeric, Files.Count)
    PBar.Cancelable = True
    For Each File In Files
        If CheckFile(File.Path) Then
            ProcessImage File.Path, Cfg
            PBar.Update
            If PBar.Canceled Then Exit Sub
        End If
    Next File
End Sub

Private Sub ProcessImage( _
                ByVal File As String, _
                ByVal Cfg As Config _
            )
    Dim Doc As Document: Set Doc = OpenDocument(File)
    Dim SourceShape As Shape: Set SourceShape = Doc.ActivePage.Shapes.First
    Dim DisplayWidth As String, DisplayHeight As String
    CalcDisplaySizes SourceShape, DisplayWidth, DisplayHeight, Cfg
        
    Dim Slices As Collection: Set Slices = SliceShape(SourceShape, Cfg)
    
    Dim FileNameCalc As FileSpec: Set FileNameCalc = FileSpec.New_(File)
    Dim SourceFileBaseName As String: SourceFileBaseName = FileNameCalc.BaseName
    With FileNameCalc
        .Path = Cfg.OutputFolder
        .Path = .Path & .BaseName
        .BaseName = _
            .BaseName _
          & "_" & DisplayWidth & "_x_" & DisplayHeight & "_" & DisplayUnit(Cfg)
        MakePath .Path
    End With
    FSO.CopyFile File, FileNameCalc
    
    SaveSlices Slices, FileNameCalc.Path, SourceFileBaseName
    
    Doc.Close
End Sub

Private Function SliceShape( _
                     ByVal Shape As Shape, _
                     ByVal Cfg As Config _
                 ) As Collection
    Set SliceShape = New Collection
    If Cfg.DivWidth = 1 And Cfg.DivHeight = 1 Then
        SliceShape.Add Shape
        Exit Function
    End If
    Dim SliceWidth As Double: SliceWidth = Shape.SizeWidth / Cfg.DivWidth
    Dim SliceHeight As Double: SliceHeight = Shape.SizeHeight / Cfg.DivHeight
    Dim HStep As Long
    Dim VStep As Long
    Dim Temp As Shape
    For VStep = 1 To Cfg.DivHeight
        For HStep = 1 To Cfg.DivWidth
            Set Temp = Shape.Duplicate
            SliceShape.Add _
                CropTool( _
                    Temp, _
                    Shape.LeftX + SliceWidth * (HStep - 1), _
                    Shape.TopY - SliceHeight * (VStep - 1), _
                    Shape.LeftX + SliceWidth * HStep, _
                    Shape.TopY - SliceHeight * VStep _
                ).FirstShape
        Next HStep
    Next VStep
    Shape.Delete
End Function

Private Sub SaveSlices( _
                ByVal Slices As Collection, _
                ByVal SavePath As String, _
                ByVal SourceFileBaseName As String _
            )
    Dim Slice As Shape
    Dim Index As Long
    For Each Slice In Slices
        Index = Index + 1
        SaveSlice Slice, Index, SavePath, SourceFileBaseName
    Next Slice
End Sub

Private Sub SaveSlice( _
                ByVal Slice As Shape, _
                ByVal Index As Long, _
                ByVal SavePath As String, _
                ByVal SourceFileBaseName As String _
            )
    
    Dim File As FileSpec: Set File = FileSpec.New_
    File.Path = SavePath
    File.BaseName = _
        SourceFileBaseName & "_" & VBA.Format(Index, "00")
    File.Ext = "png"
    Slice.Bitmap.SaveAs(File, cdrPNG).Finish
End Sub

Private Sub CalcDisplaySizes( _
                ByVal Shape As Shape, _
                ByRef DisplayWidth As String, _
                ByRef DisplayHeight As String, _
                ByVal Cfg As Config _
            )
    Dim Ratio As Double: Ratio = Shape.SizeWidth / Shape.SizeHeight
    Dim RndWidth As Double: RndWidth = RndDouble(Cfg.MinWidth, Cfg.MaxWidth)
    Dim Height As Double: Height = RndWidth / Ratio
    DisplayWidth = ToStr(VBA.Round(RndWidth, 1), ".")
    DisplayHeight = ToStr(VBA.Round(Height, 1), ".")
End Sub

Private Property Get DisplayUnit(ByVal Cfg As Config) As String
    Select Case True
        Case Cfg.OptionCentimeters: DisplayUnit = "sm"
        Case Cfg.OptionInches: DisplayUnit = "in"
    End Select
End Property

Private Property Get CheckFile(ByVal ImageFile As String) As Boolean
    Dim File As FileSpec: Set File = FileSpec.New_(ImageFile)
    If MatchAnyOf(File.Ext, "png", "jpg", "jpeg", "tif") Then CheckFile = True
End Property

Private Function ShowViewAndGetConfig(ByVal Cfg As Config) As Boolean
    With New PictureCutterView
        .SourceFolder = Cfg.SourceFolder
        .OutputFolder = Cfg.OutputFolder
        .DivWidth = Cfg.DivWidth
        .DivHeight = Cfg.DivHeight
        .MinWidth = Cfg.MinWidth
        .MaxWidth = Cfg.MaxWidth
        .OptionInches = Cfg.OptionInches
        .OptionCentimeters = Cfg.OptionCentimeters
        
        .Show vbModal
        
        Cfg.SourceFolder = .SourceFolder
        Cfg.OutputFolder = .OutputFolder
        Cfg.DivWidth = .DivWidth
        Cfg.DivHeight = .DivHeight
        Cfg.MinWidth = .MinWidth
        Cfg.MaxWidth = .MaxWidth
        Cfg.OptionInches = .OptionInches
        Cfg.OptionCentimeters = .OptionCentimeters
        
        ShowViewAndGetConfig = .IsOk
    End With
End Function

'===============================================================================
' # Tests

Private Sub testSomething()
End Sub
