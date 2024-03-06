Attribute VB_Name = "PictureCutter"
'===============================================================================
'   Макрос          : PictureCutter
'   Версия          : 2024.03.06
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "PictureCutter"
Public Const APP_DISPLAYNAME As String = APP_NAME
Public Const APP_VERSION As String = "2024.03.06"

Public Const RECTANGLE_SIZE_PX As Long = 500

'===============================================================================
' # Entry points

Sub Prepare()

    #If DebugMode = 0 Then
    On Error GoTo Catch
    Optimization = True
    #End If
    
    Dim Cfg As New Config
    If Not ShowPreprocessorViewAndGetConfig(Cfg) Then GoTo Finally
    
    PreprocessTemplates Cfg
    
Finally:
    #If DebugMode = 0 Then
    Optimization = False
    #End If
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

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
    Dim ImageSize As Size: Set ImageSize = RandomizeSize(SourceShape, Cfg)
    
    Dim FileNameCalc As FileSpec: Set FileNameCalc = FileSpec.New_(File)
    Dim SourceFileBaseName As String: SourceFileBaseName = FileNameCalc.BaseName
    With FileNameCalc
        .Path = Cfg.OutputFolder
        .Path = .Path & .BaseName
        .BaseName = _
            .BaseName & "_" _
          & ImageSize.DisplayWidth(1, ".") & "_x_" _
          & ImageSize.DisplayHeight(1, ".") & "_" & DisplayUnit(Cfg)
        MakePath .Path
    End With
    FSO.CopyFile File, FileNameCalc
    
    Dim Slices As Collection: Set Slices = SliceShape(SourceShape, Cfg)
    SaveSlices Slices, FileNameCalc.Path, SourceFileBaseName
    
    Doc.Close
    
    If Cfg.OptionImageOnRandomTemplate Then _
        ExportOnTemplates File, ImageSize, FileNameCalc.Path, Cfg
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

'-------------------------------------------------------------------------------

Public Sub PreprocessTemplates(ByVal Cfg As Config)
    Dim RawTemplateFiles As Collection
    Set RawTemplateFiles = GetValidImagesFromFolder(Cfg.RawTemplatesFolder)
    If Not CheckImagesCount(RawTemplateFiles, Cfg.RawTemplatesFolder) Then _
        Exit Sub
    
    Dim PBar As ProgressBar: Set PBar = _
        ProgressBar.New_(ProgressBarNumeric, RawTemplateFiles.Count)
    PBar.Cancelable = True
    
    Dim File As Variant
    For Each File In RawTemplateFiles
        PBar.Update
        PrepareImageAndSave File, Cfg
        If PBar.Canceled Then Exit Sub
    Next File
End Sub

Private Sub PrepareImageAndSave( _
                ByVal TemplateFile As String, ByVal Cfg As Config _
            )
    CreateDocument
    Dim Template As Shape
    Dim Frame As Shape
    Dim RectangleSize As Double: RectangleSize _
        = PixelsToDocUnits(RECTANGLE_SIZE_PX)
    With ActiveDocument
        .ColorContext.BlendingColorModel = clrColorModelRGB
        .ActiveLayer.Import TemplateFile
        Set Template = ActiveShape
        ResizeImageToDocumentResolution Template
        Template.CenterX = .ActivePage.CenterX
        Template.CenterY = .ActivePage.CenterY
        Set Frame = _
            .ActiveLayer.CreateRectangle2( _
                0, 0, RectangleSize, RectangleSize _
            )
        Frame.CenterX = .ActivePage.CenterX
        Frame.CenterY = .ActivePage.CenterY
        Frame.OrderFrontOf Template
        Frame.Fill.ApplyUniformFill CreateRGBColor(255, 255, 255)
        Frame.Outline.SetNoOutline
        Frame.CreateDropShadow _
            cdrDropShadowFlat, 50, 15, 0, 0, CreateRGBColor(0, 0, 0), _
            MergeMode:=cdrMergeMultiply
        
        Dim ExportFile As FileSpec: Set ExportFile = _
            FileSpec.New_(TemplateFile)
        ExportFile.Path = Cfg.PreparedTemplatesFolder
        ExportFile.Ext = "cdr"
        .SaveAs ExportFile
        .Close
    End With
End Sub

Private Function CheckImagesCount( _
                     ByVal Images As Collection, _
                     ByVal Folder As String _
                 ) As Boolean
    If Images.Count = 0 Then
        VBA.MsgBox "Не найдены изображения в папке " & Folder, vbCritical
    Else
        CheckImagesCount = True
    End If
End Function

Private Property Get GetValidImagesFromFolder( _
                         ByVal Folder As String _
                     ) As Collection
    With FSO
        Dim ValidFiles As New Collection
        Dim File As File
        For Each File In .GetFolder(Folder).Files
            If CheckFile(File.Name) Then
                ValidFiles.Add File
            End If
        Next File
    End With
    Set GetValidImagesFromFolder = ValidFiles
End Property

'-------------------------------------------------------------------------------

Private Sub ExportOnTemplates( _
                ByVal ImageFile As String, _
                ByVal ImageSize As Size, _
                ByVal SavePath As String, _
                ByVal Cfg As Config _
            )
    If ImageSize.Landscape Then
        ExportOnTemplatesSubset _
            Deduplicate( _
                GetRandomFilesFromFolder( _
                    Cfg.HTemplatesFolder, Cfg.ImagesQuantity _
                ) _
            ), _
            Cfg.HWidth, ImageFile, ImageSize, SavePath, Cfg
    ElseIf ImageSize.Portrait Then
        ExportOnTemplatesSubset _
            Deduplicate( _
                GetRandomFilesFromFolder( _
                    Cfg.VTemplatesFolder, Cfg.ImagesQuantity _
                ) _
            ), _
            Cfg.VHeight, ImageFile, ImageSize, SavePath, Cfg
    Else
        ExportOnTemplatesSubset _
            Deduplicate( _
                GetRandomFilesFromFolder( _
                    Cfg.ETemplatesFolder, Cfg.ImagesQuantity _
                ) _
            ), _
            Cfg.ESize, ImageFile, ImageSize, SavePath, Cfg
    End If
End Sub

Private Sub ExportOnTemplatesSubset( _
                ByVal TemplateFiles As Collection, _
                ByVal TemplateLongestSide As Double, _
                ByVal ImageFile As String, _
                ByVal ImageSize As Size, _
                ByVal SavePath As String, _
                ByVal Cfg As Config _
            )
    Dim File As Scripting.File
    For Each File In TemplateFiles
        SetOnTemplateAndExport _
            FileSpec.New_(File.Path), TemplateLongestSide, _
            ImageFile, ImageSize, SavePath, Cfg
    Next File
End Sub

Private Sub SetOnTemplateAndExport( _
                ByVal TemplateFile As FileSpec, _
                ByVal TemplateLongestSide As Double, _
                ByVal ImageFile As String, _
                ByVal ImageSize As Size, _
                ByVal SavePath As String, _
                ByVal Cfg As Config _
            )
    OpenDocument TemplateFile
    With ActiveDocument
        .ReferencePoint = cdrCenter
        .Unit = cdrCentimeter
        Dim ImageSizeCm As Size: Set ImageSizeCm = _
            ImageSize.ConvertUnits(Cfg.Unit, cdrCentimeter)
        
        .ActiveLayer.Import ImageFile
        Dim Image As Shape: Set Image = ActiveShape
        Dim Frame As Shape: Set Frame = GetFrames(1)
                
        Dim TempToImageRatio As Double: TempToImageRatio = _
            TemplateLongestSide / ImageSizeCm.Longest
          
        ImageSize.ResizeToLongest( _
            Size.NewFromShape(Frame).Longest / TempToImageRatio _
        ).ApplyToShape Frame
        
        Image.SetSize Frame.SizeWidth, Frame.SizeHeight
        Image.CenterX = Frame.CenterX
        Image.CenterY = Frame.CenterY
        Image.OrderFrontOf Frame
                
        Dim File As FileSpec: Set File = FileSpec.New_(TemplateFile)
        File.Path = SavePath
        
        If Cfg.OptionJpeg Then
            ExportTemplateAsJpeg File
        ElseIf Cfg.OptionPng Then
            ExportTemplateAsPng File
        End If
        .Close
    End With
End Sub

Private Property Get GetFrames() As Collection
    Set GetFrames = New Collection
    Dim Shape As Shape
    For Each Shape In ActivePage.Shapes
        If Shape.Type = cdrRectangleShape Then
            GetFrames.Add Shape
        End If
    Next Shape
End Property

Private Sub ExportTemplateAsPng(ByVal File As FileSpec)
    File.Ext = "png"
    With ActiveDocument
        .ExportBitmap( _
            File, cdrPNG, cdrCurrentPage, cdrRGBColorImage, , , _
            .ResolutionX, .ResolutionX _
        ).Finish
    End With
End Sub

Private Sub ExportTemplateAsJpeg(ByVal File As FileSpec)
    File.Ext = "jpg"
    With ActiveDocument
        With _
            .ExportBitmap( _
                File, cdrJPEG, cdrCurrentPage, cdrRGBColorImage, , , _
                .ResolutionX, .ResolutionX _
            )
            .Compression = 10
            .Optimized = True
            .Finish
        End With
    End With
End Sub

Private Property Get RandomizeSize( _
                ByVal Shape As Shape, _
                ByVal Cfg As Config _
            ) As Size
    Dim ImageSize As Size: Set ImageSize = Size.NewFromRect(Shape.BoundingBox)
    Dim Ratio As Double: Ratio = ImageSize.Longest / ImageSize.Shortest
    
    Dim LongestSide As Double, ShortestSide As Double
    LongestSide = RndDouble(Cfg.MinWidth, Cfg.MaxWidth)
    ShortestSide = LongestSide / Ratio
    
    If ImageSize.Landscape Then
        Set RandomizeSize = Size.New_(LongestSide, ShortestSide)
    Else
        Set RandomizeSize = Size.New_(ShortestSide, LongestSide)
    End If
End Property

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
        .OptionImageOnRandomTemplate = Cfg.OptionImageOnRandomTemplate
        .ImagesQuantity = Cfg.ImagesQuantity
        .OptionPng = Cfg.OptionPng
        .OptionJpeg = Cfg.OptionJpeg
        .HTemplatesFolder = Cfg.HTemplatesFolder
        .HWidth = Cfg.HWidth
        .VTemplatesFolder = Cfg.VTemplatesFolder
        .VHeight = Cfg.VHeight
        .ETemplatesFolder = Cfg.ETemplatesFolder
        .ESize = Cfg.ESize
        
        .Show vbModal
        
        Cfg.SourceFolder = .SourceFolder
        Cfg.OutputFolder = .OutputFolder
        Cfg.DivWidth = .DivWidth
        Cfg.DivHeight = .DivHeight
        Cfg.MinWidth = .MinWidth
        Cfg.MaxWidth = .MaxWidth
        Cfg.OptionInches = .OptionInches
        Cfg.OptionCentimeters = .OptionCentimeters
        Cfg.OptionImageOnRandomTemplate = .OptionImageOnRandomTemplate
        Cfg.ImagesQuantity = .ImagesQuantity
        Cfg.OptionPng = .OptionPng
        Cfg.OptionJpeg = .OptionJpeg
        Cfg.HTemplatesFolder = .HTemplatesFolder
        Cfg.HWidth = .HWidth
        Cfg.VTemplatesFolder = .VTemplatesFolder
        Cfg.VHeight = .VHeight
        Cfg.ETemplatesFolder = .ETemplatesFolder
        Cfg.ESize = .ESize
        
        ShowViewAndGetConfig = .IsOk
    End With
End Function

Private Function ShowPreprocessorViewAndGetConfig(ByVal Cfg As Config) As Boolean
    With New PreprocessorView
        .RawTemplatesFolder = Cfg.RawTemplatesFolder
        .PreparedTemplatesFolder = Cfg.PreparedTemplatesFolder
        
        .Show vbModal
        
        Cfg.RawTemplatesFolder = .RawTemplatesFolder
        Cfg.PreparedTemplatesFolder = .PreparedTemplatesFolder
        
        ShowPreprocessorViewAndGetConfig = .IsOk
    End With
End Function

'===============================================================================
' # Tests

Private Sub testSomething()
    Size.New_(3, 3).ApplyToShape ActiveSelectionRange.FirstShape
End Sub
