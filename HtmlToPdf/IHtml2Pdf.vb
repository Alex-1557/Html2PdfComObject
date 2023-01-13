
Imports PdfSharp.Pdf
Imports TheArtOfDev.HtmlRenderer.Core
Imports TheArtOfDev.HtmlRenderer.Core.Entities
Imports TheArtOfDev.HtmlRenderer.PdfSharp

<Runtime.InteropServices.ComVisible(True)>
<Runtime.InteropServices.Guid("864f3ae4-7c66-11ed-9649-5254008de3b3")>
<Runtime.InteropServices.InterfaceType(Runtime.InteropServices.ComInterfaceType.InterfaceIsDual)>
Public Interface IHtml2Pdf
    Sub AddFontFamilyMapping(FromFamily As String, ToFamily As String)
    Function ParseStyleSheet(Stylesheet As String, Optional combineWithDefault As Boolean = True) As CssData
    Function CreatePdfConfig(Optional MarginBottom As Integer = 10,
                    Optional MarginLeft As Integer = 10,
                    Optional MarginRight As Integer = 10,
                    Optional MarginTop As Integer = 10,
                    Optional PageOrientation As Integer = 0,
                    Optional PageSize As Integer = 4) As TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerateConfig
    Sub CreatePdfFile1(PdfFileName As String, Html As String, PageSize As PageSize)
    Sub CreatePdfFile2(PdfFileName As String, Html As String, PageSize As PageSize, Margin As Integer)
    Sub CreatePdfFile3(PdfFileName As String, Html As String, PageSize As PageSize, Margin As Integer,
                  Optional CssData As TheArtOfDev.HtmlRenderer.Core.CssData = Nothing,
                  Optional StylesSheetLoad As EventHandler(Of TheArtOfDev.HtmlRenderer.Core.Entities.HtmlStylesheetLoadEventArgs) = Nothing,
                  Optional ImageLoad As EventHandler(Of TheArtOfDev.HtmlRenderer.Core.Entities.HtmlImageLoadEventArgs) = Nothing)
    Sub CreatePdfFile4(PdfFileName As String, Html As String, Config As TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerateConfig)
    Sub CreatePdfFile5(PdfFileName As String, Html As String, Config As TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerateConfig,
                  Optional CssData As TheArtOfDev.HtmlRenderer.Core.CssData = Nothing,
                  Optional StylesSheetLoad As EventHandler(Of TheArtOfDev.HtmlRenderer.Core.Entities.HtmlStylesheetLoadEventArgs) = Nothing,
                  Optional ImageLoad As EventHandler(Of TheArtOfDev.HtmlRenderer.Core.Entities.HtmlImageLoadEventArgs) = Nothing)
    Function CreatePdfDoc1(Html As String, PageSize As PageSize) As PdfDocument
    Function CreatePdfDoc2(Html As String, PageSize As PageSize, Margin As Integer) As PdfDocument
    Function CreatePdfDoc3(Html As String, PageSize As PageSize, Margin As Integer,
                  Optional CssData As TheArtOfDev.HtmlRenderer.Core.CssData = Nothing,
                  Optional StylesSheetLoad As EventHandler(Of TheArtOfDev.HtmlRenderer.Core.Entities.HtmlStylesheetLoadEventArgs) = Nothing,
                  Optional ImageLoad As EventHandler(Of TheArtOfDev.HtmlRenderer.Core.Entities.HtmlImageLoadEventArgs) = Nothing) As PdfDocument
    Function CreatePdfDoc4(Html As String, Config As TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerateConfig) As PdfDocument
    Function CreatePdfDoc5(Html As String, Config As TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerateConfig,
                  Optional CssData As TheArtOfDev.HtmlRenderer.Core.CssData = Nothing,
                  Optional StylesSheetLoad As EventHandler(Of TheArtOfDev.HtmlRenderer.Core.Entities.HtmlStylesheetLoadEventArgs) = Nothing,
                  Optional ImageLoad As EventHandler(Of TheArtOfDev.HtmlRenderer.Core.Entities.HtmlImageLoadEventArgs) = Nothing) As PdfDocument
    Sub AddPdfPageToDoc1(Document As PdfDocument, Html As String, PageSize As PageSize,
                          Optional margin As Integer = 20,
                          Optional CssData As CssData = Nothing,
                          Optional StylesheetLoad As EventHandler(Of HtmlStylesheetLoadEventArgs) = Nothing,
                          Optional ImageLoad As EventHandler(Of HtmlImageLoadEventArgs) = Nothing)
    Sub AddPdfPageToDoc2(Document As PdfDocument, Html As String, Config As PdfGenerateConfig,
                     Optional CssData As CssData = Nothing,
                     Optional StylesheetLoad As EventHandler(Of HtmlStylesheetLoadEventArgs) = Nothing,
                     Optional ImageLoad As EventHandler(Of HtmlImageLoadEventArgs) = Nothing)
    Sub AddPdfPageToFile1(PdfFileName As String, Document As PdfDocument, Html As String, PageSize As PageSize,
                          Optional margin As Integer = 20,
                          Optional CssData As CssData = Nothing,
                          Optional StylesheetLoad As EventHandler(Of HtmlStylesheetLoadEventArgs) = Nothing,
                          Optional ImageLoad As EventHandler(Of HtmlImageLoadEventArgs) = Nothing)

    Sub AddPdfPageToFile2(PdfFileName As String, Document As PdfDocument, Html As String, Config As PdfGenerateConfig,
                     Optional CssData As CssData = Nothing,
                     Optional StylesheetLoad As EventHandler(Of HtmlStylesheetLoadEventArgs) = Nothing,
                     Optional ImageLoad As EventHandler(Of HtmlImageLoadEventArgs) = Nothing)
End Interface
