Imports System.Net.WebRequestMethods
Imports PdfSharp.Pdf
Imports TheArtOfDev.HtmlRenderer.Core
Imports TheArtOfDev.HtmlRenderer.Core.Entities
Imports TheArtOfDev.HtmlRenderer.PdfSharp

<Runtime.InteropServices.ComVisible(True)>
<Runtime.InteropServices.Guid("37191c45-7c66-11ed-9649-5254008de3b3")>
<Runtime.InteropServices.ClassInterface(Runtime.InteropServices.ClassInterfaceType.AutoDispatch)>
<Runtime.InteropServices.ProgId("TDS.Html2Pdf")>
Public Class Html2Pdf
    Implements IHtml2Pdf

    ''' <summary>
    ''' Create new object PdfGenerateConfig
    ''' </summary>
    ''' <param name="PageOrientation">0/1 - vertical/horisontal </param>
    ''' <param name="PageSize">0/1/2/3/4 - A0/A1/A2/A4/A4</param> 
    ''' <returns>new PdfGenerateConfig</returns>
    Function CreatePdfConfig(Optional MarginBottom As Integer = 10,
                    Optional MarginLeft As Integer = 10,
                    Optional MarginRight As Integer = 10,
                    Optional MarginTop As Integer = 10,
                    Optional PageOrientation As Integer = 0,
                    Optional PageSize As Integer = 4) As TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerateConfig Implements IHtml2Pdf.CreatePdfConfig
        Return New PdfGenerateConfig With {
                                        .MarginBottom = MarginBottom,
                                        .MarginLeft = MarginLeft,
                                        .MarginRight = MarginRight,
                                        .MarginTop = MarginTop,
                                        .PageOrientation = PageOrientation,
                                        .PageSize = PageSize
                                            }
    End Function

    ''' <summary>
    ''' Adds a font mapping from <paramref name="fromFamily"/> to <paramref name="toFamily"/> iff the <paramref name="fromFamily"/> is not found.
    ''' When the <paramref name="fromFamily"/> font is used in rendered html and is not found in existing 
    ''' fonts (installed or added) it will be replaced by <paramref name="toFamily"/>
    ''' </summary>
    ''' <remarks>
    ''' This fonts mapping can be used as a fallback in case the requested font is not installed in the client system.
    ''' </remarks>
    ''' <param name="FromFamily">the font family to replace</param>
    ''' <param name="ToFamily">the font family to replace with</param>
    Public Sub AddFontFamilyMapping(FromFamily As String, ToFamily As String) Implements IHtml2Pdf.AddFontFamilyMapping
        PdfGenerator.AddFontFamilyMapping(FromFamily, ToFamily)
    End Sub

    ''' <summary>
    ''' Create PDF file from given HTML.
    ''' </summary>
    ''' <param name="PdfFileName">Path to store PDF</param>
    ''' <param name="Html">HTML source to create PDF from</param>
    ''' <param name="PageSize">the page size to use for each page in the generated pdf</param>
    Public Sub CreatePdfFile1(PdfFileName As String, Html As String, PageSize As PageSize) Implements IHtml2Pdf.CreatePdfFile1
        PdfGenerator.GeneratePdf(Html, New PdfSharp.PageSize With {.value__ = PageSize}).Save(PdfFileName)
    End Sub

    ''' <summary>
    ''' Create PDF file from given HTML.
    ''' </summary>
    ''' <param name="PdfFileName">Path to store PDF</param>
    ''' <param name="Html">HTML source to create PDF from</param>
    ''' <param name="PageSize">the page size to use for each page in the generated pdf</param>
    ''' <param name="Margin">the margin to use between the HTML and the edges of each page</param>
    Public Sub CreatePdfFile2(PdfFileName As String, Html As String, PageSize As PageSize, Margin As Integer) Implements IHtml2Pdf.CreatePdfFile2
        PdfGenerator.GeneratePdf(Html, New PdfSharp.PageSize With {.value__ = PageSize}, Margin).Save(PdfFileName)
    End Sub

    ''' <summary>
    ''' Create PDF file from given HTML.
    ''' </summary>
    ''' <param name="PdfFileName">Path to store PDF</param>
    ''' <param name="Html">HTML source to create PDF from</param>
    ''' <param name="PageSize">the page size to use for each page in the generated pdf</param>
    ''' <param name="Margin">the margin to use between the HTML and the edges of each page</param>
    ''' <param name="CssData">optional: the style to use for html rendering (default - use W3 default style)</param>
    ''' <param name="StylesSheetLoad">optional: can be used to overwrite stylesheet resolution logic</param>
    ''' <param name="ImageLoad">optional: can be used to overwrite image resolution logic</param>
    Public Sub CreatePdfFile3(PdfFileName As String, Html As String, PageSize As PageSize, Margin As Integer, Optional CssData As CssData = Nothing, Optional StylesSheetLoad As EventHandler(Of HtmlStylesheetLoadEventArgs) = Nothing, Optional ImageLoad As EventHandler(Of HtmlImageLoadEventArgs) = Nothing) Implements IHtml2Pdf.CreatePdfFile3
        PdfGenerator.GeneratePdf(Html, New PdfSharp.PageSize With {.value__ = PageSize}, Margin, CssData, StylesSheetLoad, ImageLoad).Save(PdfFileName)
    End Sub

    ''' <summary>
    ''' Create PDF file from given HTML.
    ''' </summary>
    ''' <param name="PdfFileName">Path to store PDF</param>
    ''' <param name="Html">HTML source to create PDF from</param>
    ''' <param name="Config">the configuration to use for the PDF generation (page size/page orientation/margins/etc.)</param>
    Public Sub CreatePdfFile4(PdfFileName As String, Html As String, Config As PdfGenerateConfig) Implements IHtml2Pdf.CreatePdfFile4
        PdfGenerator.GeneratePdf(Html, Config).Save(PdfFileName)
    End Sub

    ''' <summary>
    ''' Create PDF file from given HTML.
    ''' </summary>
    ''' <param name="PdfFileName">Path to store PDF</param>
    ''' <param name="Html">HTML source to create PDF from</param>
    ''' <param name="Config">the configuration to use for the PDF generation (page size/page orientation/margins/etc.)</param>
    ''' <param name="CssData">optional: the style to use for html rendering (default - use W3 default style)</param>
    ''' <param name="StylesSheetLoad">optional: can be used to overwrite stylesheet resolution logic</param>
    ''' <param name="ImageLoad">optional: can be used to overwrite image resolution logic</param>
    Public Sub CreatePdfFile5(PdfFileName As String, Html As String, Config As PdfGenerateConfig, Optional CssData As CssData = Nothing, Optional StylesSheetLoad As EventHandler(Of HtmlStylesheetLoadEventArgs) = Nothing, Optional ImageLoad As EventHandler(Of HtmlImageLoadEventArgs) = Nothing) Implements IHtml2Pdf.CreatePdfFile5
        PdfGenerator.GeneratePdf(Html, Config, CssData, StylesSheetLoad, ImageLoad).Save(PdfFileName)
    End Sub

    ''' <summary>
    ''' Create PDF pages from given HTML and appends them to the provided PDF document.
    ''' </summary>
    ''' <param name="Document">PDF document to append pages to</param>
    ''' <param name="Html">HTML source to create PDF from</param>
    ''' <param name="PageSize">the page size to use for each page in the generated pdf </param>
    ''' <param name="margin">the margin to use between the HTML and the edges of each page</param>
    ''' <param name="CssData">optional: the style to use for html rendering (default - use W3 default style)</param>
    ''' <param name="StylesheetLoad">optional: can be used to overwrite stylesheet resolution logic</param>
    ''' <param name="ImageLoad">optional: can be used to overwrite image resolution logic</param>
    Public Sub AddPdfPageToDoc1(Document As PdfDocument, Html As String, PageSize As PageSize, Optional margin As Integer = 20, Optional CssData As CssData = Nothing, Optional StylesheetLoad As EventHandler(Of HtmlStylesheetLoadEventArgs) = Nothing, Optional ImageLoad As EventHandler(Of HtmlImageLoadEventArgs) = Nothing) Implements IHtml2Pdf.AddPdfPageToDoc1
        PdfGenerator.AddPdfPages(Document, Html, PageSize, margin, CssData, StylesheetLoad, ImageLoad)
    End Sub

    ''' <summary>
    ''' Create PDF pages from given HTML and appends them to the provided PDF document.
    ''' </summary>
    ''' <param name="Document">PDF document to append pages to</param>
    ''' <param name="Html">HTML source to create PDF from</param>
    ''' <param name="Config">the configuration to use for the PDF generation (page size/page orientation/margins/etc.)</param>
    ''' <param name="CssData">optional: the style to use for html rendering (default - use W3 default style)</param>
    ''' <param name="StylesheetLoad">optional: can be used to overwrite stylesheet resolution logic</param>
    ''' <param name="ImageLoad">optional: can be used to overwrite image resolution logic</param>
    Public Sub AddPdfPageToDoc2(Document As PdfDocument, Html As String, Config As PdfGenerateConfig, Optional CssData As CssData = Nothing, Optional StylesheetLoad As EventHandler(Of HtmlStylesheetLoadEventArgs) = Nothing, Optional ImageLoad As EventHandler(Of HtmlImageLoadEventArgs) = Nothing) Implements IHtml2Pdf.AddPdfPageToDoc2
        PdfGenerator.AddPdfPages(Document, Html, Config, CssData, StylesheetLoad, ImageLoad)
    End Sub

    ''' <summary>
    ''' Create PDF pages from given HTML and appends them to the provided PDF document.
    ''' </summary>
    ''' <param name="PdfFileName">Path to store PDF</param>
    ''' <param name="Document">PDF document to append pages to</param>
    ''' <param name="Html">HTML source to create PDF from</param>
    ''' <param name="PageSize">the page size to use for each page in the generated pdf </param>
    ''' <param name="margin">the margin to use between the HTML and the edges of each page</param>
    ''' <param name="CssData">optional: the style to use for html rendering (default - use W3 default style)</param>
    ''' <param name="StylesheetLoad">optional: can be used to overwrite stylesheet resolution logic</param>
    ''' <param name="ImageLoad">optional: can be used to overwrite image resolution logic</param>
    Public Sub AddPdfPageToFile1(PdfFileName As String, Document As PdfDocument, Html As String, PageSize As PageSize, Optional margin As Integer = 20, Optional CssData As CssData = Nothing, Optional StylesheetLoad As EventHandler(Of HtmlStylesheetLoadEventArgs) = Nothing, Optional ImageLoad As EventHandler(Of HtmlImageLoadEventArgs) = Nothing) Implements IHtml2Pdf.AddPdfPageToFile1
        PdfGenerator.AddPdfPages(Document, Html, PageSize, margin, CssData, StylesheetLoad, ImageLoad)
        Document.Save(PdfFileName)
    End Sub

    ''' <summary>
    ''' Create PDF pages from given HTML and appends them to the provided PDF document.
    ''' </summary>
    ''' <param name="PdfFileName">Path to store PDF</param>
    ''' <param name="Document">PDF document to append pages to</param>
    ''' <param name="Html">HTML source to create PDF from</param>
    ''' <param name="Config">the configuration to use for the PDF generation (page size/page orientation/margins/etc.)</param>
    ''' <param name="CssData">optional: the style to use for html rendering (default - use W3 default style)</param>
    ''' <param name="StylesheetLoad">optional: can be used to overwrite stylesheet resolution logic</param>
    ''' <param name="ImageLoad">optional: can be used to overwrite image resolution logic</param>
    Public Sub AddPdfPageToFile2(PdfFileName As String, Document As PdfDocument, Html As String, Config As PdfGenerateConfig, Optional CssData As CssData = Nothing, Optional StylesheetLoad As EventHandler(Of HtmlStylesheetLoadEventArgs) = Nothing, Optional ImageLoad As EventHandler(Of HtmlImageLoadEventArgs) = Nothing) Implements IHtml2Pdf.AddPdfPageToFile2
        PdfGenerator.AddPdfPages(Document, Html, Config, CssData, StylesheetLoad, ImageLoad)
        Document.Save(PdfFileName)
    End Sub

    ''' <summary>
    ''' Parse the given stylesheet to <see cref="CssData"/> object.
    ''' If <paramref name="combineWithDefault"/> is true the parsed css blocks are added to the 
    ''' default css data (as defined by W3), merged if class name already exists. If false only the data in the given stylesheet is returned.
    ''' </summary>
    ''' <param name="Stylesheet">the stylesheet source to parse</param>
    ''' <param name="combineWithDefault">true - combine the parsed css data with default css data, false - return only the parsed css data</param>
    ''' <returns>
    ''' the parsed css data
    ''' </returns>
    Public Function ParseStyleSheet(Stylesheet As String, Optional combineWithDefault As Boolean = True) As CssData Implements IHtml2Pdf.ParseStyleSheet
        Return PdfGenerator.ParseStyleSheet(Stylesheet, combineWithDefault)
    End Function

    ''' <summary>
    ''' Create PDF document from given HTML.
    ''' </summary>
    ''' <param name="Html">HTML source to create PDF from</param>
    ''' <param name="PageSize">the page size to use for each page in the generated pdf</param>
    ''' <returns>the generated image of the html</returns>
    Public Function CreatePdfDoc1(Html As String, PageSize As PageSize) As PdfDocument Implements IHtml2Pdf.CreatePdfDoc1
        Return PdfGenerator.GeneratePdf(Html, New PdfSharp.PageSize With {.value__ = PageSize})
    End Function

    ''' <summary>
    ''' Create PDF document from given HTML.
    ''' </summary>
    ''' <param name="Html">HTML source to create PDF from</param>
    ''' <param name="PageSize">the page size to use for each page in the generated pdf</param>
    ''' <param name="Margin">the margin to use between the HTML and the edges of each page</param>
    ''' <returns>the generated image of the html</returns>
    Public Function CreatePdfDoc2(Html As String, PageSize As PageSize, Margin As Integer) As PdfDocument Implements IHtml2Pdf.CreatePdfDoc2
        Return PdfGenerator.GeneratePdf(Html, New PdfSharp.PageSize With {.value__ = PageSize}, Margin)
    End Function

    ''' <summary>
    ''' Create PDF document from given HTML.
    ''' </summary>
    ''' <param name="Html">HTML source to create PDF from</param>
    ''' <param name="PageSize">the page size to use for each page in the generated pdf</param>
    ''' <param name="Margin">the margin to use between the HTML and the edges of each page</param>
    ''' <param name="CssData">optional: the style to use for html rendering (default - use W3 default style)</param>
    ''' <param name="StylesSheetLoad">optional: can be used to overwrite stylesheet resolution logic</param>
    ''' <param name="ImageLoad">optional: can be used to overwrite image resolution logic</param>
    ''' <returns>the generated image of the html</returns>
    Public Function CreatePdfDoc3(Html As String, PageSize As PageSize, Margin As Integer, Optional CssData As CssData = Nothing, Optional StylesSheetLoad As EventHandler(Of HtmlStylesheetLoadEventArgs) = Nothing, Optional ImageLoad As EventHandler(Of HtmlImageLoadEventArgs) = Nothing) As PdfDocument Implements IHtml2Pdf.CreatePdfDoc3
        Return PdfGenerator.GeneratePdf(Html, New PdfSharp.PageSize With {.value__ = PageSize}, Margin, CssData, StylesSheetLoad, ImageLoad)
    End Function

    ''' <summary>
    ''' Create PDF document from given HTML.
    ''' </summary>
    ''' <param name="Html">HTML source to create PDF from</param>
    ''' <param name="Config">the configuration to use for the PDF generation (page size/page orientation/margins/etc.)</param>
    ''' <returns>the generated image of the html</returns>
    Public Function CreatePdfDoc4(Html As String, Config As PdfGenerateConfig) As PdfDocument Implements IHtml2Pdf.CreatePdfDoc4
        Return PdfGenerator.GeneratePdf(Html, Config)
    End Function

    ''' <summary>
    ''' Create PDF document from given HTML.
    ''' </summary>
    ''' <param name="Html">HTML source to create PDF from</param>
    ''' <param name="Config">the configuration to use for the PDF generation (page size/page orientation/margins/etc.)</param>
    ''' <param name="CssData">optional: the style to use for html rendering (default - use W3 default style)</param>
    ''' <param name="StylesSheetLoad">optional: can be used to overwrite stylesheet resolution logic</param>
    ''' <param name="ImageLoad">optional: can be used to overwrite image resolution logic</param>
    ''' <returns>the generated image of the html</returns>
    Public Function CreatePdfDoc5(Html As String, Config As PdfGenerateConfig, Optional CssData As CssData = Nothing, Optional StylesSheetLoad As EventHandler(Of HtmlStylesheetLoadEventArgs) = Nothing, Optional ImageLoad As EventHandler(Of HtmlImageLoadEventArgs) = Nothing) As PdfDocument Implements IHtml2Pdf.CreatePdfDoc5
        Return PdfGenerator.GeneratePdf(Html, Config, CssData, StylesSheetLoad, ImageLoad)
    End Function

End Class
