Imports System.Net
Imports System.Security.Policy
Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Microsoft.VisualStudio.TestTools.UnitTesting.Logging
Imports NLog
Imports PdfSharp
Imports PdfSharp.Drawing
Imports TDSprint2
Imports TheArtOfDev.HtmlRenderer.Adapters
Imports TheArtOfDev.HtmlRenderer.Core
Imports TheArtOfDev.HtmlRenderer.PdfSharp
'Imports PageSize = PdfSharp.PageSize
Imports PageSize = TDSprint2.PageSize

<TestClass()> Public Class UnitTest1
    Friend Request As MyWebClient
    Friend Shared Log As NLog.Logger = LogManager.GetCurrentClassLogger()

    <TestInitialize>
    Public Sub Start()
        Request = New MyWebClient
    End Sub


    <TestMethod()>
    <DataRow("H:\TMP2\tst2.htm")>
    Public Sub TestClass(FileName As String)
        Try
            Dim Response As String = IO.File.ReadAllText(FileName)
            Dim TmpDir As String = System.IO.Path.GetTempPath
            Dim TmpFile As String = IO.Path.Combine(TmpDir, Replace(FileName, ".htm", ".pdf"))
            Dim X As New Html2Pdf
            Dim Config As New TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerateConfig With {
                    .MarginBottom = 10,
                    .MarginLeft = 10,
                    .MarginRight = 10,
                    .MarginTop = 10,
                    .PageOrientation = 0,
                    .PageSize = 4
            }
            Dim Css As CssData = X.ParseStyleSheet("", True)
            X.CreatePdfFile3(TmpFile, Response, 4, 10, Css)
            Dim Ret2 = Process.Start("C:\Program Files\Mozilla Firefox\firefox.exe", TmpFile)
        Catch ex As WebException
            Dim Resp As String = ""
            Dim Stream = ex.Response?.GetResponseStream()
            If Stream IsNot Nothing Then
                Dim Sr = New IO.StreamReader(Stream)
                Resp = Sr.ReadToEnd
            End If
            Log.Info(Resp & vbCrLf & ex.Message)
        End Try
    End Sub

    <DataRow("H:\TMP2\tst2.htm")>
    <TestMethod()>
    Public Sub TestCOM(FileName As String)
        Try
            Dim Response As String = IO.File.ReadAllText(FileName)
            Dim TmpDir As String = System.IO.Path.GetTempPath
            Dim TmpFile As String = IO.Path.Combine(TmpDir, Replace(FileName, ".htm", ".pdf"))
            '
            Dim Config As New TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerateConfig With {
                    .MarginBottom = 10,
                    .MarginLeft = 10,
                    .MarginRight = 10,
                    .MarginTop = 10,
                    .PageOrientation = 0,
                    .PageSize = 4
            }
            Dim ComObjType = Type.GetTypeFromProgID("TDS.Html2Pdf")
            Dim X As IHtml2Pdf = Activator.CreateInstance(ComObjType)
            X.CreatePdfFile2(TmpFile, Response, 4, 10)
            Dim Ret2 = Process.Start("C:\Program Files\Mozilla Firefox\firefox.exe", TmpFile)
        Catch ex As WebException
            Dim Resp As String = ""
            Dim Stream = ex.Response?.GetResponseStream()
            If Stream IsNot Nothing Then
                Dim Sr = New IO.StreamReader(Stream)
                Resp = Sr.ReadToEnd
            End If
            Log.Info(Resp & vbCrLf & ex.Message)
        End Try
    End Sub

End Class