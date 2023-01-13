Module Start
    Public Sub Main()
        Dim Str1 = IO.File.ReadAllText("H:\TMP2\ATI.htm")
        Dim X As New HTML2PDF
        X.CreatePdfFile("H:\TMP2\ATI.pdf", Str1, PageSize.A1, 10)
        Dim Ret1 = Process.Start("C:\Program Files\Mozilla Firefox\firefox.exe", "H:\TMP2\ATI.pdf")
    End Sub
End Module
