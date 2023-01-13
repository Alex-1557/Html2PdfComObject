# Html2PdfComObject
PDF creator COM-Object based on TheArtOfDev.HtmlRenderer.PdfSharp     
This COM-object allow to create PDF on anchient ASP pages, this is workable example:

     <!DOCTYPE html>
     <html>
     <body>
     <form name='TDS' method='post' ID="test1" action="/TSTPDF2.asp">
     <a href="#">Test PDF component.</a> <br><br>
     Html file<br>
     <input type="text" name="html" style="width: 500px;" value="H:\TMP2\tst2.htm">
     <br>
     <input type="submit" value="Create PDF">
     </form>
     <%
     Const ForReading = 1
     Const ForWriting = 2
     Const ForAppending = 8

     If IsPostBack() then
     Dim PdfCretor, FS, File 'as object 
     Dim TmpDir, FileContent, GUID, TmpFileName 'as string

     Set FS = Server.CreateObject("Scripting.FileSystemObject")
     Set File = FS.OpenTextFile(Request.Form("html"), ForReading)
     FileContent = File.ReadAll
     File.Close

     TmpDir = Server.MapPath("temp")
     GUID = CreateGUID()
     TmpFileName = TmpDir & "\" & GUID & ".pdf"

     Set PdfCretor = Server.CreateObject("TDS.HTML2PDF")
     PdfCretor.CreatePdfFile2 TmpFileName, FileContent, 4,10

     Response.Redirect("/temp/" & GUID & ".pdf")

     Set PdfCretor = nothing
     Set FS = nothing
     Set File = nothing
     end if

     Function CreateGUID
     Dim TypeLib
     Set TypeLib = CreateObject("Scriptlet.TypeLib")
     CreateGUID = Mid(TypeLib.Guid, 2, 36)
     End Function

     Function IsPostBack()
     If Request.ServerVariables("Request_Method") = "POST" Then
        IsPostBack="True"
     Else
        IsPostBack="False"
     End If
     End Function
     %>
     </body>
     </html>
