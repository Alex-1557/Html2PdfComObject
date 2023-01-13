Imports System.Net

Public Class MyWebClient
    Inherits WebClient
    Protected Overloads Function GetWebRequest(URL As Uri) As WebRequest
        Dim WebRequest = MyBase.GetWebRequest(URL)
        Debug.WriteLine(WebRequest.RequestUri)
        WebRequest.ContentType = "text/html"
        WebRequest.Timeout = Integer.MaxValue
        Return WebRequest
    End Function

End Class