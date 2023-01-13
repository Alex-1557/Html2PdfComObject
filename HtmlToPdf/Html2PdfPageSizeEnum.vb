Imports System.Runtime.InteropServices

<ComVisible(True)>
<Guid("1b7ef50f-3526-4a95-8c01-69e12fd053d4")>
<ClassInterface(ClassInterfaceType.None)>
<ProgId("TDS.Html2PdfPageSizeEnum")>
Public Class Html2PdfPageSizeEnum
    Private _PageSize As TDSprint2.PageSize = TDSprint2.PageSize.A4
    Public Property PageSize As TDSprint2.PageSize
        Get
            Return _PageSize
        End Get
        Set(value As TDSprint2.PageSize)
            _PageSize = value
        End Set
    End Property
End Class
