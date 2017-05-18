
Partial Class M_Source_09_Download
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim f As String = Server.MapPath(Request.QueryString("f"))
        Dim wc As New System.Net.WebClient()
        Dim a As Byte() = wc.DownloadData(f)
        Dim aa As String = System.IO.Path.GetFileName(f)
        Dim fileName As String = "attachment" + System.IO.Path.GetExtension(f)
        Response.ContentType = "application/octet-stream"
        Response.HeaderEncoding = Encoding.GetEncoding("big5")
        Response.AddHeader("Content-Disposition", String.Format("attachment;filename={0}", fileName))
        Response.BinaryWrite(a)
    End Sub

End Class
