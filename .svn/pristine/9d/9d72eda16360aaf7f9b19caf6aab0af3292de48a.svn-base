
Partial Class M_Source_04_MOA04019_1
    Inherits System.Web.UI.Page
    Public it_code As String
    Public sPicPath As String = ConfigurationManager.AppSettings("PicPathUrl")

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If CType(Request.QueryString("shcode"), String) Is Nothing Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('您的參數已遺失，請重新操作，感謝您！');")
            Response.Write("window.close();")
            Response.Write("</script>")
        Else
            Dim it_code As String = CType(Request.QueryString("shcode"), String)
            If it_code.Length <> 6 Then
                Response.Write("<script language='javascript'>")
                Response.Write("alert('您的參數不正確，請重新操作，感謝您！');")
                Response.Write("window.close();")
                Response.Write("</script>")
            End If
            lb_it_code1.Text = it_code
        End If
    End Sub

    Public Function showPic(ByVal filename As String) As Boolean
        If filename.Length <= 3 Then
            Return False
        Else
            Return True
        End If
    End Function


    Public Function showPicLB(ByVal filename As String) As Boolean
        If filename.Length <= 3 Then
            Return True
        Else
            Return False
        End If
    End Function
End Class
