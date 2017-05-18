Imports System.Web.UI
Imports System.ComponentModel
Imports System.Web.UI.HtmlControls
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.IO
Imports System.Data

Partial Class M_Source_04_MOA04102
    Inherits System.Web.UI.Page
    Public printdate As String = DateTime.Now.ToString("yyyy/MM/dd")
    Dim user_id, org_uid As String
    Public EFORMSN As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        user_id = Session("user_id")
        org_uid = Session("ORG_UID")

        'Session被清空回首頁
        If user_id = "" Or org_uid = "" Then

            Response.Write("<script language='javascript'>")
            Response.Write("alert('畫面停留太久未使用，將回首頁重新整理');")
            Response.Write("window.close();")
            Response.Write("</script>")

        Else

            '判斷登入者權限
            Dim LoginCheck As New C_Public

            If LoginCheck.LoginCheck(user_id, "MOA04100") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA04100.aspx")
                Response.End()
            End If

        End If


        If CType(Request.QueryString("streformsn"), String) Is Nothing Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('您的參數已遺失，請重新操作，感謝您！');")
            Response.Write("window.close();")
            Response.Write("</script>")
        Else
            EFORMSN = CType(Request.QueryString("streformsn"), String)
            If EFORMSN.Length <> 16 Then
                Response.Write("<script language='javascript'>")
                Response.Write("alert('您的報修單號錯誤:<" + EFORMSN + ">，請重新操作，感謝您！');")
                Response.Write("window.close();")
                Response.Write("</script>")
            Else
                lb_EFORMSN.Text = EFORMSN
                Dim name As String = "attachment;filename=房舍水電報修單_" + EFORMSN + ".doc"
                name = Server.UrlPathEncode(name)
                Response.Clear()
                Response.AddHeader("content-disposition", name)
                Response.Charset = "UTF-8"
                Response.ContentType = "application/vnd.ms-word"

                Dim stringWrite As System.IO.StringWriter = New System.IO.StringWriter
                Dim htmlWrite As System.Web.UI.HtmlTextWriter = New HtmlTextWriter(stringWrite)
                Dim hf As New HtmlForm
                Controls.Add(hf)
                hf.Controls.Add(PanelContent)
                hf.RenderControl(htmlWrite)
                Response.Write(stringWrite.ToString)
                Response.End()
            End If
        End If
    End Sub
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)


    End Sub
End Class
