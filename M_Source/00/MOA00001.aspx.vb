
Partial Class Source_00_MOA00001
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String
    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        '隱藏eformid
        e.Row.Cells(0).Visible = False
    End Sub


    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

        Dim streformid As String = ""

        Dim eformid As String = GridView1.Rows(GridView1.SelectedIndex).Cells(0).Text
        Dim eformName As String = GridView1.Rows(GridView1.SelectedIndex).Cells(1).Text

        streformid = eformid & ",MOA," & eformName & ",1"
        Server.Transfer("MOA00003.aspx?eformid=" + streformid)

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            user_id = Session("user_id")
            org_uid = Session("ORG_UID")

            'session被清空回首頁
            If user_id = "" Or org_uid = "" Then

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")

            Else

                '判斷登入者權限
                Dim LoginCheck As New C_Public

                If LoginCheck.LoginCheck(user_id, "MOA00001") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00001.aspx")
                    Response.End()
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub ImgBtn1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgBtn1.Click

        Server.Transfer("MOA00002.aspx")

    End Sub
End Class
