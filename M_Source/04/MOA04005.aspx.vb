Partial Class Source_04_MOA04005
    Inherits System.Web.UI.Page

    Protected Sub BackBtn_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BackBtn.Click

        '回上頁
        Server.Transfer("MOA04004.aspx")

    End Sub
End Class
