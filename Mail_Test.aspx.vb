Imports System.Data
Imports System.Data.SqlClient
Partial Class M_Source_01_Mail_Test
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try

            Dim FC As New C_FlowSend.C_FlowSend

            FC.F_MailGO("administrator@oa.mil.tw", "系統測試通知", "10.20.10.101", "tempu179@staff.mil.tw", "系統測試通知", "Test")

            Response.Write("成功")

        Catch ex As Exception

            Response.Write(ex.Message)

        End Try

    End Sub
End Class
