Imports System.Data
Imports System.Data.SqlClient
Partial Class OA_System_RoleLeft
    Inherits System.Web.UI.Page

    Dim MyConnection As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
    Dim HTMLstring As String = "<br>"
    Dim user_id, org_uid As String

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

                If LoginCheck.LoginCheck(user_id, "MOA00051") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00052.aspx")
                    Response.End()
                End If

                '找出群組中的連結-------------------------------------------------------------------------------
                Dim sql1 As String = "select group_Uid,group_Name from rolegroup order by Group_Order"
                MyConnection.Open()
                Dim dt1 As DataTable = New DataTable("rolegroup")
                Dim da1 As SqlDataAdapter = New SqlDataAdapter(sql1, MyConnection)
                da1.Fill(dt1)
                MyConnection.Close()
                Dim img As ImageButton
                Dim HyperLink1 As HyperLink
                For x As Integer = 0 To dt1.Rows.Count - 1
                    HyperLink1 = New HyperLink
                    img = New ImageButton
                    img.ImageUrl = "../../image/memo.gif"
                    img.PostBackUrl = "MOA00052.aspx"
                    With HyperLink1
                        .Text = dt1.Rows(x).Item(1)
                        .NavigateUrl = "MOA00053.aspx?Group_Uid=" & dt1.Rows(x).Item(0)
                        .Target = "Role_right"
                        .Font.Size = FontUnit.Parse("10pt")
                        .ForeColor = System.Drawing.Color.Black
                    End With
                    PlaceHolder1.Controls.Add(img)
                    PlaceHolder1.Controls.Add(HyperLink1)
                    PlaceHolder1.Controls.Add(New LiteralControl(HTMLstring))
                Next

            End If

        Catch SqlEx As SqlException         '連線錯誤訊息
            'MsgBox(SqlEx.Message)
        Catch ex As Exception
            'MsgBox(ex.Message)
        Finally

        End Try

    End Sub
End Class
