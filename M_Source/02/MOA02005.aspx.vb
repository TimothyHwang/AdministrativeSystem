Imports System.Data.SqlClient
Partial Class Source_02_MOA02005
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String

    Protected Sub ImgOK_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgOK.Click

        Try
            Dim tool As New C_Public
            '新增資料
            Dim connstr As String
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            If RoomName.Text = "" Or tel.Text = "" Then

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('請輸入會議室名稱或是電話');")
                Response.Write(" </script>")

            Else

                '新增會議室資料
                db.Open()
                Dim insCom As New SqlCommand("INSERT INTO P_0201(Org_Uid,MeetName,Owner,Tel,ContainNum,Share) VALUES ('" & tool.GetTopORGIDFromUserIDByLevel(user_id, 3, True) & "','" & RoomName.Text & "','" & Manager.SelectedValue & "','" & tel.Text & "','" & ContainNum.Text & "','" & Share.SelectedValue & "')", db)
                insCom.ExecuteNonQuery()
                db.Close()

                Server.Transfer("MOA02004.aspx")

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub BackBtn_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BackBtn.Click

        '回上頁
        Server.Transfer("MOA02004.aspx")

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

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

            If LoginCheck.LoginCheck(user_id, "MOA02004") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA02005.aspx")
                Response.End()
            End If

        End If

    End Sub
End Class
