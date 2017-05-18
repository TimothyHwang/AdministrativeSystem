
Partial Class Source_00_MOA00060
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String
    Protected Function FunType(ByVal str As String) As String
        Try
            '轉換共用代號
            Dim tmpStr = Eval(str)

            If tmpStr = "1" Then
                tmpStr = "全關卡"
            ElseIf tmpStr = "2" Then
                tmpStr = "上一級單位"
            ElseIf tmpStr = "3" Then
                tmpStr = "同單位"
            ElseIf tmpStr = "4" Then
                tmpStr = "指定群組"
            End If

            FunType = tmpStr

        Catch ex As Exception
            FunType = ""
        End Try
    End Function

    Protected Sub btnImgIns_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        If CType(Me.DetailsView1.FindControl("object_name"), TextBox).Text = "" Then
            CType(Me.DetailsView1.FindControl("InsErr"), Label).Text = "請輸入關卡名稱"
        ElseIf CType(Me.DetailsView1.FindControl("display_order"), TextBox).Text = "" Then
            CType(Me.DetailsView1.FindControl("InsErr"), Label).Text = "請輸入排序號碼"
        Else
            Dim randstr As New C_Public
            '新增資料
            If Me.DetailsView1.CurrentMode = DetailsViewMode.Insert Then
                CType(Me.DetailsView1.FindControl("object_uid"), TextBox).Text = randstr.randstr(10)
                'CType(Me.DetailsView1.FindControl("object_type"), TextBox).Text = "1"
                CType(Me.DetailsView1.FindControl("btnImgIns"), ImageButton).CommandName = "insert"
            End If
        End If

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

                If LoginCheck.LoginCheck(user_id, "MOA00060") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00060.aspx")
                    Response.End()
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub
End Class
