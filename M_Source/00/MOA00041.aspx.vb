Imports System.Data.SqlClient
Partial Class Source_00_MOA00041
    Inherits System.Web.UI.Page

    Dim EmpFlag As String
    Dim CancelPer As String = ""
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

                If LoginCheck.LoginCheck(user_id, "MOA00040") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00041.aspx")
                    Response.End()
                End If

                LabErr.Visible = False
                EmpFlag = ""

                '修改人員帳號
                Dim strempuid As String = ""
                strempuid = Request.QueryString("empuid")

                '找出登入者的一級單位
                Dim strParentOrg As String = ""
                Dim Org_UP As New C_Public
                strParentOrg = Org_UP.getUporg(org_uid, 1)

                If strempuid = "" Then
                    DetailsView1.DefaultMode = DetailsViewMode.Insert
                    LabTitle.Text = "人員新增"

                    '判斷登入者權限
                    If Session("Role") = "1" Then
                        SqlDataSource4.SelectCommand = "SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] ORDER BY [ORG_NAME]"
                    ElseIf Session("Role") = "2" Then
                        SqlDataSource4.SelectCommand = "SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] WHERE ORG_UID IN (" & Org_UP.getchildorg(strParentOrg) & ") ORDER BY [ORG_NAME]"
                    Else
                        SqlDataSource4.SelectCommand = ""
                    End If


                Else
                    DetailsView1.DefaultMode = DetailsViewMode.Edit
                    LabTitle.Text = "人員修改"

                    '判斷登入者權限
                    If Session("Role") = "1" Then
                        SqlDataSource4.SelectCommand = "SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] ORDER BY [ORG_NAME]"
                    ElseIf Session("Role") = "2" Then
                        SqlDataSource4.SelectCommand = "SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] WHERE ORG_UID IN (" & Org_UP.getchildorg(strParentOrg) & ") OR ORG_UID = '520' ORDER BY [ORG_NAME]"
                    Else
                        SqlDataSource4.SelectCommand = ""
                    End If

                End If



            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub DetailsView1_ItemCreated(ByVal sender As Object, ByVal e As System.EventArgs) Handles DetailsView1.ItemCreated

        DetailsView1.FindControl("member_uid").Visible = False
        DetailsView1.FindControl("PW").Visible = False

    End Sub

    Protected Sub DetailsView1_ItemInserted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DetailsViewInsertedEventArgs) Handles DetailsView1.ItemInserted

        '回查詢頁面
        Server.Transfer("MOA00040.aspx")

    End Sub

    Protected Sub DetailsView1_ItemInserting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DetailsViewInsertEventArgs) Handles DetailsView1.ItemInserting

        '人員重複取消新增
        If EmpFlag = "1" Then
            e.Cancel = True
        End If

    End Sub

    Protected Sub DetailsView1_ItemUpdated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DetailsViewUpdatedEventArgs) Handles DetailsView1.ItemUpdated

        '人員離職如有未批核表單不可移動
        If CType(DetailsView1.FindControl("DDLLeave1"), DropDownList).SelectedValue = "N" Then

            '修改資料
            Dim connstr As String
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '人員離職同時將服務單位改成待派區
            db.Open()
            Dim updCom As New SqlCommand("UPDATE EMPLOYEE SET ORG_UID='520' WHERE employee_id = '" & CType(DetailsView1.FindControl("employee_id"), TextBox).Text & "'", db)
            updCom.ExecuteNonQuery()
            db.Close()

        End If

        '回查詢頁面
        Server.Transfer("MOA00040.aspx")

    End Sub

    Protected Sub DetailsView1_ItemUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DetailsViewUpdateEventArgs) Handles DetailsView1.ItemUpdating

        '取消人員移到待派區
        If EmpFlag = "2" Then
            e.Cancel = True

            '轉移到重新分派
            Server.Transfer("MOA00042.aspx?CancelPer=" & CancelPer)

        End If

    End Sub

    Protected Sub btnImgUpd_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        Try
            '人員離職如有未批核表單不可移動
            If CType(DetailsView1.FindControl("DDLLeave1"), DropDownList).SelectedValue = "N" Then

                '修改資料
                Dim connstr As String
                Dim conn As New C_SQLFUN
                connstr = conn.G_conn_string

                '開啟連線
                Dim db As New SqlConnection(connstr)

                '人員是否有未批核表單
                db.Open()
                Dim strPer As New SqlCommand("SELECT empuid FROM flowctl WHERE empuid = '" & CType(DetailsView1.FindControl("employee_id"), TextBox).Text & "' AND (hddate IS NULL) ", db)
                Dim RdPer = strPer.ExecuteReader()
                If RdPer.read() Then
                    EmpFlag = "2"
                    CancelPer = RdPer("empuid")
                End If
                db.Close()

                If EmpFlag <> "" Then

                    LabErr.Visible = True
                    LabErr.Text = "該人員帳號有未批核表單不可辦理離職!!"
                End If

            End If


        Catch ex As Exception

        End Try

    End Sub

    Protected Sub btnImgIns_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        Try

            '新增資料
            Dim connstr As String
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '人員是否已加入
            db.Open()
            Dim strPer As New SqlCommand("SELECT employee_id FROM EMPLOYEE WHERE employee_id = '" & CType(DetailsView1.FindControl("employee_id"), TextBox).Text & "'", db)
            Dim RdPer = strPer.ExecuteReader()
            If RdPer.read() Then
                EmpFlag = "1"
            End If
            db.Close()

            '判斷人員是否重複
            If EmpFlag = "" Then
                CType(DetailsView1.FindControl("member_uid"), TextBox).Text = CType(DetailsView1.FindControl("employee_id"), TextBox).Text
                CType(DetailsView1.FindControl("PW"), TextBox).Text = CType(DetailsView1.FindControl("employee_id"), TextBox).Text
            Else
                LabErr.Visible = True
                LabErr.Text = "人員帳號已存在!!"
            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub btnImgBack_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        '清空暫存的empuid
        Server.Transfer("MOA00040.aspx?empuid=")

    End Sub

    Protected Sub btnImgBackIns_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        Server.Transfer("MOA00040.aspx")

    End Sub
End Class
