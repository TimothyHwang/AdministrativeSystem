Imports System.Data.SqlClient
Partial Class Source_00_MOA00055
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String
    Protected Sub ImageButton1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton1.Click

        '判斷查詢組織或是人員
        Dim strsql As String = ""
        strsql += "SELECT ADMINGROUP.ORG_UID, ADMINGROUP.ORG_NAME, EMPLOYEE.employee_id, EMPLOYEE.emp_chinese_name FROM EMPLOYEE INNER JOIN ADMINGROUP ON EMPLOYEE.ORG_UID = ADMINGROUP.ORG_UID"

        If RadioButtonList1.SelectedValue = "1" Then
            strsql += " WHERE ADMINGROUP.ORG_NAME like '%" & Trim(SearchValue.Text) & "%'"
        ElseIf RadioButtonList1.SelectedValue = "2" Then
            strsql += " WHERE EMPLOYEE.emp_chinese_name like '%" & Trim(SearchValue.Text) & "%'"
        ElseIf RadioButtonList1.SelectedValue = "3" Then
            strsql += " WHERE EMPLOYEE.employee_id like '%" & Trim(SearchValue.Text) & "%'"
        End If

        SqlDataSource1.SelectCommand = strsql & " ORDER BY EMPLOYEE.emp_chinese_name"


    End Sub

    Protected Sub btnImgIns_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnImgIns.Click

        '新增資料
        Dim connstr As String
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '新增群組人員
        Dim i As Integer
        For i = 0 To empList.Items.Count - 1
            If empList.Items(i).Selected = True Then
                '人員是否已加入
                db.Open()
                Dim EmpFlag As String = ""
                Dim strPer As New SqlCommand("SELECT Role_Num FROM ROLEGROUPITEM WHERE Group_Uid = '" & DropDownList1.SelectedValue & "' AND employee_id = '" & empList.Items(i).Value & "'", db)
                Dim RdPer = strPer.ExecuteReader()
                If RdPer.read() Then
                    EmpFlag = "1"
                    lbMsg.Text = "此人員已存在於您選擇的群組中!!"
                End If
                db.Close()

                If EmpFlag = "" Then
                    '新增群組人員資料
                    db.Open()
                    Dim insCom As New SqlCommand("INSERT INTO ROLEGROUPITEM(Group_Uid,employee_id) VALUES ('" & DropDownList1.SelectedValue & "','" & empList.Items(i).Value & "')", db)
                    insCom.ExecuteNonQuery()
                    db.Close()
                    lbMsg.Text = "此人員已成功加入群組中!!"
                End If
            End If
        Next
    End Sub

    Protected Sub BackBtn_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BackBtn.Click
        '回上頁
        Server.Transfer("MOA00054.aspx")
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

                If LoginCheck.LoginCheck(user_id, "MOA00054") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00055.aspx")
                    Response.End()
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub
End Class
