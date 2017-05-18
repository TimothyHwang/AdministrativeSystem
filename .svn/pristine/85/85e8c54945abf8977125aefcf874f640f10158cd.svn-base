Imports System.Data.sqlclient
Partial Class Source_00_MOA00040
    Inherits System.Web.UI.Page

    Dim conn As New C_SQLFUN
    Dim user_id, org_uid As String
    Dim strAdvance As String = ""

    Public Function SQLALL(ByVal OrgSel, ByVal PerSel, ByVal LeaveSel, ByVal PerID)

        '找出上一級單位以下全部單位
        Dim Org_Down As New C_Public

        '找出登入者的一級單位
        Dim strParentOrg As String = ""

        '找出例外單位
        Dim strExOrg As String = ""

        strParentOrg = Org_Down.getUporg(org_uid, 1)

        '整合SQL搜尋字串
        Dim strsql As String = "SELECT * FROM V_EmpInfo WHERE 1=1 "

        If OrgSel <> "" Then
            strsql += " and ORG_UID = '" & OrgSel & "'"
            strExOrg = " OR ORG_UID='520' OR ORG_UID='999' "
        End If

        If PerID <> "" Then
            strsql += " and employee_id like '%" & PerID & "%'"
        End If

        If PerSel <> "" Then
            strsql += " and emp_chinese_name like N'%" & PerSel & "%'"
        End If

        If LeaveSel <> "" Then
            strsql += " and leave = '" & LeaveSel & "'"
        End If

        '判斷登入者權限
        If Session("Role") = "1" Then
            strsql += ""
        ElseIf Session("Role") = "2" Then
            strsql += " and (ORG_UID IN (" & Org_Down.getchildorg(strParentOrg) & ") " & strExOrg & ")"
        ElseIf strAdvance = "Y" Then
            strsql += " and (ORG_UID IN (" & Org_Down.getchildorg("796A91A33O") & ") " & strExOrg & ")"
        Else
            strsql += " AND 1=2 "
        End If

        SQLALL = strsql

    End Function

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged

        '分頁
        Dim strOrd As String

        strOrd = " ORDER BY ORG_NAME,emp_chinese_name"

        SqlDataSource1.SelectCommand = SQLALL(OrgSel.SelectedValue, emp_chinese_name.Text, DropLeave.SelectedValue, empuid.Text) & strOrd

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        Dim strOrd As String

        '排序條件
        strOrd = " ORDER BY " & GridView1.SortExpression.ToString()

        SqlDataSource1.SelectCommand = SQLALL(OrgSel.SelectedValue, emp_chinese_name.Text, DropLeave.SelectedValue, empuid.Text) & strOrd


    End Sub

    Protected Sub ImgBtn1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgBtn1.Click

        Server.Transfer("MOA00043.aspx")

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

                If LoginCheck.LoginCheck(user_id, "MOA00040") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00040.aspx")
                    Response.End()
                End If

                '找出上一級單位以下全部單位
                Dim Org_Down As New C_Public

                '找出登入者的一級單位
                Dim strParentOrg As String = ""
                strParentOrg = Org_Down.getUporg(org_uid, 1)

                Dim connstr As String = conn.G_conn_string

                '開啟連線
                Dim db As New SqlConnection(connstr)

                '判斷登入者是否為高級長官管理員相關群組人員
                db.Open()
                Dim carCom As New SqlCommand("select * from ROLEGROUPITEM where employee_id = '" & user_id & "' AND (Group_Uid ='465N3W7BR1')", db)
                Dim RdvCar = carCom.ExecuteReader()
                If RdvCar.read() Then
                    strAdvance = "Y"
                End If
                db.Close()

                '判斷登入者權限
                If Session("Role") = "1" Then
                    SqlDataSource2.SelectCommand = "select '' as ORG_UID,'' as ORG_NAME union SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
                ElseIf Session("Role") = "2" Then
                    SqlDataSource2.SelectCommand = "select '' as ORG_UID,'' as ORG_NAME union SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE (ORG_UID IN (" & Org_Down.getchildorg(strParentOrg) & ") OR ORG_UID='520') ORDER BY ORG_NAME"
                ElseIf strAdvance = "Y" Then
                    SqlDataSource2.SelectCommand = "select '' as ORG_UID,'' as ORG_NAME union SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE (ORG_UID IN (" & Org_Down.getchildorg("796A91A33O") & ") OR ORG_UID='520') ORDER BY ORG_NAME"
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click

        Dim strOrd As String

        strOrd = " ORDER BY ORG_NAME,emp_chinese_name"

        SqlDataSource1.SelectCommand = SQLALL(OrgSel.SelectedValue, emp_chinese_name.Text, DropLeave.SelectedValue, empuid.Text) & strOrd

    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        If IsPostBack = False Then

            '登入馬上查詢
            ImgSearch_Click(Nothing, Nothing)

        End If

    End Sub
End Class
