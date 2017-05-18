Imports System.Data.SqlClient
Partial Class Source_00_MOA00013
    Inherits System.Web.UI.Page

    Dim eformid, user_id, org_uid, streformsn, connstr As String
    Dim SelFlag As String = ""      '1.上一級主管 2.同單位行政官 3.一級單位全部行政官
    Dim CancelFlag, streformsnOld As String
    Dim Org_Name, User_Name, Title_Name As String
    Dim strGroup_ID As String
    Dim AppUID As String
    Public read_only As String = ""

    Dim AgentEmpuid As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            '新增資料
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            ''申請人身分證ID
            AppUID = Request("AppUID").ToString

            If Session("user_id") = "" Then
                user_id = Page.User.Identity.Name.Substring(Page.User.Identity.Name.IndexOf("\") + 1)
            Else
                user_id = Session("user_id")
            End If

            read_only = Request("read_only")

            org_uid = Session("ORG_UID")

            AgentEmpuid = Session("AgentEmpuid")

            eformid = Request.QueryString("eformid")
            streformsn = Request.QueryString("eformsn")
            SelFlag = Request.QueryString("SelFlag")

            '資訊設備維修
            strGroup_ID = Request.QueryString("group_id")

            '銷假
            CancelFlag = Request.QueryString("CancelFlag")
            streformsnOld = Request.QueryString("eformsnOld")

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '登入者資料
            db.Open()
            Dim strPer As New SqlCommand("SELECT ORG_NAME,emp_chinese_name,AD_TITLE FROM V_EmpInfo WHERE employee_id = '" & user_id & "'", db)
            Dim RdPer = strPer.ExecuteReader()
            If RdPer.Read() Then
                Org_Name = RdPer("ORG_NAME")
                User_Name = RdPer("emp_chinese_name")
                Title_Name = RdPer("AD_TITLE")
            End If
            db.Close()

            '判斷登入者權限
            Dim LoginCheck As New C_Public

            If eformid = "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00013.aspx")
                Response.End()
            End If

            If SelFlag = "1" Then
                LabTitle.Text = "選擇上一級主管"
                LabName.Text = "上一級主管："
                SqlDataSource1.SelectCommand = " select employee_id,emp_chinese_name from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = (select ORG_UID from EMPLOYEE where employee_id = '" & user_id & "')) "
            ElseIf SelFlag = "2" Then
                '同單位行政官
                LabTitle.Text = "選擇單位行政官"
                LabName.Text = "行政官："
                SqlDataSource1.SelectCommand = "SELECT DISTINCT EMPLOYEE.employee_id,EMPLOYEE.emp_chinese_name FROM SYSTEMOBJUSE INNER JOIN EMPLOYEE ON SYSTEMOBJUSE.employee_id = EMPLOYEE.employee_id INNER JOIN ADMINGROUP ON EMPLOYEE.ORG_UID = ADMINGROUP.ORG_UID WHERE (SYSTEMOBJUSE.object_uid = 'p12') AND (EMPLOYEE.ORG_UID = '" & org_uid & "') ORDER BY EMPLOYEE.emp_chinese_name"
            ElseIf SelFlag = "3" Then
                '一級單位行政官
                LabTitle.Text = "選擇單位行政官"
                LabName.Text = "行政官："
                Dim strOrgTop As String = Request.QueryString("strOrgTop")
                Dim CP As New C_Public
                SqlDataSource1.SelectCommand = "SELECT DISTINCT EMPLOYEE.employee_id,EMPLOYEE.emp_chinese_name FROM SYSTEMOBJUSE INNER JOIN EMPLOYEE ON SYSTEMOBJUSE.employee_id = EMPLOYEE.employee_id INNER JOIN ADMINGROUP ON EMPLOYEE.ORG_UID = ADMINGROUP.ORG_UID WHERE (SYSTEMOBJUSE.object_uid = 'p12') AND (EMPLOYEE.ORG_UID IN (" & CP.getchildorg(strOrgTop) & ")) ORDER BY EMPLOYEE.emp_chinese_name"
            ElseIf SelFlag = "4" Then
                LabTitle.Text = "選擇資訊設備維修管制人員"
                LabName.Text = "管制人員："
                SqlDataSource1.SelectCommand = "SELECT * FROM SYSTEMOBJ A JOIN SYSTEMOBJUSE B ON A.OBJECT_UID=B.OBJECT_UID LEFT JOIN EMPLOYEE C ON B.EMPLOYEE_ID=C.EMPLOYEE_ID WHERE A.OBJECT_UID='" & strGroup_ID & "'"
            End If

        Catch ex As Exception

        End Try


    End Sub

    Protected Sub btnSend_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSend.Click

        If CancelFlag <> "" Then
            '新增撤銷單資料
            InsData()
        End If

        '表單審核
        'Dim SendVal = eformid & "," & user_id & "," & streformsn & "," & "1" & "," & DDLUser.SelectedValue
        Dim SendVal = eformid & "," & AppUID & "," & streformsn & "," & "1" & "," & DDLUser.SelectedValue

        '表單送件
        Dim Val_P As String
        Val_P = ""
        Dim db As New SqlConnection(connstr)
        Dim FCC As New CFlowSend
        Dim tool As New C_Public

        Val_P = FCC.F_Send(SendVal, connstr)

        If eformid = "BL7U2QP3IG" AndAlso SelFlag = "4" Then ''更新回資訊設備維修主資料表
            Dim flowctl As New flowctl(eformid, streformsn, "?") ''因已F_Send，所以此時抓取的是已進入下一關的簽核流程

            If flowctl.group_name = "資訊維修單位" Then ''確認流程為管制單位分派至維修人員
                Dim DC As New SQLDBControl
                Dim strSql As String = "UPDATE " & tool.GetPrimaryTableName(eformid) & " SET FIXIDNO='" & DDLUser.SelectedValue & "' WHERE EFORMSN='" & streformsn & "'"
                DC.ExecuteSQL(strSql)
                DC.Dispose()
            End If
        End If

        Dim PageUp As String = ""

        If read_only = "" Then
            PageUp = "New"
        End If

        Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & PageUp)

    End Sub

    Public Sub InsData()

        '新增連線
        Dim db As New SqlConnection(connstr)

        '開啟連線
        db.Open()

        '填表人和申請人同為登入者資料
        Dim InsP02 As String = ""
        '表單序號,填表人單位,填表人級職,填表人姓名,填表人身份證字號,申請人單位,申請人姓名,申請人級職,申請人身份證字號
        '撤銷表單編號
        InsP02 = "INSERT INTO P_0101(EFORMSN,PWUNIT,PWTITLE,PWNAME,PWIDNO,PAUNIT,PANAME, PATITLE, PAIDNO "
        InsP02 += ",nEFORMSN) "
        InsP02 += " VALUES ('" & streformsn & "','" & Org_Name & "','" & Title_Name & "',N'" & User_Name & "','" & user_id & "','" & Org_Name & "',N'" & User_Name & "','" & Title_Name & "','" & user_id & "'"
        InsP02 += ",'" & streformsnOld & "')"

        Dim insCom As New SqlCommand(InsP02, db)
        insCom.ExecuteNonQuery()
        db.Close()

    End Sub
End Class
