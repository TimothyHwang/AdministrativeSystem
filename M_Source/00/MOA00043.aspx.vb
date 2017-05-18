Imports System.DirectoryServices
Imports System.Data.SqlClient

Partial Class M_Source_00_MOA00043
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String
    Dim strcompany As String = ""
    Dim strdepartment As String = ""
    Dim strtitle As String = ""
    Dim strgivenName As String = ""
    Dim strmail As String = ""
    Dim strRevive As String = ""

    Protected Sub btnImgRead_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnImgRead.Click

        Try
            LabErr.Text = ""

            If txtUserID.Text <> "" Then

                Dim entry As New DirectoryEntry("LDAP://staff.mil.tw/OU=OU-Accounts,dc=staff,dc=mil,dc=tw", "staff\administrator", "1qaz@WSX3edc", AuthenticationTypes.Secure)

                Dim searcher As New DirectorySearcher(entry)
                searcher.Filter = "(SAMAccountName=" & txtUserID.Text & ")"

                Dim sResultSet As SearchResult
                sResultSet = searcher.FindOne

                'If sResultSet.Properties("userAccountControl").Count <> 0 Then
                '    MsgBox(sResultSet.Properties("userAccountControl").Item(0).ToString)
                'End If

                If sResultSet.Properties("company").Count <> 0 Then
                    strcompany = sResultSet.Properties("company").Item(0).ToString
                End If

                If sResultSet.Properties("department").Count <> 0 Then
                    strdepartment = sResultSet.Properties("department").Item(0).ToString
                End If

                If sResultSet.Properties("title").Count <> 0 Then
                    strtitle = sResultSet.Properties("title").Item(0).ToString
                End If

                If sResultSet.Properties("givenName").Count <> 0 Then
                    strgivenName = sResultSet.Properties("givenName").Item(0).ToString
                End If

                If sResultSet.Properties("mail").Count <> 0 Then
                    strmail = sResultSet.Properties("mail").Item(0).ToString
                End If

                LabUnit.Text = strcompany & strdepartment
                LabTitle.Text = strtitle
                LabName.Text = strgivenName
                LabMail.Text = strmail

                'MsgBox(sResultSet.Properties("name").Item(0).ToString)
                'MsgBox(sResultSet.Properties("mail").Item(0).ToString)
                'MsgBox(sResultSet.Properties("cn").Item(0).ToString)
                'MsgBox(sResultSet.Properties("sn").Item(0).ToString)
                'MsgBox(sResultSet.Properties("givenName").Item(0).ToString)
                'MsgBox(sResultSet.Properties("department").Item(0).ToString)
                'MsgBox(sResultSet.Properties("telephoneNumber").Item(0).ToString)
                'MsgBox(sResultSet.Properties("title").Item(0).ToString)

                '鎖定人員帳號
                txtUserID.Enabled = False

            End If

        Catch ex As Exception

            If Request.QueryString("empuid") <> "" Then

                btnImgUpd.Visible = False
                btnImgDel.Visible = True

            End If


            LabErr.Text = "查無資料"

        End Try

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
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00043.aspx")
                    Response.End()
                End If

                '新增資料
                Dim connstr As String
                Dim conn As New C_SQLFUN
                connstr = conn.G_conn_string

                '開啟連線
                Dim db As New SqlConnection(connstr)

                '修改人員帳號
                Dim strempuid As String = ""
                strempuid = Request.QueryString("empuid")

                '找出登入者的一級單位
                Dim strParentOrg As String = ""
                Dim strAdvance As String = ""
                Dim Org_UP As New C_Public

                strParentOrg = Org_UP.getUporg(org_uid, 1)

                '判斷登入者是否為高級長官管理員相關群組人員
                db.Open()
                Dim carCom As New SqlCommand("select * from ROLEGROUPITEM where employee_id = '" & user_id & "' AND (Group_Uid ='465N3W7BR1')", db)
                Dim RdvCar = carCom.ExecuteReader()
                If RdvCar.read() Then
                    strAdvance = "Y"
                End If
                db.Close()

                If strempuid = "" Then
                    LabCaption.Text = "人員新增"

                    '判斷登入者權限
                    If Session("Role") = "1" Then
                        SqlDataSource4.SelectCommand = "SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] ORDER BY [ORG_NAME]"
                    ElseIf Session("Role") = "2" Then
                        SqlDataSource4.SelectCommand = "SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] WHERE ORG_UID IN (" & Org_UP.getchildorg(strParentOrg) & ") ORDER BY [ORG_NAME]"
                    ElseIf strAdvance = "Y" Then
                        SqlDataSource4.SelectCommand = "SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] WHERE ORG_UID IN (" & Org_UP.getchildorg("796A91A33O") & ") ORDER BY ORG_NAME"
                    Else
                        SqlDataSource4.SelectCommand = ""
                    End If

                Else
                    LabCaption.Text = "人員修改"

                    '判斷登入者權限
                    If Session("Role") = "1" Then
                        SqlDataSource4.SelectCommand = "SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] ORDER BY [ORG_NAME]"
                    ElseIf Session("Role") = "2" Then
                        SqlDataSource4.SelectCommand = "SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] WHERE ORG_UID IN (" & Org_UP.getchildorg(strParentOrg) & ") OR ORG_UID = '520' ORDER BY [ORG_NAME]"
                    ElseIf strAdvance = "Y" Then
                        SqlDataSource4.SelectCommand = "SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] WHERE ORG_UID IN (" & Org_UP.getchildorg("796A91A33O") & ") OR ORG_UID = '520' ORDER BY ORG_NAME"
                    Else
                        SqlDataSource4.SelectCommand = ""
                    End If

                    txtUserID.Text = strempuid
                    txtUserID.Enabled = False
                    ImgClearAll.Visible = False

                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub btnImgUpd_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnImgUpd.Click

        Try

            '新增資料
            Dim connstr As String
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '修改人員帳號
            Dim strempuid As String = ""
            strempuid = Request.QueryString("empuid")

            Dim EmpFlag As String = ""
            Dim CancelPer As String = ""

            If strempuid = "" Then

                '人員是否已加入
                db.Open()
                Dim strPer As New SqlCommand("SELECT employee_id,ORG_UID FROM EMPLOYEE WHERE employee_id = '" & txtUserID.Text & "'", db)
                Dim RdPer = strPer.ExecuteReader()
                If RdPer.read() Then
                    strRevive = RdPer("ORG_UID")
                    EmpFlag = "1"
                End If
                db.Close()

                '判斷人員是否重複
                If EmpFlag = "" Then

                    If txtUserID.Text = "" Or LabName.Text = "" Then

                        'Response.Write(" <script language='javascript'>")
                        'Response.Write(" alert('全部欄位皆為必填，請全部輸入!');")
                        'Response.Write(" </script>")

                        LabErr.Text = "全部欄位皆為必填，請全部輸入!"

                    Else

                        '新增使用者資料
                        db.Open()
                        Dim insCom As New SqlCommand("INSERT INTO EMPLOYEE(employee_id, member_uid, ORG_UID, empemail, emp_chinese_name, leave, ArriveDate,title_id,TU_ID, AD_DEP, AD_TITLE) VALUES ('" & txtUserID.Text & "','" & txtUserID.Text & "','" & DDLORG.SelectedValue & "','" & LabMail.Text & "',N'" & LabName.Text & "','Y','" & Now().Date & "','9',6,'" & LabUnit.Text & "','" & LabTitle.Text & "')", db)
                        insCom.ExecuteNonQuery()
                        db.Close()

                        '新增使用者權限資料
                        db.Open()
                        Dim insAuthCom As New SqlCommand("INSERT INTO ROLEGROUPITEM(Group_Uid, employee_id) VALUES ('JROHEST34K','" & txtUserID.Text & "')", db)
                        insAuthCom.ExecuteNonQuery()
                        db.Close()

                        Server.Transfer("MOA00040.aspx")

                    End If

                Else
                    If strRevive = "999" Then

                        '人員從歷史資料區移回
                        db.Open()
                        Dim updCom As New SqlCommand("UPDATE EMPLOYEE SET ORG_UID='" & DDLORG.SelectedValue & "',Leave='N', empemail='" & LabMail.Text & "', emp_chinese_name=N'" & LabName.Text & "', AD_DEP='" & LabUnit.Text & "', AD_TITLE='" & LabTitle.Text & "' WHERE employee_id = '" & txtUserID.Text & "'", db)
                        updCom.ExecuteNonQuery()
                        db.Close()

                        '新增使用者權限資料
                        db.Open()
                        Dim insAuthCom As New SqlCommand("INSERT INTO ROLEGROUPITEM(Group_Uid, employee_id) VALUES ('JROHEST34K','" & txtUserID.Text & "')", db)
                        insAuthCom.ExecuteNonQuery()
                        db.Close()

                        LabErr.Text = "人員已從歷史資料區移回"
                    Else
                        LabErr.Text = "人員帳號已存在!!"
                    End If
                End If

            Else

                If txtUserID.Text = "" Or LabName.Text = "" Then

                    'Response.Write(" <script language='javascript'>")
                    'Response.Write(" alert('全部欄位皆為必填，請全部輸入!');")
                    'Response.Write(" </script>")

                    LabErr.Text = "全部欄位皆為必填，請全部輸入!"

                Else
                    '刪除人員的影印管理使用者資料，人員有異動皆需做該人員的影印機管理的重設作業
                    db.Open()
                    Dim delPrinterCom As New SqlCommand("DELETE FROM P_0803 WHERE (employee_id = '" & txtUserID.Text & "')", db)
                    delPrinterCom.ExecuteNonQuery()
                    db.Close()

                    '判斷人員是否離職同時將服務單位改成待派區
                    If DDLYN.SelectedValue = "N" Then
                        '人員是否有未批核表單
                        db.Open()
                        Dim strPer As New SqlCommand("SELECT empuid FROM flowctl WHERE empuid = '" & txtUserID.Text & "' AND (hddate IS NULL) ", db)
                        Dim RdPer As SqlDataReader = strPer.ExecuteReader()
                        If RdPer.read() Then
                            EmpFlag = "2"
                            CancelPer = RdPer("empuid")
                        End If
                        db.Close()

                        If EmpFlag <> "" Then

                            Response.Write(" <script language='javascript'>")
                            Response.Write(" alert('該人員帳號有未批核表單不可辦理離職!!');")
                            Response.Write(" </script>")

                            '轉移到重新分派
                            Server.Transfer("MOA00042.aspx?CancelPer=" & CancelPer)
                        Else
                            '人員移到待派區
                            db.Open()
                            Dim updCom As New SqlCommand("UPDATE EMPLOYEE SET ORG_UID='520',Leave='N', empemail='" & LabMail.Text & "', emp_chinese_name=N'" & LabName.Text & "', AD_DEP='" & LabUnit.Text & "', AD_TITLE='" & LabTitle.Text & "' WHERE employee_id = '" & txtUserID.Text & "'", db)
                            updCom.ExecuteNonQuery()
                            db.Close()

                            '取消人員群組權限
                            db.Open()
                            Dim delCom As New SqlCommand("DELETE FROM ROLEGROUPITEM WHERE (employee_id = '" & txtUserID.Text & "') AND (Group_Uid <> 'JROHEST34K')", db)
                            delCom.ExecuteNonQuery()
                            db.Close()

                            '取消人員關卡權限
                            db.Open()
                            Dim delSTCom As New SqlCommand("DELETE FROM SYSTEMOBJUSE WHERE (employee_id = '" & txtUserID.Text & "')", db)
                            delSTCom.ExecuteNonQuery()
                            db.Close()
                            '修改人員完成
                            Server.Transfer("MOA00040.aspx?empuid=")
                        End If
                    Else
                        '判斷人員原始單位
                        db.Open()
                        Dim strPer As New SqlCommand("SELECT employee_id,ORG_UID FROM EMPLOYEE WHERE employee_id = '" & txtUserID.Text & "'", db)
                        Dim RdPer = strPer.ExecuteReader()
                        If RdPer.read() Then
                            strRevive = RdPer("ORG_UID")
                        End If
                        db.Close()

                        db.Open()
                        Dim updCom As New SqlCommand("UPDATE EMPLOYEE SET ORG_UID='" & DDLORG.SelectedValue & "',Leave='Y', empemail='" & LabMail.Text & "', emp_chinese_name=N'" & LabName.Text & "', AD_DEP='" & LabUnit.Text & "', AD_TITLE='" & LabTitle.Text & "' WHERE employee_id = '" & txtUserID.Text & "'", db)
                        updCom.ExecuteNonQuery()
                        db.Close()

                        If strRevive = "999" Then
                            '歷史區移回人員加入一般權限
                            db.Open()
                            Dim insAuthCom As New SqlCommand("INSERT INTO ROLEGROUPITEM(Group_Uid, employee_id) VALUES ('JROHEST34K','" & txtUserID.Text & "')", db)
                            insAuthCom.ExecuteNonQuery()
                            db.Close()
                        Else
                            '移動使用者單位取消權限
                            If strRevive <> DDLORG.SelectedValue Then
                                '取消人員群組權限
                                db.Open()
                                Dim delCom As New SqlCommand("DELETE FROM ROLEGROUPITEM WHERE (employee_id = '" & txtUserID.Text & "') AND (Group_Uid <> 'JROHEST34K')", db)
                                delCom.ExecuteNonQuery()
                                db.Close()

                                '取消人員關卡權限
                                db.Open()
                                Dim delSTCom As New SqlCommand("DELETE FROM SYSTEMOBJUSE WHERE (employee_id = '" & txtUserID.Text & "')", db)
                                delSTCom.ExecuteNonQuery()
                                db.Close()
                            End If
                        End If
                        '修改人員完成
                        Server.Transfer("MOA00040.aspx?empuid=")
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub btnImgBack_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnImgBack.Click

        '回查詢頁面清空暫存的empuid
        Server.Transfer("MOA00040.aspx?empuid=")

    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        If Request.QueryString("empuid") <> "" Then

            DDLORG.SelectedItem.Selected = False
            DDLYN.SelectedItem.Selected = False

            '新增資料
            Dim connstr As String
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            Dim strOrg_uid As String = ""
            Dim strLeave As String = ""

            '搜尋人員資料
            db.Open()
            Dim strPer As New SqlCommand("SELECT * FROM EMPLOYEE WHERE employee_id = '" & txtUserID.Text & "'", db)
            Dim RdPer = strPer.ExecuteReader()
            If RdPer.read() Then
                strOrg_uid = RdPer("ORG_UID")
                strLeave = RdPer("leave")
            End If
            db.Close()

            Dim i As Integer = 0
            For i = 0 To DDLORG.Items.Count - 1
                If DDLORG.Items(i).Value = strOrg_uid Then
                    DDLORG.Items(i).Selected = True
                End If
            Next

            Dim j As Integer = 0
            For j = 0 To DDLYN.Items.Count - 1
                If DDLYN.Items(j).Value = strLeave Then
                    DDLYN.Items(j).Selected = True
                End If
            Next

        End If

    End Sub

    Protected Sub ImgClearAll_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgClearAll.Click

        '清除選取的AD資料

        txtUserID.Enabled = True
        txtUserID.Text = ""
        LabUnit.Text = ""
        LabTitle.Text = ""
        LabName.Text = ""
        LabMail.Text = ""


    End Sub

    Protected Sub btnImgDel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnImgDel.Click

        '新增資料
        Dim connstr As String
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '人員移到待派區
        db.Open()
        Dim updCom As New SqlCommand("UPDATE EMPLOYEE SET ORG_UID='520',Leave='N', empemail='" & LabMail.Text & "', AD_DEP='" & LabUnit.Text & "', AD_TITLE='" & LabTitle.Text & "' WHERE employee_id = '" & txtUserID.Text & "'", db)
        updCom.ExecuteNonQuery()
        db.Close()

        '取消人員群組權限
        db.Open()
        Dim delCom As New SqlCommand("DELETE FROM ROLEGROUPITEM WHERE (employee_id = '" & txtUserID.Text & "') AND (Group_Uid <> 'JROHEST34K')", db)
        delCom.ExecuteNonQuery()
        db.Close()

        '取消人員關卡權限
        db.Open()
        Dim delSTCom As New SqlCommand("DELETE FROM SYSTEMOBJUSE WHERE (employee_id = '" & txtUserID.Text & "')", db)
        delSTCom.ExecuteNonQuery()
        db.Close()

        '刪除人員的影印管理使用者資料，人員有異動皆需做該人員的影印機管理的重設作業
        db.Open()
        Dim delPrinterCom As New SqlCommand("DELETE FROM P_0803 WHERE (employee_id = '" & txtUserID.Text & "')", db)
        delPrinterCom.ExecuteNonQuery()
        db.Close()

        '修改人員完成
        Server.Transfer("MOA00040.aspx?empuid=")
    End Sub
End Class
