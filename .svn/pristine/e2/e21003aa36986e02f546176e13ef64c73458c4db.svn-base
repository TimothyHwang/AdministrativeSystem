Imports System.Data.SqlClient
Imports System.Data
Imports System.IO
Partial Class Source_01_MOA01005
    Inherits System.Web.UI.Page

    Public do_sql As New C_SQLFUN
    Dim dr As System.Data.DataRow
    Dim n_table As New System.Data.DataTable
    Dim conn As New C_SQLFUN
    Public OrgChange As String      '判斷組織是否變更
    Dim user_id, org_uid As String
    Dim STMT As String = ""
    Public print_file As String = ""
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

            Dim conn As New C_SQLFUN
            Dim connstr As String = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            Dim ParentFlag As String = ""

            '判斷是否有下一級單位
            db.Open()
            Dim strPer As New SqlCommand("SELECT ORG_UID FROM ADMINGROUP WHERE PARENT_ORG_UID = '" & org_uid & "'", db)
            Dim RdPer = strPer.ExecuteReader()
            If RdPer.read() Then
                ParentFlag = "Y"
            End If
            db.Close()

            Dim UnitFlag As String = ""

            '判斷是否有處單位人事員
            db.Open()
            Dim strUnit As New SqlCommand("SELECT Role_Num FROM ROLEGROUPITEM WHERE Group_Uid = 'JKGJZ4439V' AND employee_id = '" & user_id & "'", db)
            Dim RdUnit = strUnit.ExecuteReader()
            If RdUnit.read() Then
                UnitFlag = "Y"
            End If
            db.Close()

            '找出同級單位以下全部單位
            Dim Org_Down As New C_Public

            '找出登入者的二級單位
            Dim strParentOrg As String = ""
            strParentOrg = Org_Down.getUporg(org_uid, 2)

            '判斷登入者權限
            If Session("Role") = "1" Then
                SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
            Else

                If ParentFlag = "Y" Then
                    SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ") ORDER BY ORG_NAME"
                ElseIf UnitFlag = "Y" Then
                    SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(strParentOrg) & ") ORDER BY ORG_NAME"
                Else
                    SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID = '" & org_uid & "' ORDER BY ORG_NAME"
                End If

            End If

        End If


    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click

        '查詢
        Dim strOrd As String

        strOrd = " ORDER BY emp_chinese_name"

        SqlDataSource2.SelectCommand = SQLALL(OrgSel.SelectedValue, UserSel.SelectedValue, YearSel.SelectedValue, MonthSel.SelectedValue) & strOrd

        OrgChange = ""

    End Sub

    Protected Sub UserSel_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles UserSel.PreRender

        If OrgChange = "1" Then
            Dim tLItm As New ListItem("請選擇", "")

            '人員加請選擇
            UserSel.Items.Insert(0, tLItm)
            If UserSel.Items.Count > 1 Then
                UserSel.Items(1).Selected = False
            End If

        End If

    End Sub

    Protected Sub OrgSel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles OrgSel.SelectedIndexChanged

        '判斷組織是否變更
        OrgChange = "1"

        '清空User重新讀取
        UserSel.Items.Clear()

        If OrgSel.SelectedValue = "" Then
            SqlDataSource3.SelectCommand = "SELECT employee_id, emp_chinese_name, ORG_UID FROM EMPLOYEE WHERE 1=2"
        Else
            SqlDataSource3.SelectCommand = "SELECT employee_id, emp_chinese_name, ORG_UID FROM EMPLOYEE WHERE ORG_UID ='" & OrgSel.SelectedValue & "' ORDER BY emp_chinese_name"
        End If

    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        If Not IsPostBack Then

            Dim tLItm As New ListItem("請選擇", "")

            '系統管理員組織加請選擇

            OrgSel.Items.Insert(0, tLItm)
            If OrgSel.Items.Count > 1 Then
                OrgSel.Items(1).Selected = False
            End If

            '人員加請選擇
            UserSel.Items.Insert(0, tLItm)
            If UserSel.Items.Count > 1 Then
                UserSel.Items(1).Selected = False
            End If

            '月份加請選擇
            MonthSel.Items.Insert(0, tLItm)
            If MonthSel.Items.Count > 1 Then
                MonthSel.Items(1).Selected = False
            End If

            '登入馬上查詢
            ImgSearch_Click(Nothing, Nothing)

        End If

    End Sub

    Public Function SQLALL(ByVal OrgSel, ByVal UserSel, ByVal YearSel, ByVal MonthSel)

        '整合SQL搜尋字串
        Dim strsql As String = ""
        Dim strsel As String = ""
        Dim ParentFlag As String = ""

        If MonthSel = "" Then

            strsql = "SELECT ORG_UID, emp_chinese_name,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '事假') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "')) AS s1,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '慰勞假') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "')) AS s2,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE LIKE '%病假%') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "')) AS s3,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '值日補休') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "')) AS s4,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '產假') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "')) AS s5,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '陪產假') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "')) AS s6,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '婚假') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "')) AS s7,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '喪假') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "')) AS s8,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '公假(差)') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "')) AS s9,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '產前假') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "')) AS s10,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '流產假') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "')) AS s11"
            strsql += " FROM EMPLOYEE WHERE (leave = 'Y')"

        Else

            strsql = "SELECT ORG_UID, emp_chinese_name,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '事假') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "') AND (DATEPART(mm, nSTARTTIME) = '" & MonthSel & "') AND (DATEPART(mm, nENDTIME) = '" & MonthSel & "')) AS s1,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '慰勞假') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "') AND (DATEPART(mm, nSTARTTIME) = '" & MonthSel & "') AND (DATEPART(mm, nENDTIME) = '" & MonthSel & "')) AS s2,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE LIKE '%病假%') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "') AND (DATEPART(mm, nSTARTTIME) = '" & MonthSel & "') AND (DATEPART(mm, nENDTIME) = '" & MonthSel & "')) AS s3,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '值日補休') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "') AND (DATEPART(mm, nSTARTTIME) = '" & MonthSel & "') AND (DATEPART(mm, nENDTIME) = '" & MonthSel & "')) AS s4,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '產假') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "') AND (DATEPART(mm, nSTARTTIME) = '" & MonthSel & "') AND (DATEPART(mm, nENDTIME) = '" & MonthSel & "')) AS s5,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '陪產假') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "') AND (DATEPART(mm, nSTARTTIME) = '" & MonthSel & "') AND (DATEPART(mm, nENDTIME) = '" & MonthSel & "')) AS s6,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '婚假') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "') AND (DATEPART(mm, nSTARTTIME) = '" & MonthSel & "') AND (DATEPART(mm, nENDTIME) = '" & MonthSel & "')) AS s7,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '喪假') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "') AND (DATEPART(mm, nSTARTTIME) = '" & MonthSel & "') AND (DATEPART(mm, nENDTIME) = '" & MonthSel & "')) AS s8,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '公假(差)') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "') AND (DATEPART(mm, nSTARTTIME) = '" & MonthSel & "') AND (DATEPART(mm, nENDTIME) = '" & MonthSel & "')) AS s9,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '產前假') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "') AND (DATEPART(mm, nSTARTTIME) = '" & MonthSel & "') AND (DATEPART(mm, nENDTIME) = '" & MonthSel & "')) AS s10,"
            strsql += " (SELECT SUM(ALLHOUR) FROM V_P01_ABSENCE WHERE (PAIDNO = EMPLOYEE.employee_id) AND EFORMSN NOT IN (select nEFORMSN from p_0101 where PENDFLAG='E') AND (nTYPE = '流產假') AND (DATEPART(yy, nSTARTTIME) >= '" & YearSel & "') AND (DATEPART(yy, nENDTIME) <= '" & YearSel & "') AND (DATEPART(mm, nSTARTTIME) = '" & MonthSel & "') AND (DATEPART(mm, nENDTIME) = '" & MonthSel & "')) AS s11"
            strsql += " FROM EMPLOYEE WHERE (leave = 'Y')"

        End If


        Dim connstr As String = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '判斷是否有下一級單位
        db.Open()
        Dim strPer As New SqlCommand("SELECT ORG_UID FROM ADMINGROUP WHERE PARENT_ORG_UID = '" & org_uid & "'", db)
        Dim RdPer = strPer.ExecuteReader()
        If RdPer.read() Then
            ParentFlag = "Y"
        End If
        db.Close()

        Dim UnitFlag As String = ""

        '判斷是否有處單位人事員
        db.Open()
        Dim strUnit As New SqlCommand("SELECT Role_Num FROM ROLEGROUPITEM WHERE Group_Uid = 'JKGJZ4439V' AND employee_id = '" & user_id & "'", db)
        Dim RdUnit = strUnit.ExecuteReader()
        If RdUnit.read() Then
            UnitFlag = "Y"
        End If
        db.Close()

        '系統管理員
        If Session("Role") = "1" Then

            '組織搜尋
            If OrgSel = "" Then
                strsel = " AND 1=1"
            Else
                strsel = " AND ORG_UID='" & OrgSel & "'"
            End If

            '人員
            If UserSel <> "" Then
                strsel += " AND employee_id = '" & UserSel & "'"
            End If

        Else
            If ParentFlag = "Y" Or UnitFlag = "Y" Then

                '主官管
                '組織搜尋
                If OrgSel = "" Then
                    strsel = " AND ORG_UID='" & org_uid & "'"
                Else
                    strsel = " AND ORG_UID='" & OrgSel & "'"
                End If

                '人員
                If UserSel <> "" Then
                    strsel += " AND employee_id = '" & UserSel & "'"
                End If

            Else

                '一般人員
                '人員
                strsel = " AND employee_id = '" & user_id & "'"

            End If


        End If

        SQLALL = strsql & strsel

    End Function

    Protected Sub YearSel_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles YearSel.PreRender

        '年度
        Dim d As Date = Now()
        Dim yy As Integer = Year(d)

        Dim tLItm1 As New ListItem(yy, yy)
        Dim tLItm2 As New ListItem(yy - 1, yy - 1)
        Dim tLItm3 As New ListItem(yy - 2, yy - 2)
        Dim tLItm4 As New ListItem(yy - 3, yy - 3)
        Dim tLItm5 As New ListItem(yy - 4, yy - 4)

        If (YearSel.Items.Count = 0) Then
            YearSel.Items.Insert(0, tLItm1)
            YearSel.Items.Insert(1, tLItm2)
            YearSel.Items.Insert(2, tLItm3)
            YearSel.Items.Insert(3, tLItm4)
            YearSel.Items.Insert(4, tLItm5)
        End If

    End Sub

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged

        '分頁
        Dim strOrd As String

        strOrd = " ORDER BY emp_chinese_name"

        SqlDataSource2.SelectCommand = SQLALL(OrgSel.SelectedValue, UserSel.SelectedValue, YearSel.SelectedValue, MonthSel.SelectedValue) & strOrd

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        '排序
        Dim strOrd As String

        strOrd = " ORDER BY " & GridView1.SortExpression.ToString()

        SqlDataSource2.SelectCommand = SQLALL(OrgSel.SelectedValue, UserSel.SelectedValue, YearSel.SelectedValue, MonthSel.SelectedValue) & strOrd

    End Sub

    Protected Function FunHours(ByVal str As String) As String
        Try
            '轉換表單狀態代號
            Dim tmpStr = Eval(str)

            If tmpStr Is DBNull.Value = True Then
                FunHours = "0"
            Else
                FunHours = (CInt(tmpStr) / 8)
            End If

            'If tmpStr Is DBNull.Value = True Then
            '    FunHours = "0天0時"
            'Else
            '    FunHours = (CInt(tmpStr) \ 8) & "天" & (CInt(tmpStr) Mod 8) & "時"
            'End If

        Catch ex As Exception
            FunHours = ""
        End Try
    End Function
    Protected Function FunHours_Rep(ByVal str As String) As String
        Try
            '轉換表單狀態代號
            Dim tmpStr As String = str

            If tmpStr Is DBNull.Value = True Then
                FunHours_Rep = "0"
            Else
                FunHours_Rep = (CInt(tmpStr) / 8)
            End If

            'If tmpStr Is DBNull.Value = True Then
            '    FunHours = "0天0時"
            'Else
            '    FunHours = (CInt(tmpStr) \ 8) & "天" & (CInt(tmpStr) Mod 8) & "時"
            'End If

        Catch ex As Exception
            FunHours_Rep = ""
        End Try
    End Function

    Protected Sub ImagePrint_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImagePrint.Click
        Dim F_file As String
        Dim F_file2 As String
        Dim F_file_name As String
        Dim strOrd As String = ""
        Dim print As New C_Xprint

        STMT = SQLALL(OrgSel.SelectedValue, UserSel.SelectedValue, YearSel.SelectedValue, MonthSel.SelectedValue) & " ORDER BY emp_chinese_name"

        If do_sql.db_sql(STMT, do_sql.G_conn_string) = False Then
            do_sql.G_errmsg = "查詢資料失敗"
            Exit Sub
        End If
        If do_sql.G_table.Rows.Count > 0 Then
            n_table = do_sql.G_table
        Else
            do_sql.G_errmsg = "查無可印資料"
            Exit Sub
        End If



        Dim prn_stmt As String = ""

        F_file_name = "rpt010050"
        F_file = MapPath("../../rpt/" + F_file_name + ".txt")
        If File.Exists(F_file) = False Then
            do_sql.G_errmsg = "檔案不存在"
            Exit Sub
        End If

        print_file = "../../drs/" + F_file_name + "-" + Now.ToString("mmssff") & ".drs"
        F_file2 = MapPath(print_file)
        Dim filename As String = print_file.Split("/")(print_file.Split("/").Length - 1)
        print.C_Xprint("rpt010050.txt", filename)
        print.NewPage()

        Call do_sql.inc_file(F_file, F_file2, F_file_name)
        Dim nPage As Integer = 0
        Dim n_line As Integer = 35
        Dim tmp_add As String = ""
        'strsql = "SELECT ORG_UID, emp_chinese_name,"
        'strsql += "  (nTYPE = '事假') AND  <= '" & YearSel & "')) AS s1,"
        'strsql += "  (nTYPE = '慰勞假') AND <= '" & YearSel & "')) AS s2,"
        'strsql += "  (nTYPE LIKE '%病假%') AND  <= '" & YearSel & "')) AS s3,"
        'strsql += "  (nTYPE = '值日補休') AND  <= '" & YearSel & "')) AS s4,"
        'strsql += "  (nTYPE = '產假') AND <= '" & YearSel & "')) AS s5,"
        'strsql += "  (nTYPE = '陪產假') AND <= '" & YearSel & "')) AS s6,"
        'strsql += "   (nTYPE = '婚假') AND <= '" & YearSel & "')) AS s7,"
        'strsql += "   (nTYPE = '喪假') AND  <= '" & YearSel & "')) AS s8,"
        'strsql += "   (nTYPE = '公假(差)') AND  <= '" & YearSel & "')) AS s9,"
        'strsql += "   (nTYPE = '產前假') AND  <= '" & YearSel & "')) AS s10,"
        'strsql += "  (nTYPE = '流產假') AND  <= '" & YearSel & "')) AS s11"
        'strsql += " FROM EMPLOYEE WHERE (leave = 'Y')"
        For Each dr In n_table.Rows
            If n_line >= 34 Then
                nPage += 1
                'Call do_sql.print_sdata(F_file2, "/init " + F_file_name + ".txt 1", "/newpage null")
                'prn_stmt = CDate(Sdate.Text).ToString("yyyy/MM/dd") & "~" & CDate(Edate.Text).ToString("yyyy/MM/dd")
                'Call do_sql.print_block(F_file2, "年度：", 0, 0, YearSel.SelectedValue & MonthSel.SelectedValue)
                Call print.Add("年度：", YearSel.SelectedValue & MonthSel.SelectedValue, 0, 0)
                n_line = 0
            End If

            'Call do_sql.print_block(F_file2, "姓名", 0, n_line * 7, dr("emp_chinese_name").ToString)
            'Call do_sql.print_block(F_file2, "慰勞假", 0, n_line * 7, FunHours_Rep(dr("s2").ToString()))
            'Call do_sql.print_block(F_file2, "值日補休", 0, n_line * 7, FunHours_Rep(dr("s4").ToString()))
            'Call do_sql.print_block(F_file2, "事假", 0, n_line * 7, FunHours_Rep(dr("s1").ToString()))
            'Call do_sql.print_block(F_file2, "病假", 0, n_line * 7, FunHours_Rep(dr("s3").ToString()))
            'Call do_sql.print_block(F_file2, "公假(差)", 0, n_line * 7, FunHours_Rep(dr("s9").ToString()))
            'Call do_sql.print_block(F_file2, "婚假", 0, n_line * 7, FunHours_Rep(dr("s7").ToString()))
            'Call do_sql.print_block(F_file2, "喪假", 0, n_line * 7, FunHours_Rep(dr("s8").ToString()))
            'Call do_sql.print_block(F_file2, "陪產假", 0, n_line * 7, FunHours_Rep(dr("s6").ToString()))
            'Call do_sql.print_block(F_file2, "產假", 0, n_line * 7, FunHours_Rep(dr("s5").ToString()))
            'Call do_sql.print_block(F_file2, "產前假", 0, n_line * 7, FunHours_Rep(dr("s10").ToString()))
            'Call do_sql.print_block(F_file2, "流產假", 0, n_line * 7, FunHours_Rep(dr("s11").ToString()))
            Call print.Add("姓名", dr("emp_chinese_name").ToString, 0, n_line * 7)
            Call print.Add("慰勞假", FunHours_Rep(dr("s2").ToString()), 0, n_line * 7)
            Call print.Add("值日補休", FunHours_Rep(dr("s4").ToString()), 0, n_line * 7)
            Call print.Add("事假", FunHours_Rep(dr("s1").ToString()), 0, n_line * 7)
            Call print.Add("病假", FunHours_Rep(dr("s3").ToString()), 0, n_line * 7)
            Call print.Add("公假(差)", FunHours_Rep(dr("s9").ToString()), 0, n_line * 7)
            Call print.Add("婚假", FunHours_Rep(dr("s7").ToString()), 0, n_line * 7)
            Call print.Add("喪假", FunHours_Rep(dr("s8").ToString()), 0, n_line * 7)
            Call print.Add("陪產假", FunHours_Rep(dr("s6").ToString()), 0, n_line * 7)
            Call print.Add("產假", FunHours_Rep(dr("s5").ToString()), 0, n_line * 7)
            Call print.Add("產前假", FunHours_Rep(dr("s10").ToString()), 0, n_line * 7)
            Call print.Add("流產假", FunHours_Rep(dr("s11").ToString()), 0, n_line * 7)

            n_line = n_line + 1
        Next

        print.EndFile()

    End Sub

    Protected Sub Img_Export_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles Img_Export.Click
        Dim strsqlAll As String = SQLALL(OrgSel.SelectedValue, UserSel.SelectedValue, YearSel.SelectedValue, MonthSel.SelectedValue) & " ORDER BY emp_chinese_name"

        '設定檔案路徑
        Dim path As String = "", filename As String = ""

        path = Server.MapPath("~/Drs/")
        filename = Date.Today.ToString("yyyyMMdd") & "_010050.csv"
        path = path & filename

        Dim colname As String = ""
        Dim data As String = ""
        Dim prn_stmt As String = ""

        Dim sw As New System.IO.StreamWriter(path, False, Encoding.GetEncoding("big5"))

        prn_stmt = YearSel.SelectedValue & MonthSel.SelectedValue
        prn_stmt += " 年度休假統計 "

        sw.WriteLine(prn_stmt)

        colname = "姓名,慰勞假,值日補休,事假,病假,公假(差),婚假,喪假,陪產假,產假,產前假,流產假,"

        colname = Left(colname, Len(colname) - 1)

        sw.WriteLine(colname)


        Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

        db.Open()
        Dim dt2 As DataTable = New DataTable("")
        Dim da2 As SqlDataAdapter = New SqlDataAdapter(strsqlAll, db)
        da2.Fill(dt2)
        db.Close()
        Dim tmp_add As String = ""

        For y As Integer = 0 To dt2.Rows.Count - 1

            data += dt2.Rows(y).Item("emp_chinese_name").ToString & "," & FunHours_Rep(dt2.Rows(y).Item("s2").ToString) & ","
            data += FunHours_Rep(dt2.Rows(y).Item("s4").ToString) & "," & FunHours_Rep(dt2.Rows(y).Item("s1").ToString) & ","
            data += FunHours_Rep(dt2.Rows(y).Item("s3").ToString) & "," & FunHours_Rep(dt2.Rows(y).Item("s9").ToString) & ","
            data += FunHours_Rep(dt2.Rows(y).Item("s7").ToString) & "," & FunHours_Rep(dt2.Rows(y).Item("s8").ToString) & ","
            data += FunHours_Rep(dt2.Rows(y).Item("s6").ToString) & "," & FunHours_Rep(dt2.Rows(y).Item("s5").ToString) & ","
            data += FunHours_Rep(dt2.Rows(y).Item("s10").ToString) & "," & FunHours_Rep(dt2.Rows(y).Item("s11").ToString) & ","

            data = Left(data, Len(data) - 1)
            sw.WriteLine(data)
            data = ""

        Next

        sw.WriteLine("")
        sw.Close()

        '匯出檔案
        Response.Clear()
        Response.ContentType = "application/download"
        Response.AddHeader("content-disposition", "attachment;filename=" & filename)
        Response.WriteFile(path)
        Response.End()
    End Sub

    Protected Sub UserSel_DataBound(sender As Object, e As System.EventArgs) Handles UserSel.DataBound
        Dim DC As SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = ""
        Dim UnitFlag As String = ""
        Dim ParentFlag As String = ""

        '判斷是否有處單位人事員
        DC = New SQLDBControl
        strSql = "SELECT Role_Num FROM ROLEGROUPITEM WHERE Group_Uid = 'JKGJZ4439V' AND employee_id = '" & user_id & "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read() Then
                UnitFlag = "Y"
            End If
        End If
        DC.Dispose()

        '判斷是否有下一級單位
        DC = New SQLDBControl
        strSql = "SELECT ORG_UID FROM ADMINGROUP WHERE PARENT_ORG_UID = '" & org_uid & "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read() Then
                ParentFlag = "Y"
            End If
        End If
        DC.Dispose()

        Dim ddl As DropDownList = CType(sender, DropDownList)        
        ''設定預設下拉資料                
        If ddl.Items.Contains(ddl.Items.FindByValue(LCase(Session("user_id")))) Then
            ddl.SelectedValue = LCase(Session("user_id"))
        ElseIf ddl.Items.Contains(ddl.Items.FindByValue(UCase(Session("user_id")))) Then
            ddl.SelectedValue = UCase(Session("user_id"))
        End If

        '系統管理員
        If (Not Session("Role") = "1") And Not (ParentFlag = "Y") And (Not UnitFlag = "Y") Then
            UserSel.Enabled = False
        End If
    End Sub
End Class
