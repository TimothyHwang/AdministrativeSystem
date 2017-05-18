Imports System.Data.SqlClient
Imports System.Data
Imports System.IO
Partial Class Source_00_MOA00200
    Inherits System.Web.UI.Page

    Public do_sql As New C_SQLFUN
    Dim dr As System.Data.DataRow
    Dim n_table As New System.Data.DataTable
    Dim conn As New C_SQLFUN
    Public OrgChange As String      '判斷組織是否變更
    Dim user_id, org_uid As String
    Dim UnitFlag As String = ""
    Dim ParentFlag As String = ""
    Dim STMT As String = ""
    Public print_file As String = ""
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        If user_id.Length = 0 Then
            Dim tool As New C_Public
            user_id = Server.UrlDecode(C_Public.Decrypt(Request("param1").ToString()))
            org_uid = tool.GetOrgIDByIDNo(user_id)
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

        '判斷是否有處單位人事員
        db.Open()
        Dim strUnit As New SqlCommand("SELECT Role_Num FROM ROLEGROUPITEM WHERE Group_Uid = 'JKGJZ4439V' AND employee_id = '" & user_id & "'", db)
        Dim RdUnit = strUnit.ExecuteReader()
        If RdUnit.read() Then
            UnitFlag = "Y"
        End If
        db.Close()

        If IsPostBack = False Then

            '找出同級單位以下全部單位
            Dim Org_Down As New C_Public

            '找出登入者的一級單位
            Dim strTopOrg As String = ""
            strTopOrg = Org_Down.getUporg(org_uid, 1)

            '找出登入者的二級單位
            Dim strParentOrg As String = ""
            strParentOrg = Org_Down.getUporg(org_uid, 2)

            '判斷登入者權限
            If Session("Role") = "1" Then
                SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
            ElseIf Session("Role") = "2" Then
                SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(strTopOrg) & ") ORDER BY ORG_NAME"
            ElseIf UnitFlag = "Y" Then
                SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(strParentOrg) & ") ORDER BY ORG_NAME"
            Else
                SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ") ORDER BY ORG_NAME"
            End If

            '設定Default日期
            Dim Sdate As TextBox = Me.FindControl("Sdate")
            Dim Edate As TextBox = Me.FindControl("Edate")

            Dim dt As Date = Now()
            If (Sdate.Text = "") Then
                Sdate.Text = dt.AddDays(-7).Date
            End If

            If (Edate.Text = "") Then
                Edate.Text = dt.Date
            End If

        End If

    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click

        '查詢
        Dim strOrd As String = ""

        strOrd = " ORDER BY emp_chinese_name,ShowDay"

        SqlDataSource2.SelectCommand = SQLALL(OrgSel.SelectedValue, UserSel.SelectedValue, Sdate.Text, Edate.Text) & strOrd

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

            If Session("Role") = "1" Or Session("Role") = "2" Or UnitFlag = "Y" Or ParentFlag = "Y" Then
                SqlDataSource3.SelectCommand = "SELECT employee_id, emp_chinese_name, ORG_UID FROM EMPLOYEE WHERE ORG_UID ='" & OrgSel.SelectedValue & "' ORDER BY emp_chinese_name"
            Else
                SqlDataSource3.SelectCommand = "SELECT employee_id, emp_chinese_name, ORG_UID FROM EMPLOYEE WHERE ORG_UID ='" & OrgSel.SelectedValue & "' AND employee_id='" & user_id & "' ORDER BY emp_chinese_name"
            End If

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

            '登入馬上查詢
            ImgSearch_Click(Nothing, Nothing)

        End If

    End Sub

    Public Function SQLALL(ByVal OrgSel, ByVal UserSel, ByVal SDate, ByVal EDate)

        '整合SQL搜尋字串
        Dim strsql As String = ""
        Dim strsel As String = ""
        Dim strSqlAll As String = ""
        Dim strSqlFirst As String = ""
        Dim strSqlLast As String = ""

        '找出同級單位以下全部單位
        Dim Org_Down As New C_Public

        strsql = "SELECT employee_id,emp_chinese_name,CONVERT(nvarchar, CONVERT(datetime,'" & SDate & "'), 111) AS ShowDay,"
        strsql += " SIGNDATE_d,In_Time_nvc,Out_Time_nvc,REASON_nvc,Out_REASON"
        strsql += " FROM T_SIGN_RECORD RIGHT OUTER JOIN EMPLOYEE ON T_SIGN_RECORD.LOGONID_nvc = EMPLOYEE.employee_id AND datediff(dd,SIGNDATE_d,'" & SDate & "') = 0 WHERE 1=1 "

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
            If ParentFlag = "Y" Or UnitFlag = "Y" Or Session("Role") = "2" Then

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

        '加班
        If DrDn_add.SelectedValue <> "請選擇" Then
            strsel += " AND (Out_Time_nvc >= '" & DrDn_add.SelectedValue & "' OR Out_Time_nvc <= '06:00:00')"
        End If

        '簽到
        If DrDn_in.SelectedValue = "已簽到" Then
            strsel += " AND In_Time_nvc is not NULL "
        ElseIf DrDn_in.SelectedValue = "未簽到" Then
            strsel += " AND In_Time_nvc is NULL "
        End If

        '簽退
        If DrDn_out.SelectedValue = "已簽退" Then
            strsel += " AND Out_Time_nvc is not NULL "
        ElseIf DrDn_out.SelectedValue = "未簽退" Then
            strsel += " AND Out_Time_nvc is NULL "
        End If



        strSqlFirst = strsql & strsel

        Dim ShowDate As Integer
        If DateDiff(DateInterval.Day, CDate(SDate), CDate(EDate)) > 0 Then

            For ShowDate = 1 To DateDiff(DateInterval.Day, CDate(SDate), CDate(EDate))

                Dim strSqlUnion As String = ""

                strSqlUnion = "SELECT employee_id,emp_chinese_name,CONVERT(nvarchar, CONVERT(datetime,'" & DateAdd(DateInterval.Day, ShowDate, CDate(SDate)) & "'), 111) AS ShowDay,SIGNDATE_d,In_Time_nvc,Out_Time_nvc,REASON_nvc,Out_REASON FROM T_SIGN_RECORD RIGHT OUTER JOIN EMPLOYEE ON T_SIGN_RECORD.LOGONID_nvc = EMPLOYEE.employee_id AND datediff(dd,SIGNDATE_d,'" & DateAdd(DateInterval.Day, ShowDate, CDate(SDate)) & "') = 0 WHERE 1=1 "

                strSqlLast += " UNION " & strSqlUnion & strsel

            Next

            strSqlAll = strSqlFirst & strSqlLast

        Else
            strSqlAll = strSqlFirst
        End If

        SQLALL = strSqlAll

    End Function

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged

        '分頁
        Dim strOrd As String = ""

        strOrd = " ORDER BY emp_chinese_name,ShowDay"

        SqlDataSource2.SelectCommand = SQLALL(OrgSel.SelectedValue, UserSel.SelectedValue, Sdate.Text, Edate.Text) & strOrd

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        '排序
        Dim strOrd As String = ""

        If GridView1.SortExpression.ToString() = "emp_chinese_name" Then
            strOrd = " ORDER BY " & GridView1.SortExpression.ToString() & ",ShowDay"
        Else
            strOrd = " ORDER BY " & GridView1.SortExpression.ToString() & ",emp_chinese_name"

        End If


        SqlDataSource2.SelectCommand = SQLALL(OrgSel.SelectedValue, UserSel.SelectedValue, Sdate.Text, Edate.Text) & strOrd

    End Sub

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "90px"
        Div_grid.Style("left") = "50px"

        Calendar1.SelectedDate = Sdate.Text

    End Sub

    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click

        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "90px"
        Div_grid2.Style("left") = "180px"

        Calendar2.SelectedDate = Edate.Text

    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

        Sdate.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False

    End Sub

    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged

        Edate.Text = Calendar2.SelectedDate.Date
        Div_grid2.Visible = False

    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

        Div_grid.Visible = False

    End Sub

    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click

        Div_grid2.Visible = False

    End Sub

    Protected Sub ImagePrint_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImagePrint.Click
        Dim F_file As String
        'Dim F_file_Use As String
        Dim F_file2 As String
        Dim F_file_name As String
        Dim strOrd As String = ""
        Dim print As New C_Xprint()
        Try
            STMT = SQLALL(OrgSel.SelectedValue, UserSel.SelectedValue, Sdate.Text, Edate.Text) & " ORDER BY emp_chinese_name,ShowDay"

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



            Dim prn_stmt As String

            F_file_name = "rpt002000"
            F_file = MapPath("../../rpt/" + F_file_name + ".txt")
            'F_file_Use = MapPath("../../rpt/" + F_file_name + "_" + Rnd().ToString + ".txt")
            If File.Exists(F_file) = False Then
                do_sql.G_errmsg = "檔案不存在"
                Exit Sub
                'Else
                'File.Copy(F_file, F_file_Use)
            End If

            'F_file = F_file_Use

            print_file = "../../drs/" + F_file_name + "-" + Now.ToString("mmssff") & ".drs"
            Dim tmpPrint_file As String = print_file.Split("/")(print_file.Split("/").Length - 1)
            F_file2 = MapPath(print_file)

            print.C_Xprint(F_file_name + ".txt", tmpPrint_file)
            print.NewPage()

            Dim iPageCount As Integer = 0
            Dim iDivider As Integer = 8
            Dim iRemainder As Integer = 0
            Dim iTotalPages As Integer = 0
            Dim i_line As Integer = 0
            Dim iDetailNumber As Integer = n_table.Rows.Count

            iTotalPages = Math.DivRem(iDetailNumber, iDivider, iRemainder)
            If iRemainder > 0 Then iTotalPages += 1


            'prt.Add("任務實際里程", dv.Table.Rows(0)("RealMilage").ToString, 0, 0)
            'Call do_sql.inc_file(F_file, F_file2, F_file_name)
            Dim nPage As Integer = 0
            'Dim n_line As Integer = 35
            Dim n_line As Integer = 0
            Dim tmp_add As String = ""
            Dim tmp_name As String = ""
            Dim tmp_month As String = ""
            If n_table.Rows.Count > 0 Then
                nPage = 1
                For Each dr In n_table.Rows
                    If n_line = 34 Then
                        nPage = nPage + 1
                        print.NewPage()
                        n_line = 0
                    End If

                    'If n_line >= 31 Or tmp_name <> dr("emp_chinese_name").ToString Or tmp_month <> Mid(dr("ShowDay").ToString(), 1, 7) Then
                    tmp_name = dr("emp_chinese_name").ToString
                    tmp_month = Mid(dr("ShowDay").ToString(), 1, 7)
                    'nPage += 1
                    'Call do_sql.print_sdata(F_file2, "/init " + F_file_name + ".txt 1", "/newpage null")
                    prn_stmt = CDate(Sdate.Text).ToString("yyyy/MM/dd") & "~" & CDate(Edate.Text).ToString("yyyy/MM/dd")
                    print.Add("日期:", prn_stmt, 0, 0)
                    print.Add("加班:", DrDn_add.SelectedValue, 0, 0)
                    'Call do_sql.print_block(F_file2, "日期:", 0, 0, prn_stmt)
                    'Call do_sql.print_block(F_file2, "加班:", 0, 0, DrDn_add.SelectedValue)
                    'n_line = 0
                    'End If
                    'SELECT employee_id,emp_chinese_name,CONVERT(nvarchar, CONVERT(datetime,'" & SDate & "'), 111) AS ShowDay,SIGNDATE_d,In_Time_nvc,Out_Time_nvc,REASON_nvc
                    'Call do_sql.print_block(F_file2, "人員姓名", 0, n_line * 7, dr("emp_chinese_name").ToString)
                    'Call do_sql.print_block(F_file2, "日期", 0, n_line * 7, dr("ShowDay").ToString())
                    print.Add("人員姓名", dr("emp_chinese_name").ToString, 0, n_line * 7)
                    print.Add("日期", dr("ShowDay").ToString(), 0, n_line * 7)

                    'Call do_sql.print_block(F_file2, "簽到時間", 0, n_line * 7, dr("In_Time_nvc").ToString)
                    print.Add("簽到時間", dr("In_Time_nvc").ToString, 0, n_line * 7)
                    If (dr("Out_Time_nvc").ToString >= DrDn_add.SelectedValue & ":00" Or dr("Out_Time_nvc").ToString <= "06:00:00") And dr("Out_Time_nvc").ToString <> "" Then
                        tmp_add = "1"
                    Else
                        tmp_add = ""
                    End If
                    'Call do_sql.print_block(F_file2, "簽退時間", 0, n_line * 7, dr("Out_Time_nvc").ToString)
                    'Call do_sql.print_block(F_file2, "加班", 0, n_line * 7, tmp_add)
                    print.Add("簽退時間", dr("Out_Time_nvc").ToString, 0, n_line * 7)
                    print.Add("加班", tmp_add, 0, n_line * 7)
                    If (dr("Out_Time_nvc").ToString >= "22:00:00" Or dr("Out_Time_nvc").ToString <= "06:00:00") And dr("Out_Time_nvc").ToString <> "" Then
                        tmp_add = "1"
                    Else
                        tmp_add = ""
                    End If
                    'Call do_sql.print_block(F_file2, "夜點", 0, n_line * 7, tmp_add)
                    print.Add("夜點", tmp_add, 0, n_line * 7)

                    n_line = n_line + 1

                Next
            End If
            print.EndFile()

            'While File.Exists(F_file)
            'File.Delete(F_file)
            'End While
        Catch ex As Exception
            Response.Write(ex.Message)
            Response.End()
        End Try
        

    End Sub

    Protected Sub Img_Export_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles Img_Export.Click
        Dim strsqlAll As String = SQLALL(OrgSel.SelectedValue, UserSel.SelectedValue, Sdate.Text, Edate.Text) & " ORDER BY emp_chinese_name,ShowDay"

        '設定檔案路徑
        Dim path As String = "", filename As String = ""

        path = Server.MapPath("~/Drs/")
        filename = Date.Today.ToString("yyyyMMdd") & "_00200.csv"
        path = path & filename

        Dim colname As String = ""
        Dim data As String = ""
        Dim prn_stmt As String = ""

        Dim sw As New System.IO.StreamWriter(path, False, Encoding.GetEncoding("big5"))

        prn_stmt = (CInt(CDate(Sdate.Text).ToString("yyyy")) - 1911).ToString & " 年 " & CDate(Sdate.Text).ToString("MM") & " 月份"
        prn_stmt += "加班申報表 "

        sw.WriteLine(prn_stmt)

        colname = "人員姓名,日期,簽到時間,簽退時間,遲到理由,加班,夜點,加班理由,"

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

            data += dt2.Rows(y).Item("emp_chinese_name").ToString
            data += "," & dt2.Rows(y).Item("ShowDay").ToString
            data += "," & dt2.Rows(y).Item("In_Time_nvc").ToString
            data += "," & dt2.Rows(y).Item("Out_Time_nvc").ToString
            data += "," & dt2.Rows(y).Item("REASON_nvc").ToString
            data += ","
            If (dt2.Rows(y).Item("Out_Time_nvc").ToString >= DrDn_add.SelectedValue & ":00" Or dt2.Rows(y).Item("Out_Time_nvc").ToString <= "06:00:00") And dt2.Rows(y).Item("Out_Time_nvc").ToString <> "" Then
                tmp_add = "1"
            Else
                tmp_add = ""
            End If
            data += tmp_add & ","
            If (dt2.Rows(y).Item("Out_Time_nvc").ToString >= "22:00:00" Or dt2.Rows(y).Item("Out_Time_nvc").ToString <= "06:00:00") And dt2.Rows(y).Item("Out_Time_nvc").ToString <> "" Then
                tmp_add = "1"
            Else
                tmp_add = ""
            End If
            data += tmp_add
            data += "," & dt2.Rows(y).Item("Out_REASON").ToString
            data += ","
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

    
End Class
