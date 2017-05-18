Imports System.Data
Imports System.Data.SqlClient
Imports System.IO

Partial Class Source_01_MOA01001
    Inherits System.Web.UI.Page

    Public sel_day As String = ""
    Public do_sql As New C_SQLFUN
    Public h_count As New C_DATESUM
    Public d_pub As New C_Public
    Dim dr As System.Data.DataRow
    Dim n_table As New System.Data.DataTable
    Dim str_ORG_UID As String = ""
    Dim str_app_id As String = "" '申請人身份證字號
    Public str_EFORMSN As String = ""
    Public eformid As String = ""
    Public read_only As String = ""
    Public print_file As String = ""
    '---for date
    Public date1_flag As Boolean = False
    Public date1_x As Integer = 297
    Public date1_y As Integer = 154
    Public date2_flag As Boolean = False
    Public date2_x As Integer = 702
    Public date2_y As Integer = 161
    Dim connstr, user_id, org_uid As String
    Dim AgentEmpuid As String = ""

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged
        Txt_nSTARTTIME.Text = Calendar1.SelectedDate.ToString("yyyy/MM/dd")
        Call CHECK_NSTART()
        Call count_day()
        If do_sql.G_errmsg <> "" Then
            Txt_nSTARTTIME.Text = ""
        End If

        Div_grid.Visible = False

    End Sub
    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged
        Txt_nENDTIME.Text = Calendar2.SelectedDate.ToString("yyyy/MM/dd")
        Call CHECK_NSTART()
        Call count_day()
        If do_sql.G_errmsg <> "" Then
            Txt_nENDTIME.Text = ""
        End If

        Div_grid2.Visible = False
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        '取得登入者帳號
        If Page.User.Identity.Name.ToString.IndexOf("\") > 0 Then

            Dim LoginAll As String = Page.User.Identity.Name.ToString

            Dim LoginID() As String = Split(LoginAll, "\")

            user_id = LoginID(1)
        Else
            user_id = Page.User.Identity.Name.ToString
        End If

        Dim stmt As String = ""
        Dim p As Integer = 0
        Dim K As Integer = 0
        do_sql.G_errmsg = ""
        do_sql.G_user_id = Session("user_id")
        'user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        eformid = Request("eformid")
        str_EFORMSN = Request("eformsn")
        read_only = Request("read_only")

        'session被清空回首頁
        If user_id = "" And org_uid = "" And str_EFORMSN = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.parent.location='../../index.aspx';")
            Response.Write(" </script>")

        Else
            If str_EFORMSN = "" Then
                'str_EFORMSN = d_pub.randstr(16)  '建立唯一的eformsn
                str_EFORMSN = d_pub.CreateNewEFormSN()      '加入重覆防呆功能 20130710 paul
            End If
            '新增連線
            connstr = do_sql.G_conn_string

            '找出同級單位以下全部單位
            Dim Org_Down As New C_Public

            'do_sql.G_user_id = "tempu178"
            'eformid = "R5Y26NOFA5" '
            'str_EFORMSN = "PL2W6J4OF243157D" 'd_pub.randstr(16)
            'read_only = "1"

            If read_only = "2" Then
                but_exe.Text = "核准"
                If Session("user_id") = "" Then
                    do_sql.G_user_id = Page.User.Identity.Name.Substring(Page.User.Identity.Name.IndexOf("\") + 1)
                End If

            End If

            '找出表單審核者
            If read_only = "2" Then

                '開啟連線
                Dim db As New SqlConnection(connstr)

                db.Open()
                Dim strAgentCheck As New SqlCommand("SELECT empuid FROM flowctl WHERE eformsn = '" & str_EFORMSN & "' and hddate is null", db)
                Dim RdrAgentCheck = strAgentCheck.ExecuteReader()
                If RdrAgentCheck.Read() Then
                    AgentEmpuid = RdrAgentCheck.Item("empuid")
                End If
                db.Close()

            End If

            If read_only = "1" Or read_only = "2" Then
                Button3.Visible = True
                ImgDate1.Visible = False
                ImgDate2.Visible = False
            Else
                Button3.Visible = False
            End If

            If IsPostBack Then
                Exit Sub
            End If

            If read_only = "1" Or read_only = "2" Then
                If select_data() = False Then
                    Exit Sub
                End If
                Exit Sub
            End If

            Lab_time.Text = Now.ToString("yyyy/MM/dd HH:mm:ss")
            If do_sql.select_urname(do_sql.G_user_id) = False Then
                Lab_ORG_NAME_1.Text = ""
                Exit Sub
            End If
            If do_sql.G_usr_table.Rows.Count > 0 Then
                Lab_ORG_NAME_1.Text = do_sql.G_usr_table.Rows(0).Item("ORG_NAME").ToString.Trim
                Lab_emp_chinese_name.Text = do_sql.G_usr_table.Rows(0).Item("emp_chinese_name").ToString.Trim
                Lab_title_name_1.Text = do_sql.G_usr_table.Rows(0).Item("AD_title").ToString.Trim
                str_ORG_UID = do_sql.G_usr_table.Rows(0).Item("ORG_UID").ToString.Trim
            End If
            If str_ORG_UID <> "" Then
                'stmt = "select * from Employee where ORG_UID='" + str_ORG_UID + "' order by emp_chinese_name"
                stmt = "select * from Employee where ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ") order by emp_chinese_name"
                If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
                    Exit Sub
                End If
                If do_sql.G_table.Rows.Count > 0 Then
                    n_table = do_sql.G_table
                    p = 0
                    K = 0
                    DrDown_emp_chinese_name.Items.Clear()

                    'DrDown_emp_chinese_name.DataSource = n_table.Rows
                    'DrDown_emp_chinese_name.DataTextField = "emp_chinese_name"
                    'DrDown_emp_chinese_name.DataValueField = "employee_id"

                    For Each dr In n_table.Rows
                        DrDown_emp_chinese_name.Items.Add(Trim(dr("emp_chinese_name").ToString))
                        DrDown_emp_chinese_name.Items(p).Value = Trim(dr("employee_id").ToString)
                        If UCase(do_sql.G_user_id) = UCase(Trim(dr("employee_id").ToString)) Then
                            K = p
                        End If
                        p += 1

                    Next
                    If p > 0 Then
                        DrDown_emp_chinese_name.SelectedIndex = K
                        'DrDown_emp_chinese_name.SelectedItem.Value = do_sql.G_user_id
                        'If do_sql.select_urname(DrDown_emp_chinese_name.Items(0).Value) = False Then
                        '    Lab_ORG_NAME_2.Text = ""
                        '    Exit Sub
                        'End If
                        'If do_sql.G_usr_table.Rows.Count > 0 Then
                        '    Lab_ORG_NAME_2.Text = do_sql.G_usr_table.Rows(0).Item("ORG_UID").ToString.Trim
                        '    Lab_title_name_2.Text = do_sql.G_usr_table.Rows(0).Item("title_name").ToString.Trim
                        'End If
                        Call DrDown_emp_chinese_name_SelectedIndexChanged(sender, e)
                    End If
                End If
            End If


            Call drow_txt()

            Txt_nSTARTTIME.Text = Now.AddDays(1).ToString("yyyy/MM/dd")
            Txt_nENDTIME.Text = Now.AddDays(1).ToString("yyyy/MM/dd")

            Call count_day()

            'grid
            If str_app_id <> "" Then
                Call select_grid()
            End If

            Dim aidAndComfort As Object = GetAidAndComfort(DrDown_emp_chinese_name.SelectedValue, DateTime.Now.Year) '取得申請人今年度之慰勞假天數
            Dim strAidAndComfort As String = "(尚未設定)"
            If Not aidAndComfort Is Nothing Then
                strAidAndComfort = aidAndComfort.ToString() + " 日"
            End If
            lbAidAndComfort.Text = "本年度慰勞假日數：" + strAidAndComfort

        End If


    End Sub

    ''' <summary>
    ''' 取得目標人員年度慰勞假天數
    ''' </summary>
    ''' <param name="employeeID">目標人員ID</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetAidAndComfort(ByVal employeeID As String, ByVal year As Integer) As String
        '開啟連線
        Dim db As New SqlConnection(do_sql.G_conn_string)
        Dim i As String

        db.Open()
        Dim comm As New SqlCommand("SELECT Holidays FROM P_0106 WHERE employee_id = @employee_id AND Year = @Year AND Status = 'LATEST'", db)
        comm.Parameters.Add("@employee_id", Data.SqlDbType.VarChar).Value = employeeID
        comm.Parameters.Add("@Year", Data.SqlDbType.Int).Value = year
        i = comm.ExecuteScalar()
        db.Close()

        Return i
    End Function

    Sub select_grid()

        Dim stmt As String
        Dim td_year As String = Now.ToString("yyyy")

        'stmt = "select A.State_Name,sum(isnull(B.nDAY,0)*8 + isnull(B.nHOUR,0))/8 As nday ,"
        'stmt += "sum(isnull(B.nDAY,0)*8 + isnull(B.nHOUR,0)) % 8 as nhour "
        'stmt += "from SYSKIND A left outer join  "
        'stmt += " (select D.* from P_01 D,flowctl C "
        'stmt += " where D.EFORMSN=C.EFORMSN and D.PAIDNO ='" + str_app_id + "' and year(nENDTIME)='" + td_year + "' and c.gonogo='E' and D.EFORMSN not in (select nEFORMSN from p_0101 where PENDFLAG='E') ) B"
        'stmt += " on A.State_Name=B.ntype "
        'stmt += "where A.kind_num='2' "
        'stmt += " group by A.State_Name"

        stmt = "select A.State_Name,sum(isnull(B.nDAY,0)*8 + isnull(B.nHOUR,0))/8 As nday ,"
        stmt += "sum(isnull(B.nDAY,0)*8 + isnull(B.nHOUR,0)) % 8 as nhour "
        stmt += "from SYSKIND A left outer join  "
        stmt += " (select * from P_01 D "
        stmt += " where D.PAIDNO ='" + str_app_id + "' and year(nENDTIME)='" + td_year + "' and PENDFLAG ='E' and D.EFORMSN not in (select nEFORMSN from p_0101 where PENDFLAG='E') ) B"
        stmt += " on A.State_Name=B.ntype "
        stmt += "where A.kind_num='2' "
        stmt += " group by A.State_Name, A.State_Num order by A.State_Num"

        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Sub
        End If
        GridView1.DataSource = do_sql.G_table

        GridView1.DataBind()
    End Sub

    Protected Sub DrDown_emp_chinese_name_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrDown_emp_chinese_name.PreRender

        If read_only = "" Then
            str_app_id = DrDown_emp_chinese_name.SelectedValue
        End If

        Call select_grid()

    End Sub

    Protected Sub DrDown_emp_chinese_name_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrDown_emp_chinese_name.SelectedIndexChanged

        'str_app_id = ""

        If do_sql.select_urname(DrDown_emp_chinese_name.Items(DrDown_emp_chinese_name.SelectedIndex).Value) = False Then
            Lab_ORG_NAME_2.Text = ""
            Exit Sub
        End If

    End Sub


    Sub count_day()
        If Txt_nSTARTTIME.Text = "" Or Txt_nENDTIME.Text = "" Then
            Label22.Text = ""
            Label25.Text = ""
            Exit Sub
        End If
        If CDate(Txt_nSTARTTIME.Text).ToString("yyyy/MM/dd") > CDate(Txt_nENDTIME.Text).ToString("yyyy/MM/dd") Then
            Txt_nENDTIME.Text = Txt_nSTARTTIME.Text
        End If
        Dim begin_date As String = CDate(Txt_nSTARTTIME.Text).ToString("yyyy/MM/dd") + " " + DrDown_nSTHOUR.Items(DrDown_nSTHOUR.SelectedIndex).Value
        Dim end_date As String = CDate(Txt_nENDTIME.Text).ToString("yyyy/MM/dd") + " " + DrDown_nETHOUR.Items(DrDown_nETHOUR.SelectedIndex).Value
        Dim t_day As Integer = 0
        Dim t_hour As Integer = 0
        If h_count.date_sum(begin_date, end_date, t_day, t_hour, do_sql.G_errmsg) = False Then

        End If
        Label22.Text = t_day.ToString
        Label25.Text = t_hour.ToString
    End Sub

    Protected Sub DrDown_nSTHOUR_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrDown_nSTHOUR.SelectedIndexChanged
        Call count_day()
    End Sub

    Protected Sub DrDown_nETHOUR_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrDown_nETHOUR.SelectedIndexChanged
        Call count_day()
    End Sub

    Protected Sub but_exe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles but_exe.Click

        Dim FC As New C_FlowSend.C_FlowSend

        Dim SendVal As String = ""

        '判斷是否為批核表單
        If AgentEmpuid = "" Then
            '須抓取申請人資料,非登入人資料
            'SendVal = eformid & "," & do_sql.G_user_id & "," & str_EFORMSN & "," & "1" & ","
            SendVal = eformid & "," & DrDown_emp_chinese_name.SelectedValue & "," & str_EFORMSN & "," & "1" & ","
        Else
            SendVal = eformid & "," & AgentEmpuid & "," & str_EFORMSN & "," & "1" & ","
        End If

        Dim NextPer As Integer = 0

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '關卡為上一級主管有多少人
        NextPer = FC.F_NextStep(SendVal, connstr)

        Dim Chknextstep As Integer

        '判斷表單關卡
        db.Open()
        Dim strComCheck As New SqlCommand("SELECT nextstep FROM flowctl WHERE eformsn = '" & str_EFORMSN & "' and empuid = '" & user_id & "' and hddate is null", db)
        Dim RdrCheck = strComCheck.ExecuteReader()
        If RdrCheck.read() Then
            Chknextstep = RdrCheck.item("nextstep")
        End If
        db.Close()

        Dim strgroup_id As String = ""

        If read_only <> "" Then

            '判斷上一級主管
            db.Open()
            Dim strUpComCheck As New SqlCommand("SELECT group_id FROM flow WHERE eformid = '" & eformid & "' and stepsid = '" & Chknextstep & "' and eformrole=1 ", db)
            Dim RdrUpCheck = strUpComCheck.ExecuteReader()
            If RdrUpCheck.read() Then
                strgroup_id = RdrUpCheck.item("group_id")
            End If
            db.Close()

        End If

        If NextPer = 0 And read_only = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('無上一級主管');")
            Response.Write(" </script>")

            Exit Sub

        End If

        do_sql.G_errmsg = ""
        If read_only <> "2" Then
            If check_text() = False Then
                Exit Sub
            End If

            Dim stmt As String
            Dim str_nSTATUS As String = ""
            Dim str_nPROVEMENT As String = ""


            stmt = "select EFORMSN from P_01 where EFORMSN='" + str_EFORMSN + "'"
            If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
                Exit Sub
            End If
            If do_sql.G_table.Rows.Count > 0 Then
                do_sql.G_errmsg = "EFORMSN已重複,請重新load form"
                Exit Sub
            End If


            If Radio1.Checked = True Then
                str_nSTATUS = "假前申請"
            Else
                str_nSTATUS = "假後補請"
            End If
            If RadioButton2.Checked = True Then
                str_nPROVEMENT = "有"
            Else
                str_nPROVEMENT = "無"
            End If
            DrDown_nAGENT1_title.SelectedIndex = DrDown_nAGENT1.SelectedIndex
            stmt = "insert into P_01(EFORMSN,PWUNIT,PWTITLE," '表單序號,填表人單位,填表人級職
            stmt += "PWNAME, PWIDNO, PAUNIT," '填表人姓名,填表人身份證字號,申請人單位
            stmt += "PANAME, PATITLE, PAIDNO," '申請人姓名,申請人級職,申請人身份證字號 "
            stmt += "nSTATUS, nAPPTIME, nTYPE, " '申請類別,申請時間,假別
            stmt += "nPROVEMENT,nSTARTTIME,nSTHOUR," '請假證明,請假起始日期,請假起始時數
            stmt += "nENDTIME,nETHOUR,nDAY,nHOUR,"   '請假結束日期,請假結束時數,請假總天數,請假總小時
            stmt += "nAGENT1,nAGENT2,nAGENT3,"       '職務代理人1,職務代理人2,職務代理人3
            stmt += "nPLACE,nPHONE,nREASON,nAGENT1_title) "        '到達地點,聯絡電話,事由
            stmt += " values('" + str_EFORMSN + "','" + Lab_ORG_NAME_1.Text + "','" + Lab_title_name_1.Text + "',N'"
            stmt += Lab_emp_chinese_name.Text + "','" + do_sql.G_user_id + "','" + Lab_ORG_NAME_2.Text + "',N'"
            stmt += DrDown_emp_chinese_name.SelectedItem.Text + "','" + Lab_title_name_2.Text + "','" + DrDown_emp_chinese_name.Text + "','"
            stmt += str_nSTATUS + "',getdate(),'" + DrDown_nTYPE.SelectedItem.Text + "','"
            stmt += str_nPROVEMENT + "','" + Txt_nSTARTTIME.Text + "','" + DrDown_nSTHOUR.SelectedItem.Text + "','"
            stmt += Txt_nENDTIME.Text + "','" + DrDown_nETHOUR.SelectedItem.Text + "'," + Label22.Text + "," + Label25.Text + ",N'"
            stmt += DrDown_nAGENT1.SelectedItem.Text + "',N'" + DrDown_nAGENT2.SelectedItem.Text + "',N'" + DrDown_nAGENT3.SelectedItem.Text + "','"
            stmt += DDL_nPlace.SelectedItem.Text + "','" + TXT_nPHONE.Text + "','" + TXT_nREASON.Text.Replace("'", "") + "','"
            stmt += DrDown_nAGENT1_title.SelectedItem.Text + "')"
            If do_sql.db_exec(stmt, do_sql.G_conn_string) = False Then
                Exit Sub
            End If

        End If

        If read_only = "2" Then

            Dim strAgentName As String = ""

            '判斷是否為代理人批核
            If UCase(user_id) = UCase(AgentEmpuid) Then

                '增加批核意見
                insComment(txtcomment.Text, str_EFORMSN, user_id)

            Else

                Dim strComment As String = ""

                '找尋批核者姓名
                db.Open()
                Dim strPer As New SqlCommand("SELECT emp_chinese_name FROM EMPLOYEE WHERE employee_id = '" & user_id & "'", db)
                Dim RdPer = strPer.ExecuteReader()
                If RdPer.read() Then
                    strAgentName = RdPer("emp_chinese_name")
                End If
                db.Close()

                strComment = txtcomment.Text & "(此表單已由" & strAgentName & "代理批核)"

                '增加批核意見
                insComment(strComment, str_EFORMSN, AgentEmpuid)

            End If

        End If

        'stmt = "insert into flowctl(flowsn,eformid,eformsn,gonogo) values('78','1','" + str_EFORMSN + "',"
        'stmt += "'E') "
        'If do_sql.db_exec(stmt, do_sql.G_conn_string) = False Then
        '    Exit Sub
        'End If
        'MsgBox("新增成功", 0, "")

        Dim Val_P As String
        Val_P = ""

        '判斷下一關為上一級主管時人數是否超過一人
        If (NextPer > 1 And read_only = "") Or (NextPer > 1 And strgroup_id = "Z860") Then
            'Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & str_EFORMSN & "&SelFlag=1")
            Response.Redirect("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & str_EFORMSN & "&SelFlag=1&AppUID=" + DrDown_emp_chinese_name.SelectedValue)
        Else

            '表單審核
            Val_P = FC.F_Send(SendVal, do_sql.G_conn_string)

            '通知代理人
            If read_only = "2" Then

                '開啟連線
                Dim strPANAME As String = ""
                Dim strnAGENT1 As String = ""
                Dim strnAGENT2 As String = ""
                Dim strnAGENT3 As String = ""
                Dim strnSTARTTIME As String = ""
                Dim strnSTHOUR As String = ""
                Dim strnENDTIME As String = ""
                Dim strnETHOUR As String = ""
                Dim strPWIDNO As String = ""
                Dim strPAIDNO As String = ""
                Dim strPWNAME As String = ""

                '找尋代理人
                db.Open()
                Dim strPerAgent As New SqlCommand("SELECT PWNAME,PANAME,nAGENT1,nAGENT2,nAGENT3,nSTARTTIME,nSTHOUR,nENDTIME,nETHOUR,PWIDNO,PAIDNO FROM P_01 WHERE EFORMSN = '" & str_EFORMSN & "'", db)
                Dim RdPerAgent = strPerAgent.ExecuteReader()
                If RdPerAgent.read() Then
                    strPWNAME = RdPerAgent("PWNAME")
                    strPANAME = RdPerAgent("PANAME")
                    strnAGENT1 = RdPerAgent("nAGENT1")
                    strnAGENT2 = RdPerAgent("nAGENT2")
                    strnAGENT3 = RdPerAgent("nAGENT3")
                    strnSTARTTIME = RdPerAgent("nSTARTTIME")
                    strnSTHOUR = RdPerAgent("nSTHOUR")
                    strnENDTIME = RdPerAgent("nENDTIME")
                    strnETHOUR = RdPerAgent("nETHOUR")
                    strPWIDNO = RdPerAgent("PWIDNO")
                    strPAIDNO = RdPerAgent("PAIDNO")
                End If
                db.Close()

                If strnAGENT1 <> "" Then
                    AgentMail(strnAGENT1, str_EFORMSN, strPANAME, strnSTARTTIME, strnSTHOUR, strnENDTIME, strnETHOUR)
                End If
                If strnAGENT2 <> "" Then
                    AgentMail(strnAGENT2, str_EFORMSN, strPANAME, strnSTARTTIME, strnSTHOUR, strnENDTIME, strnETHOUR)
                End If
                If strnAGENT3 <> "" Then
                    AgentMail(strnAGENT3, str_EFORMSN, strPANAME, strnSTARTTIME, strnSTHOUR, strnENDTIME, strnETHOUR)
                End If

                '送MAIL給被代理人
                If UCase(strPWIDNO) <> UCase(strPAIDNO) Then

                    Dim MOAServer As String = ""
                    Dim SmtpHost As String = ""
                    Dim SystemMail As String = ""
                    Dim MailYN As String = ""
                    MOAServer = FC.F_MailBase("MOAServer", connstr)
                    SmtpHost = FC.F_MailBase("SmtpHost", connstr)
                    SystemMail = FC.F_MailBase("SystemMail", connstr)
                    MailYN = FC.F_MailBase("Mail_Flag", connstr)

                    Dim empemail As String = ""

                    '申請者mail
                    db.Open()
                    Dim strPer As New SqlCommand("SELECT empemail FROM EMPLOYEE WHERE employee_id = '" & strPAIDNO & "'", db)
                    Dim RdPer = strPer.ExecuteReader()
                    If RdPer.read() Then
                        empemail = RdPer("empemail")
                    End If
                    db.Close()

                    '發送Mail給請假代理人
                    Dim MailBody As String = ""
                    MailBody += "申請人:" & strPWNAME & "<br>"
                    MailBody += "被申請人:" & strPANAME & "<br>"
                    MailBody += "請假時間為" & strnSTARTTIME & "日" & strnSTHOUR & "時~" & strnENDTIME & "日" & strnETHOUR & "時" & "<br>"
                    MailBody += "詳細資料，請上行政管理系統查詢<br>"

                    '判斷是否寄送Mail
                    If MailYN = "Y" And empemail <> "" Then
                        FC.F_MailGO(SystemMail, "系統通知", SmtpHost, empemail, "被請假通知", MailBody)
                    End If

                End If

            End If

            'Val_P = "w3Rr68Mi1d,59KK8MCGPRSIM594,1,27218,6534,您所填寫的「註銷申請單」已經傳送給下一關的處理者單位資產管理員：王崑仰,1"
            'Val_P = "4rM2YFP73N," + str_EFORMSN + ",1,27218,6534,您所填寫的「註銷申請單」已經傳送給下一關的處理者單位資產管理員：王崑仰,1"
            ''val = (Request("val"))
            'GetValue = Split(Val, ",")
            'eformid = GetValue(0)
            'eformsn = GetValue(1)
            'gonogo = GetValue(2)
            'stepsid = GetValue(3)
            'goto_steps = GetValue(4)
            'sShowMsg = GetValue(5)
            'eformrole = GetValue(6)
            'MsgBox(Val_P)

            Dim PageUp As String = ""

            If read_only = "" Then
                PageUp = "New"
            End If

            Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & PageUp)
            'do_sql.G_errmsg = "新增成功"

        End If


    End Sub
    Function check_text() As Boolean
        check_text = False
        If Lab_ORG_NAME_2.Text = "" Then
            'MsgBox("申請人姓名必須選取", 0, "")
            do_sql.G_errmsg = "申請人姓名必須選取"
            Exit Function
        End If
        If DrDown_nTYPE.SelectedIndex < 1 Then
            'MsgBox("假別必須選取", 0, "")
            do_sql.G_errmsg = "假別必須選取"
            Exit Function
        End If
        If Label22.Text = "" Or Label25.Text = "" Then
            'MsgBox("請假日期時間必須選取", 0, "")
            do_sql.G_errmsg = "請假日期時間必須選取"
            Exit Function
        End If
        If CInt(Label22.Text) < 1 And CInt(Label25.Text) < 1 Then
            'MsgBox("請假日期時間錯誤", 0, "")
            do_sql.G_errmsg = "請假日期時間錯誤"
            Exit Function
        End If
        If DrDown_nAGENT1.SelectedItem.Text = "" Then
            'MsgBox("職務代理人(一):必須選取", 0, "")
            do_sql.G_errmsg = "職務代理人(一):必須選取"
            Exit Function
        End If
        If DDL_nPlace.SelectedItem.Text = "請選擇" Then
            'MsgBox("到達地點必須選取", 0, "")
            do_sql.G_errmsg = "到達地點必須選取"
            Exit Function
        End If
        'If CInt(Label22.Text) > 2 Then
        '    'MsgBox("請假不能超過2天", 0, "")
        '    do_sql.G_errmsg = "請假不能超過2天"
        '    Exit Function
        'End If
        If TXT_nREASON.Text = "" Then
            do_sql.G_errmsg = "差勤事由必須輸入"
            Exit Function
        End If
        check_text = True
    End Function

    Protected Sub Radio1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Radio1.CheckedChanged
        Call CLEAR_NSTART()
    End Sub
    Sub CLEAR_NSTART()
        If Radio1.Checked = True Then
            Txt_nSTARTTIME.Text = Now.AddDays(1).ToString("yyyy/MM/dd")
            Txt_nENDTIME.Text = Now.AddDays(1).ToString("yyyy/MM/dd")
        Else
            'If Txt_nSTARTTIME.Text = "" Then
            '    Exit Sub
            'End If
            'If Txt_nSTARTTIME.Text >= Now.ToString("yyyy/MM/dd") Then
            Txt_nSTARTTIME.Text = ""
            Txt_nENDTIME.Text = ""
            'End If
        End If
        Call count_day()
    End Sub

    Protected Sub Radio2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Radio2.CheckedChanged
        Call CLEAR_NSTART()
    End Sub
    Sub CHECK_NSTART()
        If Radio1.Checked = True Then
            If Txt_nSTARTTIME.Text < Now.ToString("yyyy/MM/dd") Then
                Txt_nSTARTTIME.Text = ""
                Txt_nENDTIME.Text = ""
                do_sql.G_errmsg = "假前申請必須日期大於等於今天"
            End If
        Else
            If Txt_nSTARTTIME.Text = "" Then
                Exit Sub
            End If
            Txt_nENDTIME.Text = Txt_nSTARTTIME.Text
            ''當日請假視為假後補請
            'If Txt_nSTARTTIME.Text >= Now.ToString("yyyy/MM/dd") Then
            '    Txt_nSTARTTIME.Text = ""
            '    Txt_nENDTIME.Text = ""
            '    do_sql.G_errmsg = "假後補請必須日期小於今天"
            'End If
        End If

    End Sub
    Function select_data() As Boolean
        select_data = False
        Dim stmt As String
        Call drow_txt()
        stmt = "select * from p_01 "
        stmt += "where EFORMSN ='" + str_EFORMSN + "'"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Function
        End If
        If do_sql.G_table.Rows.Count = 0 Then
            Lab_ORG_NAME_1.Text = ""
            Lab_emp_chinese_name.Text = ""
            Lab_title_name_1.Text = ""
            Lab_ORG_NAME_2.Text = ""
            DrDown_emp_chinese_name.Items.Clear()
            DrDown_emp_chinese_name.Items.Add("")
            DrDown_emp_chinese_name.Items(0).Value = ""

            Lab_title_name_2.Text = ""
            Lab_time.Text = ""

            Radio1.Checked = True
            Radio2.Checked = False
            DrDown_nTYPE.SelectedItem.Text = ""
            RadioButton2.Checked = True
            RadioButton3.Checked = False
            Txt_nSTARTTIME.Text = ""
            DrDown_nSTHOUR.SelectedItem.Text = ""
            Txt_nENDTIME.Text = ""
            DrDown_nETHOUR.SelectedItem.Text = ""
            Label22.Text = ""
            Label25.Text = ""
            DrDown_nAGENT1.Items.Clear()
            DrDown_nAGENT1.Items.Add("")
            DrDown_nAGENT1.Items(0).Value = ""
            DrDown_nAGENT2.Items.Clear()
            DrDown_nAGENT2.Items.Add("")
            DrDown_nAGENT2.Items(0).Value = ""
            DrDown_nAGENT3.Items.Clear()
            DrDown_nAGENT3.Items.Add("")
            DrDown_nAGENT3.Items(0).Value = ""

            TXT_nPHONE.Text = ""
            TXT_nREASON.Text = ""
            str_app_id = ""
        Else
            Lab_ORG_NAME_1.Text = do_sql.G_table.Rows(0).Item("PWUNIT").ToString
            Lab_emp_chinese_name.Text = do_sql.G_table.Rows(0).Item("PWNAME").ToString
            Lab_title_name_1.Text = do_sql.G_table.Rows(0).Item("PWTITLE").ToString
            Lab_ORG_NAME_2.Text = do_sql.G_table.Rows(0).Item("PAUNIT").ToString
            DrDown_emp_chinese_name.Items.Clear()
            DrDown_emp_chinese_name.Items.Add(Trim(do_sql.G_table.Rows(0).Item("PANAME").ToString))
            DrDown_emp_chinese_name.Items(0).Value = Trim(do_sql.G_table.Rows(0).Item("PAIDNO").ToString)

            Lab_title_name_2.Text = do_sql.G_table.Rows(0).Item("PATITLE").ToString
            Lab_time.Text = CDate(do_sql.G_table.Rows(0).Item("nAPPTIME").ToString).ToString("yyyy/MM/dd HH:mm:ss")

            If do_sql.G_table.Rows(0).Item("nSTATUS").ToString = "假前申請" Then
                Radio1.Checked = True
                Radio2.Checked = False
            Else
                Radio1.Checked = False
                Radio2.Checked = True
            End If

            DrDown_nTYPE.SelectedItem.Text = do_sql.G_table.Rows(0).Item("nTYPE").ToString
            If do_sql.G_table.Rows(0).Item("nPROVEMENT").ToString = "有" Then
                RadioButton2.Checked = True
                RadioButton3.Checked = False
            Else
                RadioButton2.Checked = False
                RadioButton3.Checked = True
            End If

            Txt_nSTARTTIME.Text = CDate(do_sql.G_table.Rows(0).Item("nSTARTTIME").ToString).ToString("yyyy/MM/dd")
            DrDown_nSTHOUR.SelectedItem.Text = do_sql.G_table.Rows(0).Item("nSTHOUR").ToString
            Txt_nENDTIME.Text = CDate(do_sql.G_table.Rows(0).Item("nENDTIME").ToString).ToString("yyyy/MM/dd")
            DrDown_nETHOUR.SelectedItem.Text = do_sql.G_table.Rows(0).Item("nETHOUR").ToString
            Label22.Text = do_sql.G_table.Rows(0).Item("nDAY").ToString
            Label25.Text = do_sql.G_table.Rows(0).Item("nHOUR").ToString
            DrDown_nAGENT1.Items.Clear()
            DrDown_nAGENT1.Items.Add("")
            DrDown_nAGENT1.Items(0).Value = ""
            DrDown_nAGENT1.Items.Add(do_sql.G_table.Rows(0).Item("nAGENT1").ToString)
            DrDown_nAGENT1.Items(1).Value = do_sql.G_table.Rows(0).Item("nAGENT1").ToString
            DrDown_nAGENT2.Items.Clear()
            DrDown_nAGENT2.Items.Add("")
            DrDown_nAGENT2.Items(0).Value = ""
            DrDown_nAGENT2.Items.Add(do_sql.G_table.Rows(0).Item("nAGENT2").ToString)
            DrDown_nAGENT2.Items(1).Value = do_sql.G_table.Rows(0).Item("nAGENT2").ToString
            DrDown_nAGENT3.Items.Clear()
            DrDown_nAGENT3.Items.Add("")
            DrDown_nAGENT3.Items(0).Value = ""
            DrDown_nAGENT3.Items.Add(do_sql.G_table.Rows(0).Item("nAGENT3").ToString)
            DrDown_nAGENT3.Items(1).Value = do_sql.G_table.Rows(0).Item("nAGENT3").ToString
            DDL_nPlace.SelectedItem.Text = do_sql.G_table.Rows(0).Item("nPLACE").ToString
            TXT_nPHONE.Text = do_sql.G_table.Rows(0).Item("nPHONE").ToString
            TXT_nREASON.Text = do_sql.G_table.Rows(0).Item("nREASON").ToString
            str_app_id = do_sql.G_table.Rows(0).Item("PAIDNO").ToString

            '取得申請人請假年度之慰勞假天數
            Dim aidAndComfort As Object = GetAidAndComfort(DrDown_emp_chinese_name.SelectedValue, CDate(do_sql.G_table.Rows(0).Item("nAPPTIME").ToString).Year)
            ' Dim aidAndComfort As Object = GetAidAndComfort(DrDown_emp_chinese_name.Items(DrDown_emp_chinese_name.SelectedIndex).Value, CDate(do_sql.G_table.Rows(0).Item("nAPPTIME").ToString).Year)
            Dim strAidAndComfort As String = "(尚未設定)"
            If Not aidAndComfort Is Nothing Then
                strAidAndComfort = aidAndComfort.ToString() + " 日"
            End If
            lbAidAndComfort.Text = "本年度慰勞假日數：" + strAidAndComfort
        End If

        If str_app_id <> "" Then
            Call select_grid()
        End If

        select_data = True

    End Function
    Sub disabel_field()

        DrDown_emp_chinese_name.Enabled = False
        Radio1.Enabled = False
        Radio2.Enabled = False
        DrDown_nTYPE.Enabled = False
        RadioButton2.Enabled = False
        RadioButton3.Enabled = False
        Txt_nSTARTTIME.Enabled = False
        DrDown_nSTHOUR.Enabled = False
        Txt_nENDTIME.Enabled = False
        DrDown_nETHOUR.Enabled = False
        DrDown_nAGENT1.SelectedItem.Enabled = False
        DrDown_nAGENT2.SelectedItem.Enabled = False
        DrDown_nAGENT3.SelectedItem.Enabled = False
        TXT_nPHONE.Enabled = False
        TXT_nREASON.Enabled = False

        If read_only = "1" Then
            but_exe.Visible = False
        End If

    End Sub
    Sub drow_txt()
        Dim stmt As String
        Dim p As Integer

        stmt = "select State_Name from SYSKIND where Kind_Num =2 order by State_Num"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Sub
        End If
        n_table = do_sql.G_table
        DrDown_nTYPE.Items.Clear()
        p = 0
        DrDown_nTYPE.Items.Add("")
        DrDown_nTYPE.Items(p).Value = ""
        p += 1
        For Each dr In n_table.Rows
            DrDown_nTYPE.Items.Add(Trim(dr("State_Name").ToString))
            DrDown_nTYPE.Items(p).Value = Trim(dr("State_Name").ToString)
            p += 1
        Next
        DrDown_nSTHOUR.Items.Clear()
        For p = 8 To 16
            DrDown_nSTHOUR.Items.Add(p.ToString("00"))
            DrDown_nSTHOUR.Items(p - 8).Value = p.ToString("00")
        Next
        DrDown_nETHOUR.Items.Clear()
        For p = 9 To 17
            DrDown_nETHOUR.Items.Add(p.ToString("00"))
            DrDown_nETHOUR.Items(p - 9).Value = p.ToString("00")
        Next
        DrDown_nETHOUR.SelectedValue = 17

    End Sub

    Protected Sub Calendar1_VisibleMonthChanged(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.MonthChangedEventArgs) Handles Calendar1.VisibleMonthChanged
        date1_flag = True
    End Sub

    Protected Sub Calendar2_VisibleMonthChanged(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.MonthChangedEventArgs) Handles Calendar2.VisibleMonthChanged
        date2_flag = True
    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim stmt As String
        Dim F_file As String
        Dim F_file2 As String
        Dim F_file_name As String

        stmt = "select * from p_01 "
        stmt += "where EFORMSN ='" + str_EFORMSN + "'"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Sub
        End If
        If do_sql.G_table.Rows.Count > 0 Then
            n_table = do_sql.G_table
        Else
            do_sql.G_errmsg = "查無可印資料"
            Exit Sub
        End If

        '開啟連線
        Dim db As New SqlConnection(connstr)

        Dim strOrgUid As String = ""
        Dim strOrgTopName As String = ""
        Dim strOrgTopParent As String = ""
        Dim strOrgSecName As String = ""
        Dim strUpName As String = ""
        Dim strUpTitle As String = ""

        Dim TY_date As Date = CDate(n_table.Rows(0).Item("nAPPTIME").ToString)
        Dim T_date As String = CStr(CInt(TY_date.ToString("yyyy")) - 1911) + "年"
        Dim BN_date As Date = CDate(n_table.Rows(0).Item("nSTARTTIME").ToString)
        Dim EN_date As Date = CDate(n_table.Rows(0).Item("nENDTIME").ToString)
        Dim PWIDNO As String = n_table.Rows(0).Item("PWIDNO").ToString
        Dim prn_stmt As String
        Dim print As New C_Xprint

        T_date += TY_date.ToString("MM") + "月" + TY_date.ToString("dd") + "日"
        F_file_name = "rpt010010"
        F_file = MapPath("../../rpt/" + F_file_name + ".txt")
        If File.Exists(F_file) = False Then
            do_sql.G_errmsg = "檔案不存在"
            Exit Sub
        End If

        print_file = "../../drs/" + F_file_name + "-" + Now.ToString("mmssff") & ".drs"
        F_file2 = MapPath(print_file)

        print.C_Xprint(F_file.Split("\")(F_file.Split("\").Length - 1), print_file.Split("/")(print_file.Split("/").Length - 1))
        'Call do_sql.inc_file(F_file, F_file2, F_file_name)
        'Call do_sql.print_sdata(F_file2, "/init " + F_file_name + ".txt 1", "/newpage null")
        print.NewPage()

        '找尋填表人單位
        db.Open()
        Dim strPerAgent As New SqlCommand("SELECT ORG_UID FROM EMPLOYEE WHERE employee_id = '" & PWIDNO & "'", db)
        Dim RdPerAgent = strPerAgent.ExecuteReader()
        If RdPerAgent.read() Then
            strOrgUid = RdPerAgent("ORG_UID")
        End If
        db.Close()

        '找出二級單位
        Dim Org_Down As New C_Public

        Dim strTopOrg As String = ""
        strTopOrg = Org_Down.getUporg(strOrgUid, 1)

        '找尋填表人一級單位
        db.Open()
        Dim strTopName As New SqlCommand("SELECT ORG_NAME,PARENT_ORG_UID FROM ADMINGROUP WHERE ORG_UID = '" & strTopOrg & "'", db)
        Dim RdTopName = strTopName.ExecuteReader()
        If RdTopName.read() Then
            strOrgTopName = RdTopName("ORG_NAME")
            strOrgTopParent = RdTopName("PARENT_ORG_UID")
        End If
        db.Close()

        Dim strParentOrg As String = ""
        strParentOrg = Org_Down.getUporg(strOrgUid, 2)

        '找尋填表人二級單位
        db.Open()
        Dim strSecName As New SqlCommand("SELECT ORG_NAME FROM ADMINGROUP WHERE ORG_UID = '" & strParentOrg & "'", db)
        Dim RdSecName = strSecName.ExecuteReader()
        If RdSecName.read() Then
            strOrgSecName = RdSecName("ORG_NAME")
        End If
        db.Close()

        Dim AddOrgTitle As String = ""
        If strOrgTopParent = "56" Then
            AddOrgTitle = "參謀本部"
        End If

        prn_stmt = n_table.Rows(0).Item("nTYPE").ToString
        prn_stmt = AddOrgTitle & strOrgTopName & strOrgSecName & "請(" + prn_stmt.Replace("假", "") + ")" + "假報告單"

        'Call do_sql.print_block(F_file2, "報告單", 0, 0, prn_stmt, 0, 67, 2)
        'Call do_sql.print_block(F_file2, "事由", 0, 0, n_table.Rows(0).Item("nTYPE").ToString, 0, 67, 2)
        'Call do_sql.print_block(F_file2, "起年", 0, 0, do_sql.add_zero(CStr(CInt(BN_date.ToString("yyyy")) - 1911), 3), 0, 67, 2)
        'Call do_sql.print_block(F_file2, "起月", 0, 0, BN_date.ToString("MM"), 0, 67, 2)
        'Call do_sql.print_block(F_file2, "起日", 0, 0, BN_date.ToString("dd"), 0, 67, 2)
        'Call do_sql.print_block(F_file2, "起時", 0, 0, n_table.Rows(0).Item("nSTHOUR").ToString + "00", 0, 67, 2)
        'Call do_sql.print_block(F_file2, "迄年", 0, 0, do_sql.add_zero(CStr(CInt(EN_date.ToString("yyyy")) - 1911), 3), 0, 67, 2)
        'Call do_sql.print_block(F_file2, "迄月", 0, 0, EN_date.ToString("MM"), 0, 67, 2)
        'Call do_sql.print_block(F_file2, "迄日", 0, 0, EN_date.ToString("dd"), 0, 67, 2)
        'Call do_sql.print_block(F_file2, "迄時", 0, 0, n_table.Rows(0).Item("nETHOUR").ToString + "00", 0, 67, 2)
        'Call do_sql.print_block(F_file2, "擬請假期_天", 0, 0, n_table.Rows(0).Item("nDAY").ToString + " 天", 0, 67, 2)
        'Call do_sql.print_block(F_file2, "擬請假期_時", 0, 0, n_table.Rows(0).Item("nHOUR").ToString + " 時", 0, 67, 2)
        'Call do_sql.print_block(F_file2, "級職", 0, 0, n_table.Rows(0).Item("nAGENT1_Title").ToString, 0, 67, 2)
        'Call do_sql.print_block(F_file2, "姓名", 0, 0, n_table.Rows(0).Item("nAGENT1").ToString, 0, 67, 2)
        'Call do_sql.print_block(F_file2, "到達地點", 0, 0, n_table.Rows(0).Item("nPLACE").ToString, 0, 67, 2)
        'Call do_sql.print_block(F_file2, "電話", 0, 0, n_table.Rows(0).Item("nPHONE").ToString, 0, 67, 2)
        '
        'Call do_sql.print_block(F_file2, "捷拓假字", 0, 0, "-------------(" & do_sql.add_zero(CStr(CInt(BN_date.ToString("yyyy")) - 1911), 3) & ")          -------字第                                 號-------------------------------", 0, 67, 2)

        print.Add("報告單", 0, 0, prn_stmt, 0, 67, 2)
        print.Add("事由", 0, 0, n_table.Rows(0).Item("nTYPE").ToString, 0, 67, 2)
        print.Add("起年", 0, 0, do_sql.add_zero(CStr(CInt(BN_date.ToString("yyyy")) - 1911), 3), 0, 67, 2)
        print.Add("起月", 0, 0, BN_date.ToString("MM"), 0, 67, 2)
        print.Add("起日", 0, 0, BN_date.ToString("dd"), 0, 67, 2)
        print.Add("起時", 0, 0, n_table.Rows(0).Item("nSTHOUR").ToString + "00", 0, 67, 2)
        print.Add("迄年", 0, 0, do_sql.add_zero(CStr(CInt(EN_date.ToString("yyyy")) - 1911), 3), 0, 67, 2)
        print.Add("迄月", 0, 0, EN_date.ToString("MM"), 0, 67, 2)
        print.Add("迄日", 0, 0, EN_date.ToString("dd"), 0, 67, 2)
        print.Add("迄時", 0, 0, n_table.Rows(0).Item("nETHOUR").ToString + "00", 0, 67, 2)
        print.Add("擬請假期_天", 0, 0, n_table.Rows(0).Item("nDAY").ToString + " 天", 0, 67, 2)
        print.Add("擬請假期_時", 0, 0, n_table.Rows(0).Item("nHOUR").ToString + " 時", 0, 67, 2)
        print.Add("級職", 0, 0, n_table.Rows(0).Item("nAGENT1_Title").ToString, 0, 67, 2)
        print.Add("姓名", 0, 0, n_table.Rows(0).Item("nAGENT1").ToString, 0, 67, 2)
        print.Add("到達地點", 0, 0, n_table.Rows(0).Item("nPLACE").ToString, 0, 67, 2)
        print.Add("電話", 0, 0, n_table.Rows(0).Item("nPHONE").ToString, 0, 67, 2)

        print.Add("捷拓假字", 0, 0, "-------------(" & do_sql.add_zero(CStr(CInt(BN_date.ToString("yyyy")) - 1911), 3) & ")          -------字第                                 號-------------------------------", 0, 67, 2)

        prn_stmt = "本屬 " + n_table.Rows(0).Item("PANAME").ToString.Trim + "因"
        prn_stmt += n_table.Rows(0).Item("nTYPE").ToString.Trim + "准假給假自"
        prn_stmt += do_sql.add_zero(CStr(CInt(BN_date.ToString("yyyy")) - 1911), 3) + "年" + BN_date.ToString("MM") + "月"
        prn_stmt += BN_date.ToString("dd") + "日" + n_table.Rows(0).Item("nSTHOUR").ToString + "時起至"
        prn_stmt += do_sql.add_zero(CStr(CInt(EN_date.ToString("yyyy")) - 1911), 3) + "年" + EN_date.ToString("MM") + "月"
        prn_stmt += EN_date.ToString("dd") + "日" + n_table.Rows(0).Item("nETHOUR").ToString + "時止共計 "
        prn_stmt += n_table.Rows(0).Item("nDAY").ToString + " 天 "
        prn_stmt += n_table.Rows(0).Item("nHOUR").ToString + " 時"

        'Call do_sql.print_block(F_file2, "姓名假期", 0, 0, prn_stmt)
        print.Add("姓名假期", prn_stmt, 0, 0)

        '找尋二級單位主管名稱
        db.Open()
        Dim strPerName As New SqlCommand("SELECT emp_chinese_name,AD_TITLE FROM EMPLOYEE WHERE ORG_UID = '" & strParentOrg & "'", db)
        Dim RdPerName = strPerName.ExecuteReader()
        If RdPerName.read() Then
            strUpName = RdPerName("emp_chinese_name")
            strUpTitle = RdPerName("AD_TITLE")
        End If
        db.Close()

        prn_stmt = strUpTitle & "  " & Left(strUpName, 1) & " ○ ○"

        'Call do_sql.print_block(F_file2, "張Sir", 0, 0, prn_stmt)
        'Call do_sql.print_block(F_file2, "日期", 0, 0, T_date)

        print.Add("張Sir", prn_stmt, 0, 0)
        print.Add("日期", T_date, 0, 0)
        print.EndFile()

    End Sub


    Protected Sub backBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles backBtn.Click

        Try

            If read_only = "2" Then
                '增加批核意見
                insComment(txtcomment.Text, str_EFORMSN, user_id)
            End If

            Dim Val_P As String
            Val_P = ""

            '表單駁回
            Dim FC As New C_FlowSend.C_FlowSend

            Dim SendVal As String = ""

            '判斷是否為代理人批核的表單
            If AgentEmpuid = "" Then
                SendVal = eformid & "," & user_id & "," & str_EFORMSN & "," & "1"
            Else
                SendVal = eformid & "," & AgentEmpuid & "," & str_EFORMSN & "," & "1"
            End If

            Val_P = FC.F_Back(SendVal, connstr)

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('表單已駁回給申請人');")
            '重新整理頁面
            Response.Write(" window.dialogArguments.location='../00/MOA00010.aspx';")
            Response.Write(" window.close();")
            Response.Write(" </script>")

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub tranBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tranBtn.Click
        Try
            If GetSuperiors(user_id).Rows.Count > 1 Then '若上一級主管為兩人以上,則切換至選擇欲呈轉主管之頁面
                Server.Transfer("../00/MOA00014.aspx?eformsn=" + str_EFORMSN)
            Else
                If read_only = "2" Then
                    '增加批核意見
                    insComment(txtcomment.Text, str_EFORMSN, user_id)
                End If

                Dim Val_P As String
                Val_P = ""

                '表單呈轉
                Dim FC As New C_FlowSend.C_FlowSend

                Dim SendVal As String = ""

                '判斷是否為代理人批核的表單
                If AgentEmpuid = "" Then
                    SendVal = eformid & "," & user_id & "," & str_EFORMSN & "," & "1"
                Else
                    SendVal = eformid & "," & AgentEmpuid & "," & str_EFORMSN & "," & "1"
                End If

                Val_P = FC.F_Transfer(SendVal, connstr)

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('表單已呈轉給上一級主管');")
                '重新整理頁面
                Response.Write(" window.dialogArguments.location='../00/MOA00010.aspx';")
                Response.Write(" window.close();")
                Response.Write(" </script>")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Function GetSuperiors(ByVal employee_id As String) As DataTable
        Dim db As New SqlConnection(connstr)
        Dim ds As New DataSet()

        db.Open()
        Dim comm As New SqlCommand("select employee_id,emp_chinese_name from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = (select ORG_UID from EMPLOYEE where employee_id = @employee_id))", db)
        comm.Parameters.Add("@employee_id", Data.SqlDbType.VarChar, 10).Value = employee_id.Trim()
        Dim da As New SqlDataAdapter(comm)
        da.Fill(ds)
        db.Close()
        Return ds.Tables(0)
    End Function

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        Try

            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '單純讀取表單不可送件
            If read_only = "1" Then
                but_exe.Visible = False
                backBtn.Visible = False
                tranBtn.Visible = False

                '讀取表單可送件
            ElseIf read_only = "2" Then

                Dim PerCount As Integer = 0

                '判斷批核者是否為副主管,假如批核者為副主管則不可送件只可呈轉
                Dim ParentFlag As String = ""
                Dim ParentVal = user_id & "," & str_EFORMSN

                Dim FC As New C_FlowSend.C_FlowSend
                ParentFlag = FC.F_NextChief(ParentVal, connstr)

                If ParentFlag = "1" Then
                    'but_exe.Visible = False

                    '上一級主管多少人
                    db.Open()
                    Dim PerCountCom As New SqlCommand("select count(empuid) as PerCount from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = (select ORG_UID from EMPLOYEE where employee_id ='" & user_id & "'))", db)
                    Dim PerRdv = PerCountCom.ExecuteReader()
                    If PerRdv.read() Then
                        PerCount = PerRdv("PerCount")
                    End If
                    db.Close()

                    '上一級沒人則呈現送件按鈕
                    If PerCount = 0 Then
                        but_exe.Visible = True
                        tranBtn.Visible = False
                    Else
                        'but_exe.Visible = False
                        but_exe.Visible = True
                    End If

                End If

                '找出同級單位以下全部單位
                Dim Org_Down As New C_Public

                '沒有批核權限不可執行動作
                If (Org_Down.ApproveAuth(str_EFORMSN, user_id)) = "" Then
                    but_exe.Visible = False
                    backBtn.Visible = False
                    tranBtn.Visible = False
                End If

            Else
                backBtn.Visible = False
                tranBtn.Visible = False
            End If

            If read_only = "" Then

                '找出同級單位以下全部單位
                Dim Org_Down As New C_Public

                If do_sql.G_usr_table.Rows.Count > 0 Then
                    Lab_ORG_NAME_2.Text = do_sql.G_usr_table.Rows(0).Item("ORG_NAME").ToString.Trim
                    Lab_title_name_2.Text = do_sql.G_usr_table.Rows(0).Item("AD_title").ToString.Trim
                    str_ORG_UID = do_sql.G_usr_table.Rows(0).Item("ORG_UID").ToString.Trim
                    str_app_id = do_sql.G_usr_table.Rows(0).Item("employee_id").ToString.Trim
                End If

                Dim stmt As String = ""
                Dim CHINE_NAME As String = ""
                Dim P As Integer = 0

                DrDown_nAGENT1.Items.Clear()
                DrDown_nAGENT2.Items.Clear()
                DrDown_nAGENT3.Items.Clear()
                DrDown_nAGENT1_title.Items.Clear()
                DrDown_nAGENT1.Items.Add("")
                DrDown_nAGENT1.Items(0).Value = ""
                DrDown_nAGENT2.Items.Add("")
                DrDown_nAGENT2.Items(0).Value = ""
                DrDown_nAGENT3.Items.Add("")
                DrDown_nAGENT3.Items(0).Value = ""
                DrDown_nAGENT1_title.Items.Add("")
                DrDown_nAGENT1_title.Items(0).Value = ""

                'stmt = "select A.*,B.title_name from Employee A,Titles B where A.title_id=B.title_id and A.ORG_UID='" + str_ORG_UID + "' order by A.emp_chinese_name"
                stmt = "select * from Employee where ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ") order by emp_chinese_name"
                If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
                    Exit Sub
                End If

                n_table = do_sql.G_table
                P = 0
                CHINE_NAME = DrDown_emp_chinese_name.Items(DrDown_emp_chinese_name.SelectedIndex).Value
                For Each dr In n_table.Rows
                    If CHINE_NAME <> Trim(dr("employee_id").ToString) And LeaveCheck(Trim(dr("employee_id").ToString), Txt_nSTARTTIME.Text, DrDown_nSTHOUR.SelectedValue, Txt_nENDTIME.Text, DrDown_nETHOUR.SelectedValue) = "" Then
                        DrDown_nAGENT1.Items.Add(Trim(dr("emp_chinese_name").ToString))
                        DrDown_nAGENT1.Items(P).Value = Trim(dr("employee_id").ToString)
                        DrDown_nAGENT2.Items.Add(Trim(dr("emp_chinese_name").ToString))
                        DrDown_nAGENT2.Items(P + 1).Value = Trim(dr("employee_id").ToString)
                        DrDown_nAGENT3.Items.Add(Trim(dr("emp_chinese_name").ToString))
                        DrDown_nAGENT3.Items(P + 1).Value = Trim(dr("employee_id").ToString)
                        DrDown_nAGENT1_title.Items.Add(Trim(dr("AD_title").ToString))
                        DrDown_nAGENT1_title.Items(P).Value = Trim(dr("AD_title").ToString)
                        P += 1
                    End If
                Next

            Else

                Call disabel_field()

                DrDown_nAGENT1.Enabled = False
                DrDown_nAGENT2.Enabled = False
                DrDown_nAGENT3.Enabled = False
                DDL_nPlace.Enabled = False

            End If


        Catch ex As Exception

        End Try

    End Sub

    Public Function insComment(ByVal strComment As String, ByVal strEformsn As String, ByVal strid As String)

        '新增批核意見
        Try

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '新增批何意見
            db.Open()
            Dim insCom As New SqlCommand("UPDATE flowctl SET comment='" & strComment & "' WHERE hddate IS NULL AND eformsn='" & strEformsn & "' AND empuid='" & strid & "'", db)

            insCom.ExecuteNonQuery()
            db.Close()

        Catch ex As Exception
        End Try

        insComment = ""

    End Function

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "200px"
        Div_grid.Style("left") = "310px"

    End Sub


    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click

        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "200px"
        Div_grid2.Style("left") = "310px"

    End Sub

    Public Function LeaveCheck(ByVal strUserID As String, ByVal strSDate As String, ByVal strSHour As String, ByVal strEDate As String, ByVal strEHour As String)

        '判斷人員是否請假
        Try

            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            Dim nSTARTTIME As String = ""
            Dim nENDTIME As String = ""
            Dim nSTHOUR As String = ""
            Dim nETHOUR As String = ""

            Dim strsql As String = "select * from P_01 "
            strsql += " WHERE ((nSTARTTIME BETWEEN '" & strSDate & "' AND '" & strEDate & "') OR (nENDTIME BETWEEN '" & strSDate & "' AND '" & strEDate & "')) "
            strsql += " AND PAIDNO = '" & strUserID & "'"
            '-:申請 0:駁回 1:送件 ?:審核中 E:完成 G:補登 B:撤銷 R:重新分派 T:呈轉
            strsql += " AND PENDFLAG <> '0' AND PENDFLAG <> 'B'"

            '是否請假
            db.Open()
            Dim PerCountCom As New SqlCommand(strsql, db)
            Dim PerRdv = PerCountCom.ExecuteReader()
            If PerRdv.read() Then
                nSTARTTIME = PerRdv("nSTARTTIME")
                nENDTIME = PerRdv("nENDTIME")
                nSTHOUR = PerRdv("nSTHOUR")
                nETHOUR = PerRdv("nETHOUR")

                '已請假同一天
                If CDate(nSTARTTIME) = CDate(nENDTIME) Then

                    If CDate(strSDate) = CDate(strEDate) Then
                        If (strSHour >= nSTHOUR And strSHour < nETHOUR) Or (strEHour > nSTHOUR And strEHour <= nETHOUR) Then
                            LeaveCheck = "Find"
                        Else
                            LeaveCheck = ""
                        End If
                    Else

                        If CDate(strSDate) = CDate(nSTARTTIME) Then

                            If strSHour >= nETHOUR Then
                                LeaveCheck = ""
                            Else
                                LeaveCheck = "Find"
                            End If

                        ElseIf CDate(strEDate) = CDate(nSTARTTIME) Then

                            If strEHour <= nSTHOUR Then
                                LeaveCheck = ""
                            Else
                                LeaveCheck = "Find"
                            End If
                        Else
                            LeaveCheck = "Find"
                        End If

                    End If

                Else

                    If CDate(strSDate) = CDate(nENDTIME) Then
                        If strSHour >= nETHOUR Then
                            LeaveCheck = ""
                        Else
                            LeaveCheck = "Find"
                        End If
                    ElseIf CDate(strEDate) = CDate(nSTARTTIME) Then
                        If strEHour <= nSTHOUR Then
                            LeaveCheck = ""
                        Else
                            LeaveCheck = "Find"
                        End If
                    Else
                        LeaveCheck = "Find"
                    End If

                End If

            Else
                LeaveCheck = ""
            End If
            db.Close()

        Catch ex As Exception
            LeaveCheck = ex.Message
        End Try

    End Function

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

        Div_grid.Visible = False

    End Sub

    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click

        Div_grid2.Visible = False

    End Sub

    Public Function AgentMail(ByVal strnAGENT As String, ByVal eformsn As String, ByVal strName As String, ByVal SDate As String, ByVal STime As String, ByVal EDate As String, ByVal ETime As String)

        '請假單寄送代理人Mail
        Try

            '開啟連線
            Dim db As New SqlConnection(connstr)

            Dim FC As New C_FlowSend.C_FlowSend

            Dim MOAServer As String = ""
            Dim SmtpHost As String = ""
            Dim SystemMail As String = ""
            Dim MailYN As String = ""
            MOAServer = FC.F_MailBase("MOAServer", connstr)
            SmtpHost = FC.F_MailBase("SmtpHost", connstr)
            SystemMail = FC.F_MailBase("SystemMail", connstr)
            MailYN = FC.F_MailBase("Mail_Flag", connstr)

            Dim emp_org As String = ""

            '找出登入者單位
            db.Open()
            Dim strPerOrg As New SqlCommand("SELECT ORG_UID FROM EMPLOYEE WHERE employee_id = '" & user_id & "'", db)
            Dim RdPerOrg = strPerOrg.ExecuteReader()
            If RdPerOrg.read() Then
                emp_org = RdPerOrg("ORG_UID")
            End If
            db.Close()

            '找出同級單位以下全部單位
            Dim Org_Down As New C_Public

            Dim empemail As String = ""

            '找尋代理人mail
            db.Open()
            Dim strPer As New SqlCommand("SELECT empemail FROM EMPLOYEE WHERE ORG_UID IN (" & Org_Down.getchildorg(emp_org) & ") AND emp_chinese_name = N'" & strnAGENT & "'", db)
            Dim RdPer = strPer.ExecuteReader()
            If RdPer.read() Then
                empemail = RdPer("empemail")
            End If
            db.Close()

            Dim strPath As String = ""
            strPath = "MOA00020.aspx?x=MOA01001&y=YAqBTxRP8P&Read_Only=1&EFORMSN=" & eformsn

            '發送Mail給請假代理人
            Dim MailBody As String = ""
            MailBody += "請假人:" & strName & "<br>"
            MailBody += "代理人:" & strnAGENT & "<br>"
            MailBody += "代理時間:" & SDate & "日" & STime & "時~" & EDate & "日" & ETime & "時" & "<br>"
            MailBody += "<a href='" & MOAServer & strPath & "'>請假申請單</a><br>"

            '判斷是否寄送Mail
            If MailYN = "Y" And empemail <> "" Then
                FC.F_MailGO(SystemMail, "系統通知", SmtpHost, empemail, "請假代理人通知", MailBody)
            End If

        Catch ex As Exception
        End Try

        AgentMail = ""

    End Function

    Protected Sub GridView2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView2.SelectedIndexChanged
        txtcomment.Text = GridView2.Rows(GridView2.SelectedRow.RowIndex).Cells(0).Text
        Div_grid10.Visible = False
    End Sub

    Protected Sub But_PHRASE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles But_PHRASE.Click
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        Dim Visitors As Integer
        Dim str_SQL As String = ""

        '否有申請過請假表單
        db.Open()
        Dim PerCountCom As New SqlCommand("SELECT count(*) as Visitors from Phrase WHERE employee_id = '" & Session("user_id") & "'", db)
        Dim PerRdv = PerCountCom.ExecuteReader()
        If PerRdv.read() Then
            Visitors = PerRdv("Visitors")
        End If
        db.Close()

        If Visitors > 0 Then   '是否需要依類別分類;如請假,派車等
            Div_grid10.Visible = True
            Div_grid10.Style("Top") = "475px"
            Div_grid10.Style("left") = "350px"
        Else

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('無個人批核片語資料,請至批核片語管理新增資料');")
            Response.Write(" </script>")

        End If
    End Sub

    Protected Sub Btn_PHclose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Btn_PHclose.Click
        Div_grid10.Visible = False
    End Sub
End Class
