Imports System.Data
Imports System.Data.sqlclient
Imports System.IO
Imports CFlowSend

Partial Class Source_05_MOA05001
    Inherits System.Web.UI.Page
    Public stmt As String
    Public do_sql As New C_SQLFUN
    Dim dr As System.Data.DataRow
    Dim n_table As New System.Data.DataTable
    Dim n_table2 As New System.Data.DataTable
    Dim n_table3 As New System.Data.DataTable
    Dim p As Integer = 0
    Dim K As Integer = 0
    Dim str_ORG_UID As String
    Dim str_app_id As String
    Public eformsn As String
    Public d_pub As New C_Public
    Dim eformid As String
    Public read_only As String = ""
    Public print_file As String = ""
    Dim connstr, user_id, org_uid As String
    Dim AgentEmpuid As String = ""
    Dim AppUID As String

	''' <summary>
    ''' 彈出警示窗並導回首頁登錄
    ''' </summary>
    ''' <param name="sMsg"></param>
    ''' <remarks></remarks>
    Protected Sub AlertAndReDirectIndex(ByVal sMsg As String)
        Response.Write(" <script language='javascript'>")
        Response.Write(" alert('" + sMsg + "');")
        Response.Write(" window.parent.parent.location='../../index.aspx';")
        Response.Write(" </script>")
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

        '新增資料
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '找出同級單位以下全部單位
        Dim Org_Down As New C_Public

        do_sql.G_errmsg = ""
        do_sql.G_user_id = Session("user_id") '"tempu180"
        'user_id = Session("user_id")
        org_uid = Session("ORG_UID")

        eformsn = Request("eformsn") '"VJ" 'd_pub.randstr(2) '
        eformid = Request("eformid")
        read_only = Request("read_only")

        'session被清空回首頁
        If user_id = "" And org_uid = "" And eformsn = "" Then
			AlertAndReDirectIndex("畫面停留太久未使用，將重新整理回首頁")   
        Else

            'do_sql.G_user_id = "tempu178"
            'eformsn = "44XW3GKT9DAJ274Z" 'd_pub.randstr(16) '
            'eformid = "R5Y26NOFA5"
            'read_only = "2"
            If read_only = "2" Then
                But_exe.Text = "核准"
                If Session("user_id") = "" Then
                    do_sql.G_user_id = Page.User.Identity.Name.Substring(Page.User.Identity.Name.IndexOf("\") + 1)
                End If
            End If

            '找出表單審核者
            If read_only = "2" Then

                '開啟連線
                Dim db As New SqlConnection(connstr)

                db.Open()
                Dim strAgentCheck As New SqlCommand("SELECT empuid FROM flowctl WHERE eformsn = '" & eformsn & "' and hddate is null", db)
                Dim RdrAgentCheck = strAgentCheck.ExecuteReader()
                If RdrAgentCheck.read() Then
                    AgentEmpuid = RdrAgentCheck.item("empuid")
                End If
                db.Close()

            End If

            Lab_eformsn.Text = eformsn
            'GridView1.DataBind()

            If read_only = "1" Or read_only = "2" Then
                Button3.Visible = True
                GridView2.Visible = False

            Else
                Button3.Visible = False
            End If
            If IsPostBack Then
                Exit Sub
            End If
            'eformsn = d_pub.randstr(2) 'Request("eformsn") '
            If read_only = "1" Or read_only = "2" Then
                If select_data() = False Then
                    Exit Sub
                End If
                Exit Sub
            End If

            Label8.Text = Now.ToString("yyyy/MM/dd HH:mm:ss")
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

                    For Each dr As DataRow In n_table.Rows
                        DrDown_emp_chinese_name.Items.Add(Trim(dr("emp_chinese_name").ToString))
                        DrDown_emp_chinese_name.Items(p).Value = Trim(dr("employee_id").ToString)
                        If UCase(do_sql.G_user_id) = UCase(Trim(dr("employee_id").ToString)) Then
                            K = p
                        End If
                        p += 1
                    Next
                    If p > 0 Then
                        DrDown_emp_chinese_name.SelectedIndex = K
                        Call DrDown_emp_chinese_name_SelectedIndexChanged(sender, e)
                    End If
                End If
            End If

            Call drow_txt()

            '新增來賓資料:
            stmt = "select A.* from (select nID,max(isnull(ncreatedate,'')) ncreatedate from p_0501 "
            stmt += "where ncreatedate > '" + Now.AddYears(-1).ToString("yyyy/MM/dd") + "'"
            stmt += " and EFORMSN in(select EFORMSN from p_05 where PWIDNO='" + do_sql.G_user_id + "'"
            stmt += " and nAPPTIME > '" + Now.AddYears(-1).ToString("yyyy/MM/dd") + "') "
            stmt += " group by nID) B inner join p_0501 A "
            stmt += " on A.nID=B.nID and A.ncreatedate=B.ncreatedate "
            stmt += " where A.ncreatedate > '" + Now.AddYears(-1).ToString("yyyy/MM/dd") + "'"
            stmt += " and A.EFORMSN in(select EFORMSN from p_05 where PWIDNO='" + do_sql.G_user_id + "'"
            stmt += " and nAPPTIME > '" + Now.AddYears(-1).ToString("yyyy/MM/dd") + "') ORDER BY nName"

            If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
                Exit Sub
            End If
            n_table = do_sql.G_table
            DrDown_guest.Items.Clear()
            p = 0
            DrDown_guest.Items.Add("--請選擇--")
            DrDown_guest.Items(p).Value = "-男--"
            p += 1
            For Each dr In n_table.Rows
                DrDown_guest.Items.Add(Trim(dr("nName").ToString) + "--" + Trim(dr("nID").ToString))
                DrDown_guest.Items(p).Value = Trim(dr("nName").ToString) + "--" + Trim(dr("nSex").ToString) + "--" + Trim(dr("nKind").ToString) + "--" + Trim(dr("nService").ToString) + "--" + Trim(dr("nID").ToString + "--" + Trim(dr("nCarNo").ToString))
                p += 1
            Next
            ''新增Grid資料:
            'stmt = "select * from p_0501 "
            'stmt += "where EFORMSN ='" + eformsn + "'"
            'stmt += " order by ncreatedate"
            'If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            '    Exit Sub
            'End If
            'GridView1.DataSource = do_sql.G_table

            GridView1.DataBind()
            RadioButton1.Checked = True
            RadioButton2.Checked = False


        End If


    End Sub

    
    Sub check_hour()
        If DrDown_nSTARTTIME.Text = "" Then
            Exit Sub
        End If
        If DrDown_nENDTIME.Text = "" Then
            Exit Sub
        End If
        If DrDown_nSTARTTIME.Text >= DrDown_nENDTIME.Text Then
            do_sql.G_errmsg = "會客時間起須小於迄"
            Exit Sub
        End If
    End Sub

    Protected Sub DrDown_nSTARTTIME_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrDown_nSTARTTIME.SelectedIndexChanged
        check_hour()
    End Sub

    Protected Sub DrDown_nENDTIME_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrDown_nENDTIME.SelectedIndexChanged
        check_hour()
    End Sub

    Protected Sub DrDown_emp_chinese_name_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrDown_emp_chinese_name.SelectedIndexChanged
        str_app_id = ""
        If do_sql.select_urname(DrDown_emp_chinese_name.Items(DrDown_emp_chinese_name.SelectedIndex).Value) = False Then
            Lab_ORG_NAME_2.Text = ""
            Exit Sub
        End If

        If do_sql.G_usr_table.Rows.Count > 0 Then
            Dim stmt As String = ""
            Dim CHINE_NAME As String = ""
            Dim P As Integer = 0

            Lab_ORG_NAME_2.Text = do_sql.G_usr_table.Rows(0).Item("ORG_NAME").ToString.Trim
            Lab_title_name_2.Text = do_sql.G_usr_table.Rows(0).Item("AD_title").ToString.Trim
            str_ORG_UID = do_sql.G_usr_table.Rows(0).Item("ORG_UID").ToString.Trim
            str_app_id = do_sql.G_usr_table.Rows(0).Item("employee_id").ToString.Trim
        End If
    End Sub

    Protected Sub DrDown_guest_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrDown_guest.SelectedIndexChanged
        If DrDown_guest.SelectedItem.Text = "" Then
            Txt_nName.Text = ""
            Radio1.Checked = True
            Radio2.Checked = False
            Txt_nService.Text = ""
            Txt_nID.Text = ""
        Else
            Dim m_value
            m_value = Split(DrDown_guest.SelectedItem.Value, "--")
            Txt_nName.Text = m_value(0)
            If m_value(1) = "男" Then
                Radio1.Checked = True
                Radio2.Checked = False
            Else
                Radio1.Checked = False
                Radio2.Checked = True
            End If
            DrDown_nkind.Text = m_value(2)
            Txt_nService.Text = m_value(3)
            Txt_nID.Text = m_value(4)
            txtCarNo.Text = m_value(5)
        End If
    End Sub

    Protected Sub But_exe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles But_exe.Click

        ''Dim FC As New C_FlowSend.C_FlowSend
		Dim FC As New CFlowSend

        Dim SendVal As String = ""

        '判斷是否為代理人批核的表單
        If AgentEmpuid = "" Then
            'SendVal = eformid & "," & user_id & "," & eformsn & "," & "1" & ","
            SendVal = eformid & "," & DrDown_emp_chinese_name.SelectedValue & "," & eformsn & "," & "1" & ","
        Else
            SendVal = eformid & "," & AgentEmpuid & "," & eformsn & "," & "1" & ","
        End If

        Dim NextPer As Integer = 0

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '關卡為上一級主管有多少人
        NextPer = FC.F_NextStep(SendVal, connstr)

        Dim Chknextstep As Integer

        '判斷表單關卡
        db.Open()
        Dim strComCheck As New SqlCommand("SELECT nextstep FROM flowctl WHERE eformsn = '" & eformsn & "' and empuid = '" & user_id & "' and hddate is null", db)
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

        Dim Val_P As String
        Val_P = ""

        '表單審核
        If read_only <> "2" Then

            If check_text() = False Then
                Exit Sub
            End If
            If ins_p501("2") = False Then
                Exit Sub
            End If
            stmt = "insert into P_05(EFORMSN,PWUNIT,PWTITLE," '表單序號,填表人單位,填表人級職
            stmt += "PWNAME, PWIDNO, PAUNIT," '填表人姓名,填表人身份證字號,申請人單位
            stmt += "PANAME, PATITLE, PAIDNO," '申請人姓名,申請人級職,申請人身份證字號 "
            stmt += "nAPPTIME, nRECROOM, " '申請時間,會客室名稱,
            stmt += "nRECEXIT,nPLACE,nPHONE," '會客入口名稱,會客地點,聯絡電話
            stmt += "nRECDATE,nSTARTTIME,nENDTIME,nSMin,nEMin,"   '會客時間(年月日),會客時間(起始-時),會客時間(結束-時),會客時間(起始-分),會客時間(結束-分)
            stmt += "nREASON) "        '事由
            stmt += " values('" + eformsn + "','" + Lab_ORG_NAME_1.Text + "','" + Lab_title_name_1.Text + "',N'"
            stmt += Lab_emp_chinese_name.Text + "','" + do_sql.G_user_id + "','" + Lab_ORG_NAME_2.Text + "',N'"
            stmt += DrDown_emp_chinese_name.SelectedItem.Text + "','" + Lab_title_name_2.Text + "','" + DrDown_emp_chinese_name.Text + "','"
            stmt += Label8.Text + "','" + DrDown_nRECROOM.SelectedItem.Text + "','"
            stmt += DrDown_nRECEXIT.SelectedItem.Text + "','" + TXT_nPLACE.Text + "','" + Txt_nPHONE.Text + "','"
            stmt += Txt_nRECDATE.Text + "','" + DrDown_nSTARTTIME.SelectedItem.Text + "','" + DrDown_nENDTIME.SelectedItem.Text + "','" + DrDown_nSMin.SelectedItem.Text + "','" + DrDown_nEMin.SelectedItem.Text + "','"
            stmt += TXT_nREASON.Text.Replace("'", "") + "')"
            If do_sql.db_exec(stmt, do_sql.G_conn_string) = False Then
                Exit Sub
            End If


            If DateDiff(DateInterval.Day, Now(), CDate(Txt_nRECDATE.Text)) < 0 Then
                '補登
                Val_P = FC.F_Supply(SendVal, do_sql.G_conn_string)

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('補登完成');")
                Response.Write(" window.parent.location='../00/MOA00020.aspx?x=MOA05001&y=U28r13D6EA';")
                Response.Write(" </script>")
            End If


        End If

        If read_only = "2" Then

            Dim strAgentName As String = ""

            '判斷是否為代理人批核
            If UCase(user_id) = UCase(AgentEmpuid) Then

                '增加批核意見
                insComment(txtcomment.Text, eformsn, user_id)

            Else
				If user_id.Length > 0 And AgentEmpuid.Length > 0 Then
					Dim strComment As String = ""

					'找尋批核者姓名
					db.Open()
					Dim strPer As New SqlCommand("SELECT emp_chinese_name FROM EMPLOYEE WHERE employee_id = '" & user_id & "'", db)
					Dim RdPer = strPer.ExecuteReader()
					If RdPer.Read() Then
						strAgentName = RdPer("emp_chinese_name")
					End If
					db.Close()
	
					strComment = txtcomment.Text & "(此表單已由" & strAgentName & "代理批核)."
	
					'增加批核意見
					insComment(strComment, eformsn, AgentEmpuid)
				Else
                    AlertAndReDirectIndex("畫面停留太久未使用，將重新整理回首頁")
                    Exit Sub
                End If
            End If

        End If

        '判斷下一關為上一級主管時人數是否超過一人
        If (NextPer > 1 And read_only = "") Or (NextPer > 1 And strgroup_id = "Z860") Then
            'Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & eformsn & "&SelFlag=1")
            '20131024 Debug Log 問題解決後須回復
            Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & eformsn & "&SelFlag=1&AgentEmpuid=" & AgentEmpuid & "&AppUID=" & DrDown_emp_chinese_name.SelectedValue)
        Else
            '正常
            Val_P = FC.F_Send(SendVal, do_sql.G_conn_string)
	    ''20130822 新增log紀錄
            'If (eformid = "S9QR2W8X6U" Or eformid = "U28r13D6EA") Then
            'If (eformid = "34XI93641R" Or eformid = "U28r13D6EA") Then
            '    Call CFlowSend.WriteAgentRecord(eformid, AgentEmpuid, db)
            '    Call CFlowSend.WriteFlowRecord(eformid, SendVal)
            'End If
            Dim PageUp As String = ""

            If read_only = "" Then
                PageUp = "New"
            End If

            Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & PageUp)
            do_sql.G_errmsg = "新增成功"
        End If



    End Sub
    Function check_text() As Boolean
        check_text = False
        If Lab_ORG_NAME_2.Text = "" Then
            do_sql.G_errmsg = "申請人姓名必須選取"
            Exit Function
        End If
        If DrDown_nRECROOM.SelectedIndex < 1 Then
            do_sql.G_errmsg = "會客室必須選取"
            Exit Function
        End If
        '取消會客入口 2013/10/24
        'If DrDown_nRECEXIT.SelectedIndex < 1 Then
        '    do_sql.G_errmsg = "會客入口必須選取"
        '    Exit Function
        'End If
        If TXT_nPLACE.Text = "" Then
            do_sql.G_errmsg = "地點:必須輸入"
            Exit Function
        End If
        If Txt_nPHONE.Text = "" Then
            do_sql.G_errmsg = "電話必須輸入"
            Exit Function
        End If
        Call chg_Fdate()
        If Txt_nRECDATE.Text = "" Then
            do_sql.G_errmsg = "時間日期必須輸入"
            Exit Function
        End If
        If TXT_nREASON.Text = "" Then
            do_sql.G_errmsg = "事由必須輸入"
            Exit Function
        End If
        If DrDown_nSTARTTIME.Text >= DrDown_nENDTIME.Text Then
            do_sql.G_errmsg = "會客時間起須小於迄"
            Exit Function
        End If
        
        check_text = True
    End Function
    Function ins_p501(ByVal df As String) As Boolean
        ins_p501 = False
        
        If Txt_nName.Text = "" Then
            If df = "2" Then
                stmt = "select * from p_0501 "
                stmt += "where EFORMSN ='" + eformsn + "'" ' and nID='" + Txt_nID.Text + "'"
                If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
                    Exit Function
                End If
                If do_sql.G_table.Rows.Count > 0 Then
                    ins_p501 = True
                    Exit Function
                End If
            End If
            do_sql.G_errmsg = "姓名必須輸入"
            Exit Function
        End If
        If Txt_nService.Text = "" Then
            do_sql.G_errmsg = "服務單位必須輸入"
            Exit Function
        End If
        If Txt_nID.Text = "" Then
            do_sql.G_errmsg = "身份證字號必須輸入"
            Exit Function
        End If
        If RadioButton1.Checked = True Then '須檢查身份證
            If do_sql.checkIDNO(Txt_nID.Text) = False Then
                do_sql.G_errmsg = "身份證字號錯誤"
                Exit Function
            End If
        End If
        Dim str_sex As String
        If Radio1.Checked = True Then
            str_sex = "男"
        Else
            str_sex = "女"
        End If
        If DrDown_nkind.SelectedItem.Text = "" Then
            do_sql.G_errmsg = "類別必需輸入"
            Exit Function
        End If
		
		If txtCarNo.Text.IndexOf("-") >= 0 Then
            do_sql.G_errmsg = "車號不可輸入「-」，已自動去除"
            txtCarNo.Text = txtCarNo.Text.Replace("-", "")
        End If
		
        stmt = "select * from p_0501 "
        stmt += "where EFORMSN ='" + eformsn + "' and nID='" + Txt_nID.Text + "'"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Function
        End If
        If do_sql.G_table.Rows.Count = 0 Then
            stmt = "insert into P_0501(EFORMSN," '表單序號
            stmt += "nName,nSex,nService," '姓名,性別,服務單位,
            stmt += "nID,nKIND,nCarNo)" '身分證字號 "        
            stmt += " values('" + eformsn + "','" + Txt_nName.Text + "','" + str_sex + "','"
            stmt += Txt_nService.Text + "','" + Txt_nID.Text + "','" + DrDown_nkind.SelectedItem.Text + "'"
            stmt += ",'" + txtCarNo.Text + "')"

            If do_sql.db_exec(stmt, do_sql.G_conn_string) = False Then
                Exit Function
            End If
            ''新增Grid資料:
            'stmt = "select * from p_0501 "
            'stmt += "where EFORMSN ='" + eformsn + "'"
            'stmt += " order by ncreatedate"
            'If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            '    Exit Function
            'End If
            'GridView1.DataSource = do_sql.G_table

            GridView1.DataBind()
            Txt_nName.Text = ""
            Txt_nService.Text = ""
            Txt_nID.Text = ""
            txtCarNo.Text = ""
            Radio1.Checked = True
            Radio2.Checked = False
        End If
        ins_p501 = True
    End Function

    Protected Sub But_ins_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles But_ins.Click
        If ins_p501("1") = False Then
            Exit Sub
        End If
    End Sub

    Protected Sub RadioButton1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        Txt_nID.Text = ""
    End Sub

    Protected Sub RadioButton2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        Txt_nID.Text = ""
    End Sub
    Function select_data() As Boolean
        select_data = False
        Call drow_txt()
        stmt = "select * from p_05 "
        stmt += "where EFORMSN ='" + eformsn + "'"
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
            Label8.Text = ""

            DrDown_nRECROOM.SelectedItem.Text = ""
            DrDown_nRECEXIT.SelectedItem.Text = ""
            TXT_nPLACE.Text = ""
            Txt_nPHONE.Text = ""
            Txt_nRECDATE.Text = ""
            DrDown_nSTARTTIME.SelectedItem.Text = ""
            DrDown_nENDTIME.SelectedItem.Text = ""
            TXT_nREASON.Text = ""
        Else
            Lab_ORG_NAME_1.Text = do_sql.G_table.Rows(0).Item("PWUNIT").ToString
            Lab_emp_chinese_name.Text = do_sql.G_table.Rows(0).Item("PWNAME").ToString
            Lab_title_name_1.Text = do_sql.G_table.Rows(0).Item("PWTITLE").ToString
            Lab_ORG_NAME_2.Text = do_sql.G_table.Rows(0).Item("PAUNIT").ToString
            DrDown_emp_chinese_name.Items.Clear()
            DrDown_emp_chinese_name.Items.Add(Trim(do_sql.G_table.Rows(0).Item("PANAME").ToString))
            DrDown_emp_chinese_name.Items(0).Value = Trim(do_sql.G_table.Rows(0).Item("PANAME").ToString)

            Lab_title_name_2.Text = do_sql.G_table.Rows(0).Item("PATITLE").ToString
            Label8.Text = do_sql.G_table.Rows(0).Item("nAPPTIME").ToString

            DrDown_nRECROOM.SelectedItem.Text = do_sql.G_table.Rows(0).Item("nRECROOM").ToString
            DrDown_nRECEXIT.SelectedItem.Text = do_sql.G_table.Rows(0).Item("nRECEXIT").ToString
            TXT_nPLACE.Text = do_sql.G_table.Rows(0).Item("nPLACE").ToString
            Txt_nPHONE.Text = do_sql.G_table.Rows(0).Item("nPHONE").ToString
            Txt_nRECDATE.Text = CDate(do_sql.G_table.Rows(0).Item("nRECDATE").ToString).ToString("yyyy/MM/dd")
            DrDown_nSTARTTIME.SelectedItem.Text = do_sql.G_table.Rows(0).Item("nSTARTTIME").ToString
            DrDown_nENDTIME.SelectedItem.Text = do_sql.G_table.Rows(0).Item("nENDTIME").ToString
            DrDown_nSMin.SelectedItem.Text = do_sql.G_table.Rows(0).Item("nSMin").ToString
            DrDown_nEMin.SelectedItem.Text = do_sql.G_table.Rows(0).Item("nEMin").ToString
            TXT_nREASON.Text = do_sql.G_table.Rows(0).Item("nREASON").ToString
        End If
        'stmt = "select * from p_0501 "
        'stmt += "where EFORMSN ='" + eformsn + "'"
        ''stmt += " order by ncreatedate"
        'If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
        '    Exit Function
        'End If
        'GridView1.DataSource = do_sql.G_table

        'GridView1.DataBind()
        Call disabel_field()
        select_data = True
    End Function
    Sub disabel_field()
        DrDown_emp_chinese_name.Enabled = False
        DrDown_nRECROOM.Enabled = False
        DrDown_nRECEXIT.Enabled = False
        TXT_nPLACE.Enabled = False
        Txt_nPHONE.Enabled = False
        Txt_nRECDATE.Enabled = False
        DrDown_nSTARTTIME.Enabled = False
        DrDown_nENDTIME.Enabled = False
        DrDown_nSMin.Enabled = False
        DrDown_nEMin.Enabled = False
        TXT_nREASON.Enabled = False
        'GridView1.Enabled = False
        If read_only = "1" Then
            But_exe.Enabled = False
        End If
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
        But_ins.Enabled = False
    End Sub
    Sub drow_txt()
        '會客室名稱
        stmt = "select State_Name from SYSKIND where Kind_Num =3 and State_enabled=1"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Sub
        End If
        n_table = do_sql.G_table
        DrDown_nRECROOM.Items.Clear()
        p = 0
        DrDown_nRECROOM.Items.Add("--請選擇--")
        DrDown_nRECROOM.Items(p).Value = ""
        p += 1
        For Each dr In n_table.Rows
            DrDown_nRECROOM.Items.Add(Trim(dr("State_Name").ToString))
            DrDown_nRECROOM.Items(p).Value = Trim(dr("State_Name").ToString)
            p += 1
        Next

        '會客入口名稱
        '入口暫時不使用,放空白入資料庫 2013/10/24
        DrDown_nRECEXIT.Items.Clear()
        DrDown_nRECEXIT.Items.Add(New ListItem("", ""))
        'stmt = "select State_Name from SYSKIND where Kind_Num =4 and State_enabled=1"
        'If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
        '    Exit Sub
        'End If
        'n_table = do_sql.G_table
        'DrDown_nRECEXIT.Items.Clear()
        'p = 0
        'DrDown_nRECEXIT.Items.Add("--請選擇--")
        'DrDown_nRECEXIT.Items(p).Value = ""
        'p += 1
        'For Each dr In n_table.Rows
        '    DrDown_nRECEXIT.Items.Add(Trim(dr("State_Name").ToString))
        '    DrDown_nRECEXIT.Items(p).Value = Trim(dr("State_Name").ToString)
        '    p += 1
        'Next

        DrDown_nSTARTTIME.Items.Clear()
        For p = 8 To 16
            DrDown_nSTARTTIME.Items.Add(p.ToString("00"))
            DrDown_nSTARTTIME.Items(p - 8).Value = p.ToString("00")
        Next
        DrDown_nENDTIME.Items.Clear()
        For p = 9 To 17
            DrDown_nENDTIME.Items.Add(p.ToString("00"))
            DrDown_nENDTIME.Items(p - 9).Value = p.ToString("00")
        Next
    End Sub

    Sub chg_Fdate()
        If Txt_nRECDATE.Text = "" Then
            Exit Sub
        End If
        If Txt_nRECDATE.Text.Length = 10 Then
            Exit Sub
        End If
        Txt_nRECDATE.Text = CDate(Txt_nRECDATE.Text).ToString("yyyy/MM/dd")
    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim stmt As String
        Dim F_file As String
        Dim F_file2 As String
        Dim F_file_name As String
        Dim tool As New C_Public
        Dim dtFormDataPrint As DataTable
        Dim dtDetailDataPrint As DataTable        

        stmt = "select * from p_05 "
        stmt += "where EFORMSN ='" + Lab_eformsn.Text + "'"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Sub
        End If
        If do_sql.G_table.Rows.Count > 0 Then
            ''n_table = do_sql.G_table
            dtFormDataPrint = do_sql.G_table
        Else
            do_sql.G_errmsg = "查無可印資料"
            Exit Sub
        End If
        stmt = "select * from p_0501 "
        stmt += "where EFORMSN ='" + Lab_eformsn.Text + "'"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Sub
        End If
        ''n_table2 = do_sql.G_table
        dtDetailDataPrint = do_sql.G_table

        'Dim TY_date As Date = CDate(n_table.Rows(0).Item("nAPPTIME").ToString)
        'Dim T_date As String = CStr(CInt(TY_date.ToString("yyyy")) - 1911) + "年"
        'Dim BN_date As Date = CDate(n_table.Rows(0).Item("nRECDATE").ToString)
        'Dim BN_date As Date = CDate(dtFormDataPrint.Rows(0).Item("nRECDATE").ToString)


        'T_date += TY_date.ToString("MM") + "月" + TY_date.ToString("dd") + "日"
        F_file_name = "rpt050010"
        F_file = MapPath("../../rpt/" + F_file_name + ".txt")
        If File.Exists(F_file) = False Then
            do_sql.G_errmsg = "檔案不存在"
            Exit Sub
        End If

        'print_file = "../../drs/" + F_file_name + "-" + Now.ToString("mmssff") & ".drs"
        'F_file2 = MapPath(print_file)
        'Call do_sql.inc_file(F_file, F_file2, F_file_name)
        'Call do_sql.print_sdata(F_file2, "/init " + F_file_name + ".txt 1", "/newpage null")

        'Call do_sql.print_block(F_file2, "事由", "0", "0", n_table.Rows(0).Item("nREASON").ToString)
        'Call do_sql.print_block(F_file2, "單位", "0", "0", n_table.Rows(0).Item("PAUNIT").ToString)
        'Call do_sql.print_block(F_file2, "級職", "0", "0", n_table.Rows(0).Item("PATITLE").ToString)
        'Call do_sql.print_block(F_file2, "姓名", "0", "0", n_table.Rows(0).Item("PANAME").ToString)
        'Call do_sql.print_block(F_file2, "接送人_單位", "0", "0", n_table.Rows(0).Item("PAUNIT").ToString)
        'Call do_sql.print_block(F_file2, "接送人_級職", "0", "0", n_table.Rows(0).Item("PATITLE").ToString)
        'Call do_sql.print_block(F_file2, "接送人_姓名", "0", "0", n_table.Rows(0).Item("PANAME").ToString)
        'Call do_sql.print_block(F_file2, "接送人_電話", "0", "0", n_table.Rows(0).Item("nPHONE").ToString)
        'Call do_sql.print_block(F_file2, "會客地點", "0", "0", n_table.Rows(0).Item("nPLACE").ToString)
        'Call do_sql.print_block(F_file2, "年", "0", "0", do_sql.add_zero(CStr(CInt(BN_date.ToString("yyyy")) - 1911), 3))
        'Call do_sql.print_block(F_file2, "月", "0", "0", BN_date.ToString("MM"))
        'Call do_sql.print_block(F_file2, "日", "0", "0", BN_date.ToString("dd"))
        'Call do_sql.print_block(F_file2, "起時", "0", "0", n_table.Rows(0).Item("nSTARTTIME").ToString)
        'Call do_sql.print_block(F_file2, "迄時", "0", "0", n_table.Rows(0).Item("nENDTIME").ToString)
        'Dim y As Integer = 0
        'For pi As Integer = 0 To n_table2.Rows.Count - 1
        '    Call do_sql.print_block(F_file2, "來賓_姓名", "0", y.ToString(), n_table2.Rows(pi).Item("nName").ToString)
        '    Call do_sql.print_block(F_file2, "來賓_性別", "0", y.ToString(), n_table2.Rows(pi).Item("nSex").ToString)
        '    Call do_sql.print_block(F_file2, "來賓_身份證", "0", y.ToString(), tool.IDMask(n_table2.Rows(pi).Item("nID").ToString, 3, 2))
        '    Call do_sql.print_block(F_file2, "來賓_服務單位", "0", y.ToString(), n_table2.Rows(pi).Item("nService").ToString)
        '    Call do_sql.print_block(F_file2, "來賓_會客證", "0", y.ToString(), n_table2.Rows(pi).Item("nPassID").ToString)
        '    Call do_sql.print_block(F_file2, "來賓_車號", 0, y, n_table2.Rows(pi).Item("nCarNo").ToString)
        '    y += 8
        'Next

        'Call do_sql.print_block(F_file2, "底年", "0", "0", do_sql.add_zero(CStr(CInt(BN_date.ToString("yyyy")) - 1911), 3))
        'Call do_sql.print_block(F_file2, "底月", "0", "0", BN_date.ToString("MM"))
        'Call do_sql.print_block(F_file2, "底日", "0", "0", BN_date.ToString("dd"))
        Dim Print As New C_Xprint

        Dim filename As String = F_file_name + "-" + Now.ToString("mmssff") & ".drs"
        ''print_file = "../../drs/" + F_file_name + "-" + Now.ToString("mmssff") & ".drs"
        F_file2 = MapPath("../../drs/" + filename)

        Print.C_Xprint(F_file_name + ".txt", filename)


        'Call do_sql.inc_file(F_file, F_file2, F_file_name)
        Dim iPageCount As Integer = 0
        Dim iDivider As Integer = 8
        Dim iRemainder As Integer = 0
        Dim iTotalPages As Integer = 0
        Dim i_line As Integer = 0
        Dim iDetailNumber As Integer = dtDetailDataPrint.Rows.Count

        iTotalPages = Math.DivRem(iDetailNumber, iDivider, iRemainder)
        If iRemainder > 0 Then iTotalPages += 1
        Dim yShift = 3
        If dtFormDataPrint.Rows.Count > 0 Then
            iPageCount = 0
            For i = 0 To iTotalPages - 1
                Print.NewPage()
                ''換頁
                'Call do_sql.print_sdata(F_file2, "/init " + F_file_name + ".txt 1", "/newpage null")
                ''主資料
                Dim BN_date As Date = CDate(dtFormDataPrint.Rows(0).Item("nRECDATE").ToString)

                'Call do_sql.print_block(F_file2, "事由", "0", "0", dtFormDataPrint.Rows(0).Item("nREASON").ToString)
                'Call do_sql.print_block(F_file2, "單位", "0", "0", dtFormDataPrint.Rows(0).Item("PAUNIT").ToString)
                'Call do_sql.print_block(F_file2, "級職", "0", "0", dtFormDataPrint.Rows(0).Item("PATITLE").ToString)
                'Call do_sql.print_block(F_file2, "姓名", "0", "0", dtFormDataPrint.Rows(0).Item("PANAME").ToString)
                'Call do_sql.print_block(F_file2, "接送人_單位", "0", "0", dtFormDataPrint.Rows(0).Item("PAUNIT").ToString)
                'Call do_sql.print_block(F_file2, "接送人_級職", "0", "0", dtFormDataPrint.Rows(0).Item("PATITLE").ToString)
                'Call do_sql.print_block(F_file2, "接送人_姓名", "0", "0", dtFormDataPrint.Rows(0).Item("PANAME").ToString)
                'Call do_sql.print_block(F_file2, "接送人_電話", "0", "0", dtFormDataPrint.Rows(0).Item("nPHONE").ToString)
                'Call do_sql.print_block(F_file2, "會客地點", "0", "0", dtFormDataPrint.Rows(0).Item("nPLACE").ToString)
                'Call do_sql.print_block(F_file2, "年", "0", "0", do_sql.add_zero(CStr(CInt(BN_date.ToString("yyyy")) - 1911), 3))
                'Call do_sql.print_block(F_file2, "月", "0", "0", BN_date.ToString("MM"))
                'Call do_sql.print_block(F_file2, "日", "0", "0", BN_date.ToString("dd"))
                'Call do_sql.print_block(F_file2, "起時", "0", "0", dtFormDataPrint.Rows(0).Item("nSTARTTIME").ToString)
                'Call do_sql.print_block(F_file2, "迄時", "0", "0", dtFormDataPrint.Rows(0).Item("nENDTIME").ToString)

                Print.Add("事由", dtFormDataPrint.Rows(0).Item("nREASON").ToString, "0", "0")
                Print.Add("單位", dtFormDataPrint.Rows(0).Item("PAUNIT").ToString, "0", "0")
                Print.Add("級職", dtFormDataPrint.Rows(0).Item("PATITLE").ToString, "0", "0")
                Print.Add("姓名", dtFormDataPrint.Rows(0).Item("PANAME").ToString, "0", "0")
                Print.Add("接送人_單位", dtFormDataPrint.Rows(0).Item("PAUNIT").ToString, "0", "0")
                Print.Add("接送人_級職", dtFormDataPrint.Rows(0).Item("PATITLE").ToString, "0", "0")
                Print.Add("接送人_姓名", dtFormDataPrint.Rows(0).Item("PANAME").ToString, "0", "0")
                Print.Add("接送人_電話", dtFormDataPrint.Rows(0).Item("nPHONE").ToString, "0", "0")
                Print.Add("會客地點", dtFormDataPrint.Rows(0).Item("nPLACE").ToString, "0", "0")
                Print.Add("年", do_sql.add_zero(CStr(CInt(BN_date.ToString("yyyy")) - 1911), 3), "0", "0")
                Print.Add("月", BN_date.ToString("MM"), "0", "0")
                Print.Add("日", BN_date.ToString("dd"), "0", "0")
                Print.Add("起時", dtFormDataPrint.Rows(0).Item("nSTARTTIME").ToString, "0", "0")
                Print.Add("迄時", dtFormDataPrint.Rows(0).Item("nENDTIME").ToString, "0", "0")

                ''人員資料
                Dim y As Integer = 0
                ''每頁8筆滿頁狀態
                If (iPageCount < iTotalPages - 1) Or (iPageCount = 0 And iTotalPages = 1 And dtDetailDataPrint.Rows.Count = iDivider) Then
                    For j As Integer = 0 To iDivider - 1
                        'Call do_sql.print_block(F_file2, "來賓_姓名", "0", y.ToString(), dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nName").ToString)
                        'Call do_sql.print_block(F_file2, "來賓_性別", "0", y.ToString(), dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nSex").ToString)
                        'Call do_sql.print_block(F_file2, "來賓_身份證", "0", y.ToString(), tool.IDMask(dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nID").ToString, 3, 2))
                        'Call do_sql.print_block(F_file2, "來賓_服務單位", "0", y.ToString(), dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nService").ToString)
                        'Call do_sql.print_block(F_file2, "來賓_會客證", "0", y.ToString(), dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nPassID").ToString)
                        'Call do_sql.print_block(F_file2, "來賓_車號", "0", y.ToString(), dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nCarNo").ToString)

                        Print.Add("來賓_姓名", dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nName").ToString, "0", y.ToString())
                        Print.Add("來賓_性別", dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nSex").ToString, "0", y.ToString())
                        Print.Add("來賓_身份證", tool.IDMask(dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nID").ToString, 3, 2), "0", y.ToString())
                        Print.Add("來賓_服務單位", dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nService").ToString, "0", y.ToString())
                        Print.Add("來賓_會客證", dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nPassID").ToString, "0", y.ToString())
                        Print.Add("來賓_車號", dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nCarNo").ToString, "0", y.ToString())

                        y += 8
                    Next
                    ''每頁8筆未滿頁狀態
                ElseIf iPageCount = iTotalPages - 1 Then

                    For j As Integer = 0 To iRemainder - 1
                        'Call do_sql.print_block(F_file2, "來賓_姓名", "0", y.ToString(), dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nName").ToString)
                        'Call do_sql.print_block(F_file2, "來賓_性別", "0", y.ToString(), dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nSex").ToString)
                        'Call do_sql.print_block(F_file2, "來賓_身份證", "0", y.ToString(), tool.IDMask(dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nID").ToString, 3, 2))
                        'Call do_sql.print_block(F_file2, "來賓_服務單位", "0", y.ToString(), dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nService").ToString)
                        'Call do_sql.print_block(F_file2, "來賓_會客證", "0", y.ToString(), dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nPassID").ToString)
                        'Call do_sql.print_block(F_file2, "來賓_車號", "0", y.ToString(), dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nCarNo").ToString)

                        Print.Add("來賓_姓名", dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nName").ToString, "0", y.ToString())
                        Print.Add("來賓_性別", dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nSex").ToString, "0", y.ToString())
                        Print.Add("來賓_身份證", tool.IDMask(dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nID").ToString, 3, 2), "0", y.ToString())
                        Print.Add("來賓_服務單位", dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nService").ToString, "0", y.ToString())
                        Print.Add("來賓_會客證", dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nPassID").ToString, "0", y.ToString())
                        Print.Add("來賓_車號", dtDetailDataPrint.Rows(j + iPageCount * iDivider).Item("nCarNo").ToString, "0", y.ToString())

                        y += 8
                    Next

                End If
                iPageCount += 1
            Next
        End If

        Print.EndFile()
        If (Print.ErrMsg <> "") Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('" & Print.ErrMsg & "');")
            Response.Write("</script>")
        Else
            Response.Write("<script language='javascript'>")
            Response.Write("javascript:window.open('DownloadFile.aspx?file=" + filename + "');")
            Response.Write("</script>")
            
            '    Response.Write("<script language='javascript'>")
            '    Response.Write("window.onload = function() {")
            '    Response.Write("window.location.replace('../../drs/" & filename & "');")
            '    Response.Write("}")
            '    Response.Write("</script>")
        End If
    End Sub
    
    Function select_data2() As Boolean
        select_data2 = False
        'Call drow_txt()
        stmt = "select * from p_05 "
        'stmt += "where nAPPTIME >= '" + Now.ToString("yyyy/MM/dd") + "'"
        'stmt += "  and nAPPTIME < '" + Now.AddDays(3).ToString("yyyy/MM/dd") + "'"
        stmt += "where EFORMSN='" + GridView2.Rows(GridView2.SelectedRow.RowIndex).Cells(3).Text + "'"

        'stmt += " order by nAPPTIME desc"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Function
        End If
        If do_sql.G_table.Rows.Count = 0 Then
            Exit Function
            '下面不做
            Lab_ORG_NAME_1.Text = ""
            Lab_emp_chinese_name.Text = ""
            Lab_title_name_1.Text = ""
            Lab_ORG_NAME_2.Text = ""
            DrDown_emp_chinese_name.Items.Clear()
            DrDown_emp_chinese_name.Items.Add("")
            DrDown_emp_chinese_name.Items(0).Value = ""

            Lab_title_name_2.Text = ""
            Label8.Text = ""

            DrDown_nRECROOM.SelectedItem.Text = ""
            DrDown_nRECEXIT.SelectedItem.Text = ""
            TXT_nPLACE.Text = ""
            Txt_nPHONE.Text = ""
            Txt_nRECDATE.Text = ""
            DrDown_nSTARTTIME.SelectedItem.Text = ""
            DrDown_nENDTIME.SelectedItem.Text = ""
            TXT_nREASON.Text = ""
        Else
            ''填表單資料
            do_sql.set_drop_down(DrDown_nRECROOM, do_sql.G_table.Rows(0).Item("nRECROOM").ToString.Trim)
            do_sql.set_drop_down(DrDown_nRECEXIT, do_sql.G_table.Rows(0).Item("nRECEXIT").ToString.Trim)
            TXT_nPLACE.Text = do_sql.G_table.Rows(0).Item("nPLACE").ToString
            Txt_nPHONE.Text = do_sql.G_table.Rows(0).Item("nPHONE").ToString
            Txt_nRECDATE.Text = "" 'CDate(do_sql.G_table.Rows(0).Item("nRECDATE").ToString).ToString("yyyy/MM/dd")
            do_sql.set_drop_down(DrDown_nSTARTTIME, do_sql.G_table.Rows(0).Item("nSTARTTIME").ToString.Trim)
            do_sql.set_drop_down(DrDown_nENDTIME, do_sql.G_table.Rows(0).Item("nENDTIME").ToString.Trim)
            do_sql.set_drop_down(DrDown_nSMin, do_sql.G_table.Rows(0).Item("nSMin").ToString.Trim)
            do_sql.set_drop_down(DrDown_nEMin, do_sql.G_table.Rows(0).Item("nEMin").ToString.Trim)

            TXT_nREASON.Text = do_sql.G_table.Rows(0).Item("nREASON").ToString
            n_table = do_sql.G_table
        End If
        
        If do_sql.begin_tran(do_sql.G_conn_string) = False Then
            Exit Function
        End If
        'stmt = "delete from p_05 "
        'stmt += "where EFORMSN ='" + eformsn + "'"
        'If do_sql.db_exec(stmt, do_sql.G_Trans) = False Then
        '    Call do_sql.rollback_tran()
        '    Exit Function
        'End If
        stmt = "delete from p_0501 where EFORMSN='" + eformsn + "'"
        If do_sql.db_exec(stmt, do_sql.G_Trans) = False Then
            Call do_sql.rollback_tran()
            Exit Function
        End If
        stmt = "insert into p_0501(EFORMSN,nName,nSex,nKind,nService,nID,nCarNo) "
        stmt += "select '" + eformsn + "',nName,nSex,nKind,nService,nID,nCarNo " '
        stmt += " from p_0501 where EFORMSN='" + n_table.Rows(0).Item("EFORMSN").ToString + "'"
        If do_sql.db_exec(stmt, do_sql.G_Trans) = False Then
            Call do_sql.rollback_tran()
            Exit Function
        End If
        'Call do_sql.rollback_tran()
        Call do_sql.commit_tran()

        'GridView1.DataBind()
        'Call disabel_field()
        GridView1.DataBind()

        select_data2 = True
    End Function

    Protected Sub GridView2_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowCreated
        If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then
            '隱藏eformsn
            e.Row.Cells(3).Visible = False
        End If
        
    End Sub

    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        If e.Row.RowIndex >= 0 Then
            'e.Row.Cells(3).Text = "111,222,333"
            stmt = "select nName from p_0501 where eformsn='" + e.Row.Cells(3).Text + "'"
            If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
                Exit Sub
            End If
            If do_sql.G_table.Rows.Count > 0 Then
                n_table = do_sql.G_table
                Dim xname As String = ""
                For Each dr In n_table.Rows
                    If xname = "" Then
                        xname = Trim(dr("nName").ToString)
                    Else
                        xname += "," + Trim(dr("nName").ToString)
                    End If
                Next
                e.Row.Cells(1).Text = xname
            End If
        End If
    End Sub

    Protected Sub GridView2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView2.SelectedIndexChanged
        'Dim ppp As String
        ''ppp = 1
        'ppp = GridView2.Rows(GridView2.SelectedRow.RowIndex).Cells(0).Text

        If select_data2() = False Then
            Exit Sub
        End If

        Div_grid.Visible = False



    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '找出同級單位以下全部單位
        Dim Org_Down As New C_Public

        '單純讀取表單不可送件
        If read_only = "1" Then
            But_exe.Visible = False
            backBtn.Visible = False
            tranBtn.Visible = False
            SelOldEform.Visible = False

            '讀取表單可送件
        ElseIf read_only = "2" Then
            SelOldEform.Visible = False

            Dim PerCount As Integer = 0

            '判斷批核者是否為副主管,假如批核者為副主管則不可送件只可呈轉
            Dim ParentFlag As String = ""
            Dim ParentVal = user_id & "," & eformsn

            Dim FC As New C_FlowSend.C_FlowSend
            ParentFlag = FC.F_NextChief(ParentVal, connstr)

            If ParentFlag = "1" Then
                But_exe.Visible = False

                '上一級主管多少人
                db.Open()
                Dim PerCountCom As New SqlCommand("select count(empuid) as PerCount from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = (select ORG_UID from EMPLOYEE where employee_id ='" & user_id & "'))", db)
                Dim PerRdv = PerCountCom.ExecuteReader()
                If PerRdv.Read() Then
                    PerCount = CInt(PerRdv("PerCount"))
                End If
                db.Close()

                '上一級沒人則呈現送件按鈕
                If PerCount = 0 Then
                    But_exe.Visible = True
                    tranBtn.Visible = False
                Else

                    '登入者為二級單位主管則可核准
                    Dim strParentOrg As String = ""
                    strParentOrg = Org_Down.getUporg(org_uid, 2)

                    If UCase(strParentOrg) = UCase(org_uid) Then
                        But_exe.Visible = True
                    Else

                        Dim strOrgUid As String = ""

                        '判斷被代理人單位
                        db.Open()
                        Dim UpPerCom As New SqlCommand("select ORG_UID from EMPLOYEE where employee_id ='" & AgentEmpuid & "'", db)
                        Dim UpPerRdv = UpPerCom.ExecuteReader()
                        If UpPerRdv.read() Then
                            strOrgUid = UpPerRdv("ORG_UID")
                        End If
                        db.Close()

                        '被代理人為二級單位主管則可以送件
                        If UCase(strParentOrg) = UCase(strOrgUid) Then
                            But_exe.Visible = True
                        Else
                            But_exe.Visible = False
                        End If

                    End If

                End If
	    Else

                '上一級主管多少人
                db.Open()
                Dim PerCountCom As New SqlCommand("select count(empuid) as PerCount from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = (select ORG_UID from EMPLOYEE where employee_id ='" & user_id & "'))", db)
                Dim PerRdv = PerCountCom.ExecuteReader()
                If PerRdv.Read() Then
                    PerCount = CInt(PerRdv("PerCount"))
                End If
                db.Close()

                '上一級沒人則呈現送件按鈕
                If PerCount = 0 Then
                    But_exe.Visible = True
                    tranBtn.Visible = False
                End If
            End If
            
	    '沒有批核權限不可執行動作
            If (Org_Down.ApproveAuth(eformsn, user_id)) = "" Then
                But_exe.Visible = False
                backBtn.Visible = False
                tranBtn.Visible = False
            End If

        Else
            backBtn.Visible = False
            tranBtn.Visible = False
        End If

    End Sub

    Protected Sub backBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles backBtn.Click

        Try

            If read_only = "2" Then
                '增加批核意見
                insComment(txtcomment.Text, eformsn, user_id)
            End If

            Dim Val_P As String
            Val_P = ""

            '表單駁回
            Dim FC As New C_FlowSend.C_FlowSend

            Dim SendVal As String = ""

            '判斷是否為代理人批核的表單
            If AgentEmpuid = "" Then
                SendVal = eformid & "," & user_id & "," & eformsn & "," & "1"
            Else
                SendVal = eformid & "," & AgentEmpuid & "," & eformsn & "," & "1"
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
            Dim tool As New C_Public
            If tool.GetSuperiors(user_id).Count > 1 Then '若上一級主管為兩人以上,則切換至選擇欲呈轉主管之頁面
                Server.Transfer("../00/MOA00014.aspx?eformsn=" + eformsn)
            ElseIf tool.GetSuperiors(user_id).Count = 0 Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('無上一級主管');")
                Response.Write(" </script>")
                Exit Sub
            Else
                If read_only = "2" Then
                    '增加批核意見
                    insComment(txtcomment.Text, eformsn, user_id)
                End If

                Dim Val_P As String = ""

                '表單呈轉
                'Dim FC As New C_FlowSend.C_FlowSend
				Dim FC As New CFlowSend

                Dim SendVal As String = ""

                '判斷是否為代理人批核的表單
                If AgentEmpuid = "" Then
                    SendVal = eformid & "," & user_id & "," & eformsn & "," & "1"
                Else
                    SendVal = eformid & "," & AgentEmpuid & "," & eformsn & "," & "1"
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

    Protected Sub SelOldEform_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SelOldEform.Click

        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        Dim Visitors As Integer
        Dim str_SQL As String = ""

        '最近是否有申請過會客表單
        db.Open()
        Dim PerCountCom As New SqlCommand("SELECT count(P_NUM) as Visitors from P_05 WHERE (DATEDIFF(day, nAPPTIME, GETDATE()) < 90) AND PAIDNO = '" & user_id & "' AND p_05.eformsn IN (select p_0501.eformsn from p_0501 where P_0501.nKIND IS NOT NULL)", db)
        Dim PerRdv = PerCountCom.ExecuteReader()
        If PerRdv.read() Then
            Visitors = PerRdv("Visitors")
        End If
        db.Close()

        If Visitors > 0 Then
            Div_grid.Visible = True
            Div_grid.Style("Top") = "280px"
            Div_grid.Style("left") = "380px"
        Else

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('最近沒有申請會客洽公申請單資料');")
            Response.Write(" </script>")

        End If


    End Sub

    Protected Sub CloseVistor_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CloseVistor.Click

        '隱藏表單快速新增
        Div_grid.Visible = False

    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated

        If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then

            '新表單才可刪除
            If read_only <> "" Then

                '隱藏刪除按鈕
                e.Row.Cells(8).Visible = False

            End If

        End If

    End Sub

    Protected Sub DrDown_nRECROOM_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrDown_nRECROOM.SelectedIndexChanged


        'If DrDown_nRECROOM.SelectedItem.Text = "第一會客室" Then

        '    '會客入口名稱
        '    stmt = "select State_Name from SYSKIND where Kind_Num =4 and (State_Num='1' or State_Num='2') and State_enabled=1"
        '    If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
        '        Exit Sub
        '    End If
        '    n_table = do_sql.G_table
        '    DrDown_nRECEXIT.Items.Clear()
        '    p = 0
        '    DrDown_nRECEXIT.Items.Add("--請選擇--")
        '    DrDown_nRECEXIT.Items(p).Value = ""
        '    p += 1
        '    For Each dr In n_table.Rows
        '        DrDown_nRECEXIT.Items.Add(Trim(dr("State_Name").ToString))
        '        DrDown_nRECEXIT.Items(p).Value = Trim(dr("State_Name").ToString)
        '        p += 1
        '    Next

        'ElseIf DrDown_nRECROOM.SelectedItem.Text = "第二會客室" Then

        '    '會客入口名稱
        '    stmt = "select State_Name from SYSKIND where Kind_Num =4 and (State_Num='3' or State_Num='4') and State_enabled=1"
        '    If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
        '        Exit Sub
        '    End If
        '    n_table = do_sql.G_table
        '    DrDown_nRECEXIT.Items.Clear()
        '    p = 0
        '    DrDown_nRECEXIT.Items.Add("--請選擇--")
        '    DrDown_nRECEXIT.Items(p).Value = ""
        '    p += 1
        '    For Each dr In n_table.Rows
        '        DrDown_nRECEXIT.Items.Add(Trim(dr("State_Name").ToString))
        '        DrDown_nRECEXIT.Items(p).Value = Trim(dr("State_Name").ToString)
        '        p += 1
        '    Next
        'Else

        '    DrDown_nRECEXIT.Items.Clear()
        '    '不使用會客入口 2013/10/24
        '    p = 0
        '    DrDown_nRECEXIT.Items.Add("--請選擇--")
        '    DrDown_nRECEXIT.Items(p).Value = ""

        'End If


    End Sub

    Protected Sub GridView10_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView10.SelectedIndexChanged
        txtcomment.Text = GridView10.Rows(GridView10.SelectedRow.RowIndex).Cells(0).Text
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
            Div_grid10.Style("Top") = "500px"
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

    Protected Sub GridView1_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim tool As New C_Public
        
        Select Case (e.Row.RowType)
            Case DataControlRowType.DataRow, DataControlRowType.Header
                ''Dim datarow As DataRowView = CType(e.Row.DataItem, DataRowView)
                If read_only = "1" Then
                    e.Row.Cells(6).Text = tool.IDMask(e.Row.Cells(6).Text, 3, 2)                    
                End If

        End Select
    End Sub
End Class
