Imports System.Data
Imports System.Data.SqlClient

Partial Class Source_06_MOA06001
    Inherits System.Web.UI.Page
    Public do_sql As New C_SQLFUN
    Dim dr As System.Data.DataRow
    Dim n_table As New System.Data.DataTable
    Public eformsn As String
    Public d_pub As New C_Public
    Dim eformid As String
    Dim str_ORG_UID As String
    Public stmt As String
    Dim p As Integer = 0
    Dim K As Integer = 0
    Dim str_app_id As String
    Public read_only As String = ""
    Dim connstr, user_id, org_uid As String
    Dim sql As String
    Dim D2 As New System.Data.DataTable
    Dim D1 As New System.Data.DataTable
    Dim AgentEmpuid As String = ""
    Dim AppUID As String

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

        do_sql.G_errmsg = ""
        do_sql.G_user_id = Session("user_id") '"tempu180"
        'user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        eformsn = Request("eformsn") '"VJ" 'd_pub.randstr(2) '
        eformid = Request("eformid")
        read_only = Request("read_only")

        'session被清空回首頁
        If user_id = "" And org_uid = "" And eformsn = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.parent.location='../../index.aspx';")
            Response.Write(" </script>")

        Else

            '找出同級單位以下全部單位
            Dim Org_Down As New C_Public

            'do_sql.G_user_id = "tempu178"
            'eformsn = "R5Y26NOFA5123456" 'd_pub.randstr(16) '
            'eformid = "R5Y26NOFA5"
            'read_only = "1"
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
                Lab_PWUNIT.Text = ""
                Exit Sub
            End If
            If do_sql.G_usr_table.Rows.Count > 0 Then
                Lab_PWUNIT.Text = do_sql.G_usr_table.Rows(0).Item("ORG_NAME").ToString.Trim
                Lab_PWNAME.Text = do_sql.G_usr_table.Rows(0).Item("emp_chinese_name").ToString.Trim
                Lab_PWTITLE.Text = do_sql.G_usr_table.Rows(0).Item("AD_title").ToString.Trim
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
                    DrDown_PANAME.Items.Clear()

                    For Each dr In n_table.Rows
                        DrDown_PANAME.Items.Add(Trim(dr("emp_chinese_name").ToString))
                        DrDown_PANAME.Items(p).Value = Trim(dr("employee_id").ToString)
                        If UCase(do_sql.G_user_id) = UCase(Trim(dr("employee_id").ToString)) Then
                            K = p
                        End If
                        p += 1
                    Next
                    If p > 0 Then
                        DrDown_PANAME.SelectedIndex = K
                        Call DrDown_PANAME_SelectedIndexChanged(sender, e)
                    End If
                End If
            End If

            Call Add_DropDownList()

            '新增Grid資料:
            'stmt = "select * from p_0601 "
            'stmt += "where EFORMSN ='" + eformsn + "'"
            ''stmt += " order by ncreatedate"
            'If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            '    Exit Sub
            'End If
            'GridView1.DataSource = do_sql.G_table

            GridView1.DataBind()



        End If

    End Sub

    Protected Sub DrDown_PANAME_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrDown_PANAME.SelectedIndexChanged
        str_app_id = ""
        If do_sql.select_urname(DrDown_PANAME.Items(DrDown_PANAME.SelectedIndex).Value) = False Then
            DrDown_PANAME.Text = ""
            Exit Sub
        End If

        If do_sql.G_usr_table.Rows.Count > 0 Then
            Dim stmt As String = ""
            Dim CHINE_NAME As String = ""
            p = 0
            Lab_PAUNIT.Text = do_sql.G_usr_table.Rows(0).Item("ORG_NAME").ToString.Trim
            Lab_PATITLE.Text = do_sql.G_usr_table.Rows(0).Item("AD_title").ToString.Trim
            str_ORG_UID = do_sql.G_usr_table.Rows(0).Item("ORG_UID").ToString.Trim
            str_app_id = do_sql.G_usr_table.Rows(0).Item("employee_id").ToString.Trim
        End If
    End Sub
    Sub Add_DropDownList()
        '請假事由
        DrDown_nREASON.Items.Clear()
        p = 0
        DrDown_nREASON.Items.Add("--請選擇--")
        DrDown_nREASON.Items(p).Value = ""
        p += 1
        DrDown_nREASON.Items.Add("演訓運用")
        DrDown_nREASON.Items(p).Value = "演訓運用"
        p += 1
        DrDown_nREASON.Items.Add("業務督考")
        DrDown_nREASON.Items(p).Value = "業務督考"
        p += 1
        DrDown_nREASON.Items.Add("上級調閱")
        DrDown_nREASON.Items(p).Value = "上級調閱"
        p += 1
        DrDown_nREASON.Items.Add("營外會報")
        DrDown_nREASON.Items(p).Value = "營外會報"
        p += 1
        DrDown_nREASON.Items.Add("任務協調")
        DrDown_nREASON.Items(p).Value = "任務協調"
        p += 1
        DrDown_nREASON.Items.Add("國會備詢")
        DrDown_nREASON.Items(p).Value = "國會備詢"
        p += 1
        DrDown_nREASON.Items.Add("司法約詢")
        DrDown_nREASON.Items(p).Value = "司法約詢"
        p += 1
        DrDown_nREASON.Items.Add("公務出國")
        DrDown_nREASON.Items(p).Value = "公務出國"
        p += 1
        DrDown_nREASON.Items.Add("保修維護")
        DrDown_nREASON.Items(p).Value = "保修維護"
        p += 1
        DrDown_nREASON.Items.Add("其他")
        DrDown_nREASON.Items(p).Value = "其他"

        '區分 
        p = 0
        DrDown_nKind.Items.Add("--請選擇--")
        DrDown_nKind.Items(p).Value = ""
        p += 1
        DrDown_nKind.Items.Add("電腦主機")
        DrDown_nKind.Items(p).Value = "電腦主機"
        p += 1
        DrDown_nKind.Items.Add("磁片光碟")
        DrDown_nKind.Items(p).Value = "磁片光碟"
        p += 1
        DrDown_nKind.Items.Add("行動硬碟")
        DrDown_nKind.Items(p).Value = "行動硬碟"
        p += 1
        DrDown_nKind.Items.Add("文書圖表")
        DrDown_nKind.Items(p).Value = "文書圖表"
        p += 1
        DrDown_nKind.Items.Add("保密裝備")
        DrDown_nKind.Items(p).Value = "保密裝備"
        p += 1
        DrDown_nKind.Items.Add("其他")
        DrDown_nKind.Items(p).Value = "其他"

        '機密等級 
        p = 0
        DrDown_nClass.Items.Add("--請選擇--")
        DrDown_nClass.Items(p).Value = ""
        p += 1
        DrDown_nClass.Items.Add("普通")
        DrDown_nClass.Items(p).Value = "普通"
        p += 1
        DrDown_nClass.Items.Add("密")
        DrDown_nClass.Items(p).Value = "密"
        p += 1
        DrDown_nClass.Items.Add("機密")
        DrDown_nClass.Items(p).Value = "機密"
        p += 1
        DrDown_nClass.Items.Add("極機密")
        DrDown_nClass.Items(p).Value = "極機密"
        p += 1
        DrDown_nClass.Items.Add("絕對機密")
        DrDown_nClass.Items(p).Value = "絕對機密"
    End Sub


    Protected Sub But_ins_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles But_ins.Click
        If ins_p601("1") = False Then
            Exit Sub
        End If
    End Sub
    Function ins_p601(ByVal df As String) As Boolean
        ins_p601 = False

        If DrDown_nKind.SelectedItem.Value = "" Then
            If df = "2" Then
                stmt = "select * from p_0601 "
                stmt += "where EFORMSN ='" + eformsn + "'"
                If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
                    Exit Function
                End If
                If do_sql.G_table.Rows.Count > 0 Then
                    ins_p601 = True
                    Exit Function
                End If
                'Exit Function
            End If
            do_sql.G_errmsg = "區分必須輸入"
            Exit Function
        End If
        If Txt_nMName.Text = "" Then
            do_sql.G_errmsg = "名稱機型必須輸入"
            Exit Function
        End If
        'If Txt_nDocnum.Text = "" Then
        '    do_sql.G_errmsg = "編(文)號必須輸入"
        '    Exit Function
        'End If
        If Txt_nAmount.Text = "" Then
            do_sql.G_errmsg = "數量必須輸入"
            Exit Function
        End If
        If Txt_nContent.Text = "" Then
            do_sql.G_errmsg = "編內容概要必須輸入"
            Exit Function
        End If
        If DrDown_nClass.SelectedItem.Value = "" Then
            do_sql.G_errmsg = "機密等級必須輸入"
            Exit Function
        End If

        'stmt = "select * from p_0601 "
        'stmt += "where EFORMSN ='" + eformsn + "'"
        'If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
        '    Exit Function
        'End If
        'If do_sql.G_table.Rows.Count = 0 Then
        stmt = "insert into P_0601(EFORMSN," '表單序號
        stmt += "nKind,nMName,nDocnum," '區分,名稱機型,編(文)號,
        stmt += "nAmount,nContent,nClass)" '數量,內容摘要,機密等級 "        
        stmt += " values('" + eformsn + "','" + DrDown_nKind.SelectedItem.Value + "','" + Txt_nMName.Text + "','"
        stmt += Txt_nDocnum.Text + "'," + Txt_nAmount.Text + ",'" + Txt_nContent.Text + "','" + DrDown_nClass.SelectedItem.Value + "')"

        If do_sql.db_exec(stmt, do_sql.G_conn_string) = False Then
            Exit Function
        End If
        '新增Grid資料:
        'stmt = "select * from p_0601 "
        'stmt += "where EFORMSN ='" + eformsn + "'"
        ''stmt += " order by ncreatedate"
        'If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
        '    Exit Function
        'End If
        'GridView1.DataSource = do_sql.G_table

        GridView1.DataBind()
        DrDown_nKind.SelectedIndex = 0
        Txt_nMName.Text = ""
        Txt_nDocnum.Text = ""
        Txt_nAmount.Text = ""
        Txt_nContent.Text = ""
        DrDown_nClass.SelectedIndex = 0
        'End If
        ins_p601 = True
    End Function

    Protected Sub But_exe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles But_exe.Click

        Dim FC As New C_FlowSend.C_FlowSend

        Dim SendVal As String = ""

        '判斷是否為代理人批核的表單
        If AgentEmpuid = "" Then
            'SendVal = eformid & "," & user_id & "," & eformsn & "," & "1" & ","
            SendVal = eformid & "," & DrDown_PANAME.SelectedValue & "," & eformsn & "," & "1" & ","
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
            If ins_p601("2") = False Then
                Exit Sub
            End If
            stmt = "insert into P_06(EFORMSN,PWUNIT,PWTITLE," '表單序號,填表人單位,填表人級職
            stmt += "PWNAME, PWIDNO, PAUNIT," '填表人姓名,填表人身份證字號,申請人單位
            stmt += "PANAME, PATITLE, PAIDNO," '申請人姓名,申請人級職,申請人身份證字號 "
            stmt += "nAPPTIME, nREASON, " '申請時間,申請事由,
            stmt += "nDATE,nPLACE) " '申請出入日期,使用地點        
            stmt += " values('" + eformsn + "','" + Lab_PWUNIT.Text + "','" + Lab_PWTITLE.Text + "',N'"
            stmt += Lab_PWNAME.Text + "','" + do_sql.G_user_id + "','" + Lab_PAUNIT.Text + "',N'"
            stmt += DrDown_PANAME.SelectedItem.Text + "','" + Lab_PATITLE.Text + "','" + DrDown_PANAME.SelectedItem.Value + "','"
            stmt += Label8.Text + "','" + DrDown_nREASON.SelectedItem.Text + "','"
            stmt += Txt_nDATE.Text + "','" + TXT_nPLACE.Text + "')"
            If do_sql.db_exec(stmt, do_sql.G_conn_string) = False Then
                Exit Sub
            End If

            If DateDiff(DateInterval.Day, Now(), CDate(Txt_nDATE.Text)) < 0 Then

                '補登
                Val_P = FC.F_Supply(SendVal, do_sql.G_conn_string)

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('補登完成');")
                Response.Write(" window.parent.location='../00/MOA00020.aspx?x=MOA06001&y=D6Y95Y5XSU';")
                Response.Write(" </script>")
                Exit Sub

            End If

        End If

        If read_only = "2" Then

            Dim strAgentName As String = ""

            '判斷是否為代理人批核
            If UCase(user_id) = UCase(AgentEmpuid) Then

                '增加批核意見
                insComment(txtcomment.Text, eformsn, user_id)

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
                insComment(strComment, eformsn, AgentEmpuid)

            End If

        End If

        Val_P = ""

        '判斷下一關為上一級主管時人數是否超過一人
        If (NextPer > 1 And read_only = "") Or (NextPer > 1 And strgroup_id = "Z860") Then
            Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & eformsn & "&SelFlag=1&AppUID=" & DrDown_PANAME.SelectedValue)
        Else

            '正常
            Val_P = FC.F_Send(SendVal, do_sql.G_conn_string)
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
        If Lab_PAUNIT.Text = "" Then
            do_sql.G_errmsg = "申請人姓名必須選取"
            Exit Function
        End If
        If DrDown_nREASON.SelectedIndex < 1 Then
            do_sql.G_errmsg = "申請事由必須選取"
            Exit Function
        End If
        Call chg_Fdate()
        If Txt_nDATE.Text = "" Then
            do_sql.G_errmsg = "申請出入日期必須選取"
            Exit Function
        End If
        If TXT_nPLACE.Text = "" Then
            do_sql.G_errmsg = "地點:必須輸入"
            Exit Function
        End If
        check_text = True
    End Function
    Function select_data() As Boolean
        select_data = False
        Call Add_DropDownList()
        stmt = "select * from p_06 "
        stmt += "where EFORMSN ='" + eformsn + "'"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Function
        End If
        If do_sql.G_table.Rows.Count = 0 Then
            Lab_PWUNIT.Text = ""
            Lab_PWNAME.Text = ""
            Lab_PWTITLE.Text = ""
            Lab_PAUNIT.Text = ""
            DrDown_PANAME.Items.Clear()
            DrDown_PANAME.Items.Add("")
            DrDown_PANAME.Items(0).Value = ""

            Lab_PATITLE.Text = ""
            Label8.Text = ""

            DrDown_nREASON.SelectedItem.Text = ""
            Txt_nDATE.Text = ""
            Txt_nPLACE.Text = ""
        Else
            Lab_PWUNIT.Text = do_sql.G_table.Rows(0).Item("PWUNIT").ToString
            Lab_PWNAME.Text = do_sql.G_table.Rows(0).Item("PWNAME").ToString
            Lab_PWTITLE.Text = do_sql.G_table.Rows(0).Item("PWTITLE").ToString
            Lab_PAUNIT.Text = do_sql.G_table.Rows(0).Item("PAUNIT").ToString
            DrDown_PANAME.Items.Clear()
            DrDown_PANAME.Items.Add(Trim(do_sql.G_table.Rows(0).Item("PANAME").ToString))
            DrDown_PANAME.Items(0).Value = Trim(do_sql.G_table.Rows(0).Item("PANAME").ToString)

            Lab_PATITLE.Text = do_sql.G_table.Rows(0).Item("PATITLE").ToString
            Label8.Text = do_sql.G_table.Rows(0).Item("nAPPTIME").ToString

            DrDown_nREASON.SelectedItem.Text = do_sql.G_table.Rows(0).Item("nREASON").ToString
            Txt_nDATE.Text = CDate(do_sql.G_table.Rows(0).Item("nDATE").ToString).ToString("yyyy/MM/dd")
            TXT_nPLACE.Text = do_sql.G_table.Rows(0).Item("nPLACE").ToString
        End If
        'stmt = "select * from p_0601 "
        'stmt += "where EFORMSN ='" + eformsn + "'"
        ''stmt += " order by ncreatedate"
        'If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
        '    Exit Function
        'End If
        'GridView1.DataSource = do_sql.G_table

        GridView1.DataBind()
        Call disabel_field()
        select_data = True
    End Function
    Sub disabel_field()
        DrDown_PANAME.Enabled = False
        DrDown_nREASON.Enabled = False
        Txt_nPLACE.Enabled = False
        Txt_nDATE.Enabled = False
        GridView1.Enabled = False
        If read_only = "1" Then
            But_exe.Visible = False
            But_ins.Visible = False
        ElseIf read_only = "2" Then
            But_ins.Visible = False
        End If

    End Sub


    Sub chg_Fdate()
        If Txt_nDATE.Text = "" Then
            Exit Sub
        End If
        If Txt_nDATE.Text.Length = 10 Then
            Exit Sub
        End If
        Txt_nDATE.Text = CDate(Txt_nDATE.Text).ToString("yyyy/MM/dd")
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
            If GetSuperiors(user_id).Rows.Count > 1 Then '若上一級主管為兩人以上,則切換至選擇欲呈轉主管之頁面
                Server.Transfer("../00/MOA00014.aspx?eformsn=" + eformsn)
            Else
                If read_only = "2" Then
                    '增加批核意見
                    insComment(txtcomment.Text, eformsn, user_id)
                End If


                Dim Val_P As String
                Val_P = ""

                '表單呈轉
                Dim FC As New C_FlowSend.C_FlowSend

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

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '單純讀取表單不可送件
        If read_only = "1" Then
            But_exe.Visible = False
            backBtn.Visible = False
            tranBtn.Visible = False
            ImgPrint.Visible = False

            '讀取表單可送件
        ElseIf read_only = "2" Then

            Dim PerCount As Integer = 0

            '判斷批核者是否為副主管,假如批核者為副主管則不可送件只可呈轉
            Dim ParentFlag As String = ""
            Dim ParentVal = user_id & "," & eformsn

            Dim FC As New C_FlowSend.C_FlowSend
            ParentFlag = FC.F_NextChief(ParentVal, connstr)

            If ParentFlag = "1" Then
                'But_exe.Visible = False
                ImgPrint.Visible = False

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
                    But_exe.Visible = True
                    tranBtn.Visible = False
                Else
                    But_exe.Visible = True
                    'But_exe.Visible = False
                End If

            End If

            '找出同級單位以下全部單位
            Dim Org_Down As New C_Public

            '沒有批核權限不可執行動作
            If (Org_Down.ApproveAuth(eformsn, user_id)) = "" Then
                But_exe.Visible = False
                backBtn.Visible = False
                tranBtn.Visible = False
                ImgPrint.Visible = False
            End If

        Else
            backBtn.Visible = False
            tranBtn.Visible = False
        End If

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

    Protected Sub ImgPrint_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgPrint.Click

        '新增資料
        Dim connstr As String
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '新增資料
        db.Open()

        If check_text() = False Then
            Exit Sub
        End If
        If ins_p601("2") = False Then
            Exit Sub
        End If

        stmt = "insert into P_06(EFORMSN,PWUNIT,PWTITLE," '表單序號,填表人單位,填表人級職
        stmt += "PWNAME, PWIDNO, PAUNIT," '填表人姓名,填表人身份證字號,申請人單位
        stmt += "PANAME, PATITLE, PAIDNO," '申請人姓名,申請人級職,申請人身份證字號 "
        stmt += "nAPPTIME, nREASON, " '申請時間,申請事由,
        stmt += "nDATE,nPLACE,nApply) " '申請出入日期,使用地點,是否送件        
        stmt += " values('" + eformsn + "','" + Lab_PWUNIT.Text + "','" + Lab_PWTITLE.Text + "','"
        stmt += Lab_PWNAME.Text + "','" + do_sql.G_user_id + "','" + Lab_PAUNIT.Text + "','"
        stmt += DrDown_PANAME.SelectedItem.Text + "','" + Lab_PATITLE.Text + "','" + DrDown_PANAME.SelectedItem.Value + "','"
        stmt += Label8.Text + "','" + DrDown_nREASON.SelectedItem.Text + "','"
        stmt += Txt_nDATE.Text + "','" + TXT_nPLACE.Text + "',1)"

        Dim insCom As New SqlCommand(stmt, db)
        insCom.ExecuteNonQuery()
        db.Close()

        '列印
        Dim filename As String = "prt_06003" & Rnd() & ".drs"
        Dim print As New C_Xprint
        print.C_Xprint("rpt060030.txt", filename)
        print.NewPage()
        Prt_06003(print)
        print.EndFile()
        If (print.ErrMsg <> "") Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('" & print.ErrMsg & "');")
            Response.Write("</script>")
        Else
            Response.Write("<script language='javascript'>")
            'Response.Write("window.onload = function() {")
            'Response.Write("window.location.replace('../../drs/" & filename & "');")
            Response.Write("window.parent.location='../00/MOA00020.aspx?x=MOA06001&y=D6Y95Y5XSU&fn=" & filename & "';")
            'Response.Write("}")
            Response.Write("</script>")
        End If

    End Sub

    Protected Sub Prt_06003(ByVal prt As C_Xprint)

        sql = "SELECT PWUNIT, PWTITLE, PWNAME, PWIDNO, PAUNIT, PANAME, PATITLE, PAIDNO, nAPPTIME, nREASON, convert(nvarchar,nDATE,111) nDATE, nPLACE"
        sql += " FROM P_06"
        sql += " where EFORMSN='" + eformsn + "'"

        If do_sql.db_sql(sql, do_sql.G_conn_string) = False Then
            Exit Sub
        End If
        If do_sql.G_table.Rows.Count > 0 Then
            D2 = do_sql.G_table
        Else
            Exit Sub
        End If
        sql = "SELECT Info_Num, nKind, nMName, nDocnum, nAmount, nContent, nClass"
        sql += " FROM P_0601"
        sql += " where EFORMSN='" + eformsn + "' order by 2"
        If do_sql.db_sql(sql, do_sql.G_conn_string) = False Then
            Exit Sub
        End If

        Dim h2 As Int16 = 93
        Dim i As Int16
        Dim append_file As Boolean = False
        Dim hi As Integer
        For i = 0 To 2
            D1 = do_sql.G_table

            prt.Add("申請出入日期", CDate(D2.Rows(0).Item("nDATE").ToString()).ToString("yyyy/MM/dd"), 0, i * h2)
            prt.Add("申請人", D2.Rows(0).Item("PANAME").ToString(), 0, i * h2)
            If D2.Rows(0).Item("nREASON").ToString() = "其他" Then
                prt.Add("事由_其他", "V", 0, i * h2)
            Else
                prt.Add(D2.Rows(0).Item("nREASON").ToString(), "V", 0, i * h2)
            End If
            prt.Add("使用地點", D2.Rows(0).Item("nPLACE").ToString(), 0, i * h2)

            Dim block_name As String = ""
            Dim si As String = ""
            Dim tabled(7) As Integer
            For hi = 1 To 6
                tabled(hi) = 0
            Next
            For hi = 0 To D1.Rows.Count - 1
                Select Case D1.Rows(hi).Item("nKind").ToString()
                    Case "電腦主機"
                        si = "1"
                    Case "磁片光碟"
                        si = "2"
                    Case "行動硬碟"
                        si = "3"
                    Case "文書圖表"
                        si = "4"
                    Case "保密裝備"
                        si = "5"
                    Case Else
                        si = "6"
                End Select
                tabled(CInt(si)) = tabled(CInt(si)) + 1
                If tabled(CInt(si)) < 2 Then
                    If D1.Rows(hi).Item("nKind").ToString() = "其他" Then
                        prt.Add("區分_其他", "V", 0, i * h2)
                    Else
                        prt.Add(D1.Rows(hi).Item("nKind").ToString(), "V", 0, i * h2)
                    End If
                    prt.Add("名稱機型_" & D1.Rows(hi).Item("nKind").ToString(), D1.Rows(hi).Item("nMName").ToString(), 0, i * h2)
                    prt.Add("編（文）號_" & D1.Rows(hi).Item("nKind").ToString(), D1.Rows(hi).Item("nDocnum").ToString(), 0, i * h2)
                    prt.Add("數量_" & D1.Rows(hi).Item("nKind").ToString(), D1.Rows(hi).Item("nAmount").ToString(), 0, i * h2)
                    prt.Add("內容概要_" & D1.Rows(hi).Item("nKind").ToString(), D1.Rows(hi).Item("nContent").ToString(), 0, (i * h2) - 1)
                    prt.Add("機密等級_" & D1.Rows(hi).Item("nClass").ToString() & "_" & si, "V", 0, i * h2)
                Else
                    append_file = True
                End If
            Next
        Next
        '附件
        If append_file = True Then
            prt.NewPage("rpt060031.txt")

            Dim h As Integer = 8
            Dim row As Integer = 0
            For hi = 0 To D1.Rows.Count - 1
                prt.Add("區分", D1.Rows(hi).Item("nKind").ToString(), 0, row * h)
                prt.Add("名稱機型", D1.Rows(hi).Item("nMName").ToString(), 0, row * h)
                prt.Add("編（文）號", D1.Rows(hi).Item("nDocnum").ToString(), 0, row * h)
                prt.Add("數量", Space(9 - D1.Rows(hi).Item("nAmount").ToString().Length) + D1.Rows(hi).Item("nAmount").ToString(), 0, row * h)
                prt.Add("內容概要", D1.Rows(hi).Item("nContent").ToString(), 0, row * h)
                prt.Add("機密等級", D1.Rows(hi).Item("nClass").ToString(), 0, row * h)
                row = row + 1
            Next
        End If
    End Sub

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "185px"
        Div_grid.Style("left") = "300px"

    End Sub
    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

        Txt_nDATE.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False

    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

        Div_grid.Visible = False

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
            Div_grid10.Style("Top") = "375px"
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
