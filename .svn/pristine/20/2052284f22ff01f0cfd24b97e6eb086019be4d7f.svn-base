Imports System.Data
Imports System.IO
Imports System.Data.SqlClient

Partial Class Source_03_MOA03001
    Inherits System.Web.UI.Page
    Public do_sql As New C_SQLFUN
    Dim dr As System.Data.DataRow
    Dim n_table As New System.Data.DataTable
    Dim n_table2 As New System.Data.DataTable
    Dim n_table3 As New System.Data.DataTable
    Public print_file As String = ""
    Public grid_table As New System.Data.DataTable
    Public eformsn As String
    Public d_pub As New C_Public
    Dim eformid As String
    Dim str_ORG_UID As String
    Public stmt As String
    Dim p As Integer = 0
    Dim K As Integer = 0
    Dim str_app_id As String
    Public read_only As String = ""
    Dim str_nSTATUS As String = ""
    '---for date
    Public date1_flag As Boolean = False
    Public date1_x As Integer = 273
    Public date1_y As Integer = 249
    Dim connstr, user_id, org_uid As String
    Dim AgentEmpuid As String = ""
    Dim nRECDATE_Flag As String = ""
    '判斷前端頁面是否顯示列印按鈕用變數,表單關卡為倒數最後一關或是已通過全部關卡審核完成才顯示
    Public strDisplayPrintButton As String = ""
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

        do_sql.G_errmsg = ""
        do_sql.G_user_id = Session("user_id") '"tempu180"
        'user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        eformsn = Request("eformsn") '"VJ" 'd_pub.randstr(2) '
        eformid = Request("eformid")
        read_only = Request("read_only")
        nRECDATE_Flag = Request("nRECDATE_Flag")

        'session被清空回首頁
        If user_id = "" And org_uid = "" And eformsn = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.parent.location='../../index.aspx';")
            Response.Write(" </script>")

        Else

            '新增連線
            connstr = do_sql.G_conn_string
            '取得表單審核關卡是否已全部完成,或已審核至倒數最後一關,是的話則將該變數設定為YES,
            '用以判斷前端頁面是否顯示列印按鈕
            strDisplayPrintButton = GetIfFlowStepsInCondition(eformsn)

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
                If RdrAgentCheck.Read() Then
                    AgentEmpuid = RdrAgentCheck.Item("empuid")
                End If
                db.Close()

            End If

            If IsPostBack Then
                Exit Sub
            End If
            'eformsn = d_pub.randstr(2) 'Request("eformsn") '

            If read_only = "1" Or read_only = "2" Then

                ImgCalender.Visible = False

                If select_data() = False Then
                    Exit Sub
                End If
                Exit Sub
            End If
            Lab_time.Text = Now.ToString("yyyy/MM/dd HH:mm:ss")
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
                stmt = "select * from Employee where ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ")  order by emp_chinese_name"

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

        End If

    End Sub

    ''' <summary>
    ''' '取得該表單之Flow流程是否已跑至最後一關或已全部審核完成
    ''' </summary>
    ''' <param name="EFORMSN">表單ID</param>
    ''' <returns>空字串或YES</returns>
    ''' <remarks></remarks>
    Private Function GetIfFlowStepsInCondition(ByVal EFORMSN As String) As String
        Dim str As String = ""
        Dim db As New SqlConnection(connstr)
        db.Open()
        Dim dt As New DataTable
        Dim ds As New DataSet
        Dim strComm As String = "SELECT * FROM flowctl WHERE eformsn = '" + EFORMSN + "' AND stepsid IN (" + GetLastAndCompleteStepID(db) + ")"
        Dim da As SqlDataAdapter = New SqlDataAdapter(strComm, db)
        da.Fill(ds)
        dt = ds.Tables(0)
        db.Close()
        '若符合條件則回傳YES字串
        If dt.Rows.Count > 0 Then
            str = "YES"
        End If
        Return str
    End Function

    ''' <summary>
    ''' 取得倒數最後一關及審核完成之關卡ID
    ''' </summary>
    ''' <param name="db">SqlConnection</param>
    ''' <returns>關卡ID字串</returns>
    ''' <remarks></remarks>
    Private Function GetLastAndCompleteStepID(ByRef db As SqlConnection) As String
        Dim sqlComm As New SqlCommand("SELECT TOP 2 stepsid FROM flow WHERE eformid='j2mvKYe3l9' ORDER BY y", db)
        Dim reader = sqlComm.ExecuteReader()
        Dim str As String = "''"
        While reader.Read()
            str = str + ",'" + reader.Item("stepsid").ToString() + "'"
        End While
        reader.Close()
        Return str
    End Function

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
        stmt = "select * from p_0302 order by pck_name"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Sub
        End If
        n_table = do_sql.G_table
        DrDown_nSTYLE.Items.Clear()
        p = 0
        DrDown_nSTYLE.Items.Add("")
        DrDown_nSTYLE.Items(p).Value = ""
        p += 1
        For Each dr In n_table.Rows
            DrDown_nSTYLE.Items.Add(Trim(dr("pck_name").ToString.Trim))
            DrDown_nSTYLE.Items(p).Value = Trim(dr("pck_name").ToString.Trim)
            p += 1
        Next

        DrDown_nSTHOUR.Items.Clear()
        For p = 0 To 23
            DrDown_nSTHOUR.Items.Add(p.ToString("00"))
            DrDown_nSTHOUR.Items(p).Value = p.ToString("00")
        Next
        DrDown_nEDHOUR.Items.Clear()
        For p = 0 To 59
            DrDown_nEDHOUR.Items.Add(p.ToString("00"))
            DrDown_nEDHOUR.Items(p).Value = p.ToString("00")
        Next
        '
        DrDown_nSTUSEHOUR.Items.Clear()
        For p = 0 To 23
            DrDown_nSTUSEHOUR.Items.Add(p.ToString("00"))
            DrDown_nSTUSEHOUR.Items(p).Value = p.ToString("00")
        Next
        DrDown_nSTUSEMIN.Items.Clear()
        For p = 0 To 59
            DrDown_nSTUSEMIN.Items.Add(p.ToString("00"))
            DrDown_nSTUSEMIN.Items(p).Value = p.ToString("00")
        Next
        '
        DrDown_nEDUSEHOUR.Items.Clear()
        For p = 0 To 23
            DrDown_nEDUSEHOUR.Items.Add(p.ToString("00"))
            DrDown_nEDUSEHOUR.Items(p).Value = p.ToString("00")
        Next
        DrDown_nEDUSEMIN.Items.Clear()
        For p = 0 To 59
            DrDown_nEDUSEMIN.Items.Add(p.ToString("00"))
            DrDown_nEDUSEMIN.Items(p).Value = p.ToString("00")
        Next
    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged
        Dim tmp_date As String = CDate(Calendar1.SelectedDate).ToString("yyyy/MM/dd")
        If tmp_date <= Now.ToString("yyyy/MM/dd") Then
            do_sql.G_errmsg = "今天以前不能新增"
            Exit Sub
        ElseIf tmp_date = Now.AddDays(1).ToString("yyyy/MM/dd") Then
            If Now.ToString("HH:mm") >= "16:00" Then
                do_sql.G_errmsg = "16:00不能新增明日派車單"
                Exit Sub
            End If
        End If

        Txt_nARRDATE.Text = tmp_date
        Txt_nUSEDATE.Text = Txt_nARRDATE.Text
        Txt_nEDUSEDATE.Text = Txt_nARRDATE.Text

        Div_grid2.Visible = False
    End Sub

    Protected Sub But_ins_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles But_ins.Click
        If ins_p305("1") = False Then
            Exit Sub
        End If
    End Sub
    Function ins_p305(ByVal df As String) As Boolean
        ins_p305 = False
        If GridView1.Rows.Count >= 5 Then
            do_sql.G_errmsg = "路徑起迄點最多只能輸入五筆"
            Exit Function
        End If
        If Txt_GoLocal.Text = "" Then
            do_sql.G_errmsg = "從必須輸入"
            Exit Function
        End If
        If Txt_EndLocal.Text = "" Then
            do_sql.G_errmsg = "至必須輸入"
            Exit Function
        End If


        stmt = "insert into P_0305(EFORMSN," '表單序號
        stmt += "GoLocal,EndLocal)" '起點,迄點
        stmt += " values('" + eformsn + "','" + Txt_GoLocal.Text + "','" + Txt_EndLocal.Text + "')"
        If do_sql.db_exec(stmt, do_sql.G_conn_string) = False Then
            Exit Function
        End If
        ''新增Grid資料:
        'stmt = "select * from p_0305 "
        'stmt += "where EFORMSN ='" + eformsn + "'"
        ''stmt += " order by ncreatedate"
        'If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
        '    Exit Function
        'End If
        'GridView1.DataSource = do_sql.G_table

        GridView1.DataBind()

        Txt_GoLocal.Text = ""
        Txt_EndLocal.Text = ""
        ins_p305 = True
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
        If RdrCheck.Read() Then
            Chknextstep = RdrCheck.Item("nextstep")
        End If
        db.Close()

        Dim strgroup_id As String = ""

        If read_only <> "" Then

            '判斷上一級主管
            db.Open()
            Dim strUpComCheck As New SqlCommand("SELECT group_id FROM flow WHERE eformid = '" & eformid & "' and stepsid = '" & Chknextstep & "' and eformrole=1 ", db)
            Dim RdrUpCheck = strUpComCheck.ExecuteReader()
            If RdrUpCheck.Read() Then
                strgroup_id = RdrUpCheck.Item("group_id")
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

            '新增派車資料
            InsData()

            Call do_sql.commit_tran()
        End If

        If read_only = "2" Then

            Dim strAgentName As String = ""

            If DrDown_nSTYLE.Enabled Then '若為審核表單,且車輛品名型式選單為可選
                If String.IsNullOrEmpty(DrDown_nSTYLE.SelectedItem.Text) Then
                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('車輛品名型式不可為空白');")
                    Response.Write(" <script language='javascript'>")
                    Exit Sub
                Else
                    '則更新表單之車輛品名型式資料
                    UpdateStyleOfCar(db, DrDown_nSTYLE.SelectedItem.Text, eformsn)
                End If
            End If

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
                If RdPer.Read() Then
                    strAgentName = RdPer("emp_chinese_name")
                End If
                db.Close()

                strComment = txtcomment.Text & "(此表單已由" & strAgentName & "代理批核)"

                '增加批核意見
                insComment(strComment, eformsn, AgentEmpuid)

            End If

        End If

        Dim Val_P As String
        Val_P = ""


        '判斷下一關為上一級主管時人數是否超過一人
        If (NextPer > 1 And read_only = "") Or (NextPer > 1 And strgroup_id = "Z860") Then
            Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & eformsn & "&SelFlag=1" & "&AppUID=" & DrDown_PANAME.SelectedValue)
        Else

            If nRECDATE_Flag = "" Then

                '表單審核
                Val_P = FC.F_Send(SendVal, do_sql.G_conn_string)

                Dim PageUp As String = ""

                If read_only = "" Then
                    PageUp = "New"
                End If

                Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & PageUp)
                do_sql.G_errmsg = "新增成功"

            Else

                '表單審核
                Val_P = FC.F_Send(SendVal, do_sql.G_conn_string)

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('批核完成');")
                Response.Write(" window.location='MOA03009.aspx?nRECDATE_Flag=" & nRECDATE_Flag & "';")
                Response.Write(" </script>")

            End If


        End If

    End Sub

    ''' <summary>
    ''' 更新表單之車輛品名型式資料
    ''' </summary>
    ''' <param name="db">SqlConnection</param>
    ''' <param name="strStyle">車輛品名型式</param>
    ''' <param name="strEFORMSN">表單sn</param>
    ''' <remarks></remarks>
    Private Sub UpdateStyleOfCar(ByRef db As SqlConnection, ByVal strStyle As String, ByVal strEFORMSN As String)
        db.Open()
        Dim strComm As New SqlCommand("UPDATE P_03 SET nSTYLE = @nSTYLE WHERE EFORMSN = @EFORMSN", db)
        strComm.Parameters.Add("@nSTYLE", SqlDbType.VarChar, 50).Value = strStyle
        strComm.Parameters.Add("@EFORMSN", SqlDbType.VarChar, 16).Value = strEFORMSN
        strComm.ExecuteNonQuery()
        db.Close()
    End Sub

    Function check_text() As Boolean
        check_text = False
        'If DrDown_PANAME.SelectedIndex < 0 Then
        '    do_sql.G_errmsg = "申請人姓名必須選取"
        '    Exit Function
        'End If

        do_sql.G_errmsg = ""
        If Lab_PAUNIT.Text = "" Then
            do_sql.G_errmsg = "申請人姓名必須選取"
            Exit Function
        End If

        If TXT_nPHONE.Text = "" Then
            do_sql.G_errmsg = "聯絡電話必須輸入"
            Exit Function
        End If
        If TXT_nREASON.Text = "" Then
            do_sql.G_errmsg = "任務理由必須輸入"
            Exit Function
        End If
        If TXT_nITEM.Text = "" Then
            do_sql.G_errmsg = "人員項目必須輸入"
            Exit Function
        End If
        Dim varOutPut As Integer
        If Not Integer.TryParse(TXT_nITEM.Text, varOutPut) Then
            do_sql.G_errmsg = "人員項目只允許輸入數字"
            Exit Function
        End If

        If TXT_nARRIVEPLACE.Text = "" Then
            do_sql.G_errmsg = "車輛報到地點必須輸入"
            Exit Function
        End If

        If TXT_nARRIVETO.Text = "" Then
            do_sql.G_errmsg = "向何人報到必須輸入"
            Exit Function
        End If

        'If DrDown_nCARNUM.SelectedItem.Value = "" Then
        '    do_sql.G_errmsg = "車次數必須輸入"
        '    Exit Function
        'End If

        If Txt_nSTARTPOINT.Text = "" Then
            do_sql.G_errmsg = "起點必須輸入"
            Exit Function
        End If
        If Txt_nENDPOINT.Text = "" Then
            do_sql.G_errmsg = "目的地必須輸入"
            Exit Function
        End If
        If Txt_nARRDATE.Text = "" Then
            do_sql.G_errmsg = "車輛報到日期必須輸入"
            Exit Function
        End If

        Dim tomorrow As Date = Date.Now.Date.AddDays(1)
        Dim applicationDate As Date = Convert.ToDateTime(Txt_nARRDATE.Text).Date
        If DateAndTime.Now.TimeOfDay > New TimeSpan(16, 0, 0) And applicationDate.Equals(tomorrow) Then
            do_sql.G_errmsg = "目前時間已超過下午四點,不可申請當日隔天之派車"
            Exit Function
        End If

        If Radio1.Checked = True Then
            str_nSTATUS = "一般車"
        ElseIf Radio2.Checked = True Then
            str_nSTATUS = "經常性支援"
        ElseIf Radio3.Checked Then             
            str_nSTATUS = "主官管"
        End If
        If DrDown_nSTYLE.SelectedIndex < 1 Then
            do_sql.G_errmsg = "車輛品名型式必須選取"
            Exit Function
        End If
        If Txt_nUSEDATE.Text = "" Then
            do_sql.G_errmsg = "任務使用時間起日期必須輸入"
            Exit Function
        End If
        If Txt_nEDUSEDATE.Text = "" Then
            do_sql.G_errmsg = "任務使用時間迄日期必須輸入"
            Exit Function
        End If
        check_text = True
    End Function
    Function select_data() As Boolean
        select_data = False
        Call Add_DropDownList()
        stmt = "select * from p_03 p,p_0301 p1 "
        stmt += "where p.EFORMSN ='" + eformsn + "' and p.EFORMSN = p1.EFORMSN "
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
            Lab_time.Text = ""

            TXT_nPHONE.Text = ""
            TXT_nREASON.Text = ""
            TXT_nITEM.Text = ""
            TXT_nARRIVEPLACE.Text = ""
            TXT_nARRIVETO.Text = ""
            'DrDown_nCARNUM.SelectedItem.Value = ""
            'TXT_nUSECARNUM.Text = ""
            Txt_nSTARTPOINT.Text = ""
            Txt_nENDPOINT.Text = ""
            Txt_nARRDATE.Text = ""
            DrDown_nSTHOUR.SelectedIndex = 0
            DrDown_nEDHOUR.SelectedIndex = 0
            Radio1.Checked = True
            Radio1.Checked = False
            Radio1.Checked = False

            Txt_nUSEDATE.Text = ""
            DrDown_nSTUSEHOUR.SelectedIndex = 0
            DrDown_nSTUSEMIN.SelectedIndex = 0

            Txt_nEDUSEDATE.Text = ""
            DrDown_nEDUSEHOUR.SelectedIndex = 0
            DrDown_nEDUSEMIN.SelectedIndex = 0

            LabCarNum.Text = ""

        Else
            Lab_PWUNIT.Text = do_sql.G_table.Rows(0).Item("PWUNIT").ToString
            Lab_PWNAME.Text = do_sql.G_table.Rows(0).Item("PWNAME").ToString
            Lab_PWTITLE.Text = do_sql.G_table.Rows(0).Item("PWTITLE").ToString
            Lab_PAUNIT.Text = do_sql.G_table.Rows(0).Item("PAUNIT").ToString
            DrDown_PANAME.Items.Clear()
            DrDown_PANAME.Items.Add(Trim(do_sql.G_table.Rows(0).Item("PANAME").ToString))
            DrDown_PANAME.Items(0).Value = Trim(do_sql.G_table.Rows(0).Item("PANAME").ToString)

            Lab_PATITLE.Text = do_sql.G_table.Rows(0).Item("PATITLE").ToString
            Lab_time.Text = do_sql.G_table.Rows(0).Item("nAPPLYTIME").ToString

            TXT_nPHONE.Text = do_sql.G_table.Rows(0).Item("nPHONE").ToString
            TXT_nREASON.Text = do_sql.G_table.Rows(0).Item("nREASON").ToString
            TXT_nITEM.Text = do_sql.G_table.Rows(0).Item("nITEM").ToString
            TXT_nARRIVEPLACE.Text = do_sql.G_table.Rows(0).Item("nARRIVEPLACE").ToString
            TXT_nARRIVETO.Text = do_sql.G_table.Rows(0).Item("nARRIVETO").ToString
            'DrDown_nCARNUM.SelectedItem.Value = do_sql.G_table.Rows(0).Item("nCARNUM").ToString
            'TXT_nUSECARNUM.Text = do_sql.G_table.Rows(0).Item("nUSECARNUM").ToString
            Txt_nSTARTPOINT.Text = do_sql.G_table.Rows(0).Item("nSTARTPOINT").ToString
            Txt_nENDPOINT.Text = do_sql.G_table.Rows(0).Item("nENDPOINT").ToString
            Txt_nARRDATE.Text = CDate(do_sql.G_table.Rows(0).Item("nARRDATE").ToString).ToString("yyyy/MM/dd")

            do_sql.set_drop_down(DrDown_nSTHOUR, do_sql.G_table.Rows(0).Item("nSTHOUR").ToString.Trim)
            do_sql.set_drop_down(DrDown_nEDHOUR, do_sql.G_table.Rows(0).Item("nEDHOUR").ToString.Trim)
            If do_sql.G_table.Rows(0).Item("nSTATUS").ToString = "一般車" Then
                Radio1.Checked = True
                Radio2.Checked = False
                Radio3.Checked = False
            ElseIf do_sql.G_table.Rows(0).Item("nSTATUS").ToString = "經常性支援" Then
                Radio1.Checked = False
                Radio2.Checked = True
                Radio3.Checked = False
            Else
                Radio1.Checked = False
                Radio2.Checked = False
                Radio3.Checked = True
            End If
            'DrDown_nSTYLE.SelectedItem.Value = "1" 'do_sql.G_table.Rows(0).Item("nSTYLE").ToString.Trim
            do_sql.set_drop_down(DrDown_nSTYLE, do_sql.G_table.Rows(0).Item("nSTYLE").ToString.Trim)

            Txt_nUSEDATE.Text = CDate(do_sql.G_table.Rows(0).Item("nUSEDATE").ToString).ToString("yyyy/MM/dd")
            do_sql.set_drop_down(DrDown_nSTUSEHOUR, do_sql.G_table.Rows(0).Item("nSTUSEHOUR").ToString.Trim)
            do_sql.set_drop_down(DrDown_nSTUSEMIN, do_sql.G_table.Rows(0).Item("nSTUSEMIN").ToString.Trim)

            Txt_nEDUSEDATE.Text = CDate(do_sql.G_table.Rows(0).Item("nEDUSEDATE").ToString).ToString("yyyy/MM/dd")
            do_sql.set_drop_down(DrDown_nEDUSEHOUR, do_sql.G_table.Rows(0).Item("nEDUSEHOUR").ToString.Trim)
            do_sql.set_drop_down(DrDown_nEDUSEMIN, do_sql.G_table.Rows(0).Item("nEDUSEMIN").ToString.Trim)

            LabCarNum.Text = do_sql.G_table.Rows(0).Item("CarNumber").ToString
        End If

        'GridView1.DataBind()
        Call disabel_field()
        select_data = True
    End Function
    Sub disabel_field()

        DrDown_PANAME.Enabled = False
        TXT_nPHONE.ReadOnly = True
        TXT_nREASON.ReadOnly = True
        TXT_nITEM.ReadOnly = True
        TXT_nARRIVEPLACE.ReadOnly = True
        TXT_nARRIVETO.ReadOnly = True
        'DrDown_nCARNUM.Enabled = False
        'TXT_nUSECARNUM.ReadOnly = True
        Txt_nSTARTPOINT.ReadOnly = True
        Txt_nENDPOINT.ReadOnly = True
        Txt_nARRDATE.ReadOnly = True
        DrDown_nSTHOUR.Enabled = False
        DrDown_nEDHOUR.Enabled = False
        Radio1.Enabled = False
        Radio2.Enabled = False
        Radio3.Enabled = False
        '若當前表單審核關卡非第一次之派車調度單位或派車管制單位
        If Not IsThirdOrForthStep(eformsn, eformid) Then
            '則車種品名型式選單設為disable(第一次之派車調度單位或派車管制單位審核關卡可再重新選擇車種品名型式)
            DrDown_nSTYLE.Enabled = False
        End If
        Txt_nUSEDATE.ReadOnly = True
        DrDown_nSTUSEHOUR.Enabled = False
        DrDown_nSTUSEMIN.Enabled = False

        Txt_nEDUSEDATE.ReadOnly = True
        DrDown_nEDUSEHOUR.Enabled = False
        DrDown_nEDUSEMIN.Enabled = False
        GridView1.Enabled = False
        If read_only = "1" Then
            But_exe.Visible = False
        End If
        But_ins.Visible = False
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="eformsn">表單sn</param>
    ''' <param name="eformid">表單種類ID</param>
    ''' <returns>目前表單審核關卡是否為第一次之派車調度單位</returns>
    ''' <remarks></remarks>
    Private Function IsThirdOrForthStep(ByVal eformsn As String, ByVal eformid As String) As Boolean
        Dim result As Boolean = False
        Dim db As New SqlConnection(connstr)
        db.Open()

        Dim currentFlowStepsID As String = GetStepsID(db, eformsn)  '表單目前審核關卡ID
        Dim thirdStepsID As String = GetThirdStepsID(db, eformid) '第三審核關卡(即第一次之派車調度單位)之stepsid
        Dim forthStepsID As String = GetForthStepsID(db, eformid) '第四審核關卡(即第一次之派車調度單位)之stepsid
        If Not String.IsNullOrEmpty(currentFlowStepsID) And Not String.IsNullOrEmpty(thirdStepsID) _
            And (currentFlowStepsID.Equals(thirdStepsID) Or currentFlowStepsID.Equals(forthStepsID)) Then
            '若目前表單審核關卡為第一次之派車調度單位或派車管制單位
            result = True '則回傳表單審核關卡目前為第一次之派車調度單位或即派車管制單位
        End If

        db.Close()
        Return result
    End Function

    ''' <summary>
    ''' 取得表單之stepsid
    ''' </summary>
    ''' <param name="db">SqlConnection</param>
    ''' <param name="eformsn">表單sn</param>
    ''' <returns>表單stepsid</returns>
    ''' <remarks></remarks>
    Private Function GetStepsID(ByRef db As SqlConnection, ByVal eformsn As String) As String
        Dim stepsID As String = ""

        Dim sqlComm = New SqlCommand("SELECT TOP 1 stepsid FROM flowctl WHERE eformsn = @eformsn ORDER BY flowsn DESC", db)
        sqlComm.Parameters.Add("@eformsn", SqlDbType.VarChar, 16).Value = eformsn
        stepsID = sqlComm.ExecuteScalar()

        Return stepsID
    End Function

    ''' <summary>
    ''' 取得第三審核關卡(即第一次之派車調度單位)之stepsid
    ''' </summary>
    ''' <param name="db">SqlConnection</param>
    ''' <param name="eformid">表單種類ID</param>
    ''' <returns>第三審核關卡(即第一次之派車調度單位)之stepsid</returns>
    ''' <remarks></remarks>
    Private Function GetThirdStepsID(ByRef db As SqlConnection, ByVal eformid As String) As String
        Dim stepsID As String = ""

        Dim strComm As String = "SELECT f.stepsid FROM flow AS f JOIN SYSTEMOBJ AS S ON f.group_id = S.object_uid "
        strComm = strComm + "WHERE f.eformid = @eformid AND S.object_name = N'派車調度單位' AND f.nextstep <> -1"
        Dim sqlComm = New SqlCommand(strComm, db)
        sqlComm.Parameters.Add("@eformid", SqlDbType.VarChar, 10).Value = eformid
        stepsID = sqlComm.ExecuteScalar().ToString()

        Return stepsID
    End Function

    ''' <summary>
    ''' 取得第四審核關卡(即派車管制單位)之stepsid
    ''' </summary>
    ''' <param name="db">SqlConnection</param>
    ''' <param name="eformid">表單種類ID</param>
    ''' <returns>第四審核關卡(即派車管制單位)之stepsid</returns>
    ''' <remarks></remarks>
    Private Function GetForthStepsID(ByRef db As SqlConnection, ByVal eformid As String) As String
        Dim stepsID As String = ""

        Dim strComm As String = "SELECT f.stepsid FROM flow AS f JOIN SYSTEMOBJ AS S ON f.group_id = S.object_uid "
        strComm = strComm + "WHERE f.eformid = @eformid AND S.object_name = N'派車管制單位'"
        Dim sqlComm = New SqlCommand(strComm, db)
        sqlComm.Parameters.Add("@eformid", SqlDbType.VarChar, 10).Value = eformid
        stepsID = sqlComm.ExecuteScalar().ToString()

        Return stepsID
    End Function

    'Sub set_drop_down(ByVal drop_box As DropDownList, ByVal value As String)
    '    Dim box_index As Integer = drop_box.Items.Count
    '    Dim ci As Integer
    '    For ci = 0 To box_index - 1
    '        If drop_box.Items(ci).Value = value Then
    '            drop_box.SelectedIndex = ci
    '            Exit Sub
    '        End If
    '    Next
    'End Sub

    Protected Sub Calendar1_VisibleMonthChanged(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.MonthChangedEventArgs) Handles Calendar1.VisibleMonthChanged
        date1_flag = True
    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call add_grid()
    End Sub

    Protected Sub But_cencel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles But_cencel.Click
        Div_grid.Visible = False
        Div_grid.Style("Top") = "624px"
        Div_grid.Style("left") = "237px"
    End Sub
    Sub add_grid()
        If Txt_nARRDATE.Text = "" Then
            do_sql.G_errmsg = "車輛報到日期必需輸入"
            Exit Sub
        End If
        stmt = "select A.pck_name,(A.PCN_Use-isnull(B.per,0)) As num_car from "
        stmt += "(select pck_name,PCN_Use from p_0306 where  pcn_date = '" + Txt_nARRDATE.Text + "') A left outer join "
        stmt += "(select nSTYLE,count(*) As per from P_03  where  nARRDATE = '" + Txt_nARRDATE.Text + "'"
        stmt += " group by nSTYLE) B  on A.pck_name=B.nSTYLE"
        stmt += " order by A.pck_name"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Sub
        End If
        If do_sql.G_table.Rows.Count = 0 Then
            do_sql.G_errmsg = "查無車輛資料"
            Exit Sub
        End If
        grid_table = do_sql.G_table
        'GridView2.DataSource = grid_table
        'GridView2.DataBind()
        SqlDataSource2.SelectCommand = stmt
        Div_grid.Visible = True
        Div_grid.Style("Top") = "122px"
        Div_grid.Style("left") = "400px"
    End Sub

    Protected Sub GridView2_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView2.PageIndexChanging
        add_grid()
    End Sub

    Protected Sub supplyBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles supplyBtn.Click

        Try

            '判斷是否輸入資料
            If check_text() = True Then

                '新增申請派車資料
                InsData()

                Dim Val_P As String
                Val_P = ""

                '表單補登
                Dim FC As New C_FlowSend.C_FlowSend

                Dim SendVal As String = ""

                '判斷是否為代理人批核的表單
                If AgentEmpuid = "" Then
                    SendVal = eformid & "," & user_id & "," & eformsn & "," & "1"
                Else
                    SendVal = eformid & "," & AgentEmpuid & "," & eformsn & "," & "1"
                End If

                Val_P = FC.F_Supply(SendVal, connstr)

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('補登完成');")
                Response.Write(" window.parent.location='../00/MOA00020.aspx?x=MOA03001&y=j2mvKYe3l9';")
                Response.Write(" </script>")

            End If

        Catch ex As Exception

        End Try

    End Sub

    Public Sub InsData()

        '新增資料
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '新增會議日期資料
        db.Open()

        Dim InsP03 As String = ""
        InsP03 = "insert into P_03(EFORMSN,PWUNIT,PWTITLE," '表單序號,填表人單位,填表人級職
        InsP03 += "PWNAME, PWIDNO, PAUNIT," '填表人姓名,填表人身份證字號,申請人單位
        InsP03 += "PANAME, PATITLE, PAIDNO," '申請人姓名,申請人級職,申請人身份證字號 "
        InsP03 += "nAPPLYTIME, nPHONE,nREASON, " '申請時間,聯絡電話,申請事由,
        InsP03 += "nITEM,nARRIVEPLACE,nARRIVETO," '人員項目,車輛報到地點,向何人報到        
        InsP03 += "nCARNUM,nUSECARNUM,nSTARTPOINT, " '車次數,使用車輛數,起點        
        InsP03 += "nENDPOINT,nARRDATE,nSTHOUR,nEDHOUR," '目的地, 車輛報到日期,車輛報到日期(時),車輛報到日期(分)     
        InsP03 += "nSTATUS,nSTYLE,nUSEDATE,nSTUSEHOUR,nSTUSEMIN," '車輛狀態(車輛類型),車輛品名型式,任務使用時間(起),任務使用時間(時)-起,任務使用時間(分)-起     
        InsP03 += "nEDUSEDATE,nEDUSEHOUR,nEDUSEMIN) " '任務使用時間(迄),任務使用時間(時)-迄,任務使用時間(分)-迄
        InsP03 += " values('" + eformsn + "','" + Lab_PWUNIT.Text + "','" + Lab_PWTITLE.Text + "',N'"
        InsP03 += Lab_PWNAME.Text + "','" + do_sql.G_user_id + "','" + Lab_PAUNIT.Text + "',N'"
        InsP03 += DrDown_PANAME.SelectedItem.Text + "','" + Lab_PATITLE.Text + "','" + DrDown_PANAME.SelectedItem.Value + "','"
        InsP03 += Lab_time.Text + "','" + TXT_nPHONE.Text.Replace("'", "''") + "','" + TXT_nREASON.Text + "','"
        InsP03 += TXT_nITEM.Text + "','" + TXT_nARRIVEPLACE.Text + "','" + TXT_nARRIVETO.Text + "',"
        InsP03 += "'1','1','" + Txt_nSTARTPOINT.Text + "','"
        InsP03 += Txt_nENDPOINT.Text + "','" + Txt_nARRDATE.Text + "','" + DrDown_nSTHOUR.SelectedItem.Value + "','" + DrDown_nEDHOUR.SelectedItem.Value + "','"
        InsP03 += str_nSTATUS + "','" + DrDown_nSTYLE.SelectedItem.Value + "','" + Txt_nUSEDATE.Text + "','" + DrDown_nSTUSEHOUR.SelectedItem.Value + "','" + DrDown_nSTUSEMIN.SelectedItem.Value + "','"
        InsP03 += Txt_nEDUSEDATE.Text + "','" + DrDown_nEDUSEHOUR.SelectedItem.Value + "','" + DrDown_nEDUSEMIN.SelectedItem.Value + "')"

        Dim InsP0301 As String = ""
        InsP0301 = "insert into P_0301(EFORMSN) values('" + eformsn + "')"

        Dim insCom As New SqlCommand(InsP03, db)
        insCom.ExecuteNonQuery()
        Dim InsCom1 As New SqlCommand(InsP0301, db)
        InsCom1.ExecuteNonQuery()

        db.Close()



    End Sub

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
            supplyBtn.Visible = False

            '讀取表單可送件
        ElseIf read_only = "2" Then

            LabCarTitle.Visible = True
            LabCarNum.Visible = True

            supplyBtn.Visible = False

            Dim PerCount As Integer = 0

            '判斷批核者是否為副主管,假如批核者為副主管則不可送件只可呈轉
            Dim ParentFlag As String = ""
            Dim ParentVal = user_id & "," & eformsn

            Dim FC As New C_FlowSend.C_FlowSend
            ParentFlag = FC.F_NextChief(ParentVal, connstr)

            If ParentFlag = "1" Then
                'But_exe.Visible = False

                '上一級主管多少人
                db.Open()
                Dim PerCountCom As New SqlCommand("select count(empuid) as PerCount from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = (select ORG_UID from EMPLOYEE where employee_id ='" & user_id & "'))", db)
                Dim PerRdv = PerCountCom.ExecuteReader()
                If PerRdv.Read() Then
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
                supplyBtn.Visible = False
            End If

        Else
            backBtn.Visible = False
            tranBtn.Visible = False
        End If

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

    Public Sub insComment(ByVal strComment As String, ByVal strEformsn As String, ByVal strid As String)

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

    End Sub

    Protected Sub ImgCalender_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgCalender.Click

        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "150px"
        Div_grid2.Style("left") = "370px"

    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

        Div_grid2.Visible = False

    End Sub

    Protected Sub ImagePrint_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImagePrint.Click
        Dim F_file As String
        Dim F_file2 As String
        Dim F_file_name As String


        stmt = "select * from p_03  "
        stmt += "where EFORMSN ='" + eformsn + "'"

        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            do_sql.G_errmsg = "查詢資料失敗"
            Exit Sub
        End If
        If do_sql.G_table.Rows.Count > 0 Then
            n_table = do_sql.G_table
        Else
            do_sql.G_errmsg = "查無可印資料"
            Exit Sub
        End If
        stmt = "SELECT * FROM p_0305 WHERE eformsn = '" & eformsn & "'"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            do_sql.G_errmsg = "查詢資料失敗"
            Exit Sub
        End If
        If do_sql.G_table.Rows.Count > 0 Then
            n_table2 = do_sql.G_table
        End If
        stmt = "SELECT * FROM flowctl WHERE eformsn = '" & eformsn & "'"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            do_sql.G_errmsg = "查詢資料失敗"
            Exit Sub
        End If
        If do_sql.G_table.Rows.Count > 0 Then
            n_table3 = do_sql.G_table
        End If

        Dim prn_stmt As String = ""

        F_file_name = "rpt030010"
        F_file = MapPath("../../rpt/" + F_file_name + ".txt")
        If File.Exists(F_file) = False Then
            do_sql.G_errmsg = "檔案不存在"
            Exit Sub
        End If

        print_file = "../../drs/" + F_file_name + "-" + Now.ToString("mmssff") & ".drs"
        F_file2 = MapPath(print_file)

        Call do_sql.inc_file(F_file, F_file2, F_file_name)
        Dim tmp_date As String = ""
        Dim nPage As Integer = 0
        Dim n_line As Integer = 0
        nPage += 1
        Call do_sql.print_sdata(F_file2, "/init " + F_file_name + ".txt 1", "/newpage null")
        For Each dr In n_table.Rows
            Call do_sql.print_block(F_file2, "填表人單位", 0, n_line * 7, dr("PWUNIT").ToString)
            Call do_sql.print_block(F_file2, "姓名", 0, n_line * 7, dr("PWNAME").ToString)
            Call do_sql.print_block(F_file2, "級職", 0, n_line * 7, dr("PWTITLE").ToString)
            Call do_sql.print_block(F_file2, "申請人單位", 0, n_line * 7, dr("PAUNIT").ToString)


            Call do_sql.print_block(F_file2, "姓名2", 0, n_line * 7, dr("PANAME").ToString)
            Call do_sql.print_block(F_file2, "級職2", 0, n_line * 7, dr("PATITLE").ToString)
            Call do_sql.print_block(F_file2, "申請時間", 0, n_line * 7, CDate(dr("nAPPLYTIME").ToString).ToString("yyyy/MM/dd HH:mm:ss"))
            Call do_sql.print_block(F_file2, "聯絡電話", 0, n_line * 7, dr("nPHONE").ToString)
            Call do_sql.print_block(F_file2, "事由", 0, n_line * 7, dr("nREASON").ToString)
            Call do_sql.print_block(F_file2, "車輛報到地點", 0, n_line * 7, dr("nARRIVEPLACE").ToString)
            Call do_sql.print_block(F_file2, "起點", 0, n_line * 7, dr("nSTARTPOINT").ToString)
            Call do_sql.print_block(F_file2, "目的地", 0, n_line * 7, dr("nENDPOINT").ToString)
            tmp_date = CDate(dr("nARRDATE").ToString).ToString("yyyy/MM/dd") & " " & dr("nSTHOUR").ToString & ":" & dr("nEDHOUR").ToString
            Call do_sql.print_block(F_file2, "車輛報到日期", 0, n_line * 7, tmp_date)
            Call do_sql.print_block(F_file2, "人員項目", 0, n_line * 7, dr("nITEM").ToString)
            Call do_sql.print_block(F_file2, "車輛類型", 0, n_line * 7, dr("nSTATUS").ToString)
            Call do_sql.print_block(F_file2, "向何人報到", 0, n_line * 7, dr("nARRIVETO").ToString)
            Call do_sql.print_block(F_file2, "車輛品名型式", 0, n_line * 7, dr("nSTYLE").ToString)
            tmp_date = CDate(dr("nUSEDATE").ToString).ToString("yyyy/MM/dd") & " " & dr("nSTUSEHOUR").ToString & ":" & dr("nSTUSEMIN").ToString
            tmp_date += " 至 " & CDate(dr("nEDUSEDATE").ToString).ToString("yyyy/MM/dd") & " " & dr("nEDUSEHOUR").ToString & ":" & dr("nEDUSEMIN").ToString
            Call do_sql.print_block(F_file2, "任務使用時間", 0, n_line * 7, tmp_date)
        Next
        n_line = 0
        For Each dr In n_table2.Rows
            Call do_sql.print_block(F_file2, "從", 0, n_line * 5, dr("GoLocal").ToString)
            Call do_sql.print_block(F_file2, "至", 0, n_line * 5, dr("EndLocal").ToString)
            n_line += 1
        Next
        n_line = 0
        For Each dr In n_table3.Rows
            If dr("hddate") Is DBNull.Value = True Then
                tmp_date = ""
            Else
                tmp_date = CDate(dr("hddate").ToString).ToString("yyyy/MM/dd HH:mm:ss")
            End If
            Call do_sql.print_block(F_file2, "備註-姓名", 0, n_line * 5, dr("emp_chinese_name").ToString)
            Call do_sql.print_block(F_file2, "備註-群組", 0, n_line * 5, dr("group_name").ToString)
            Call do_sql.print_block(F_file2, "備註-日期", 0, n_line * 5, tmp_date)
            Call do_sql.print_block(F_file2, "備註", 0, n_line * 5, dr("comment").ToString)
            n_line += 1
        Next

        'InsP03 = "insert into P_03(EFORMSN,PWUNIT,PWTITLE," '表單序號,填表人單位,填表人級職
        'InsP03 += "PWNAME, PWIDNO, PAUNIT," '填表人姓名,填表人身份證字號,申請人單位
        'InsP03 += "PANAME, PATITLE, PAIDNO," '申請人姓名,申請人級職,申請人身份證字號 "
        'InsP03 += "nAPPLYTIME, nPHONE,nREASON, " '申請時間,聯絡電話,申請事由,
        'InsP03 += "nITEM,nARRIVEPLACE,nARRIVETO," '人員項目,車輛報到地點,向何人報到        
        'InsP03 += "nCARNUM,nUSECARNUM,nSTARTPOINT, " '車次數,使用車輛數,起點        
        'InsP03 += "nENDPOINT,nARRDATE,nSTHOUR,nEDHOUR," '目的地, 車輛報到日期,車輛報到日期(時),車輛報到日期(分)     
        'InsP03 += "nSTATUS,nSTYLE,nUSEDATE,nSTUSEHOUR,nSTUSEMIN," '車輛狀態(車輛類型),車輛品名型式,任務使用時間(起),任務使用時間(時)-起,任務使用時間(分)-起     
        'InsP03 += "nEDUSEDATE,nEDUSEHOUR,nEDUSEMIN) " '任務使用時間(迄),任務使用時間(時)-迄,任務使用時間(分)-迄
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
        If PerRdv.Read() Then
            Visitors = PerRdv("Visitors")
        End If
        db.Close()

        If Visitors > 0 Then   '是否需要依類別分類;如請假,派車等
            Div_grid10.Visible = True
            Div_grid10.Style("Top") = "490px"
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
