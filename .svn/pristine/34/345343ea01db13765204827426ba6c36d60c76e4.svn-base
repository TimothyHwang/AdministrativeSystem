Imports System.Data
Imports System.Data.SqlClient
Imports System.IO


Partial Class M_Source_08_MOA08001
    Inherits System.Web.UI.Page
    Public FCC As New CFlowSend
    Public eformid As String = ""
    Public do_sql As New C_SQLFUN
    Public h_count As New C_DATESUM
    Public d_pub As New C_Public
    Dim dr As System.Data.DataRow
    Dim n_table As New System.Data.DataTable
    Dim str_ORG_UID As String = ""
    Dim str_app_id As String = "" '申請人身份證字號
    Public str_EFORMSN As String = ""
    Public read_only As String = ""
    Public tu_id As String = ""
    Dim connstr, user_id, org_uid As String
    Public executeDBScript As String = String.Empty
    Dim AgentEmpuid As String = ""
    Dim CP As New C_Public
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

        Dim stmt As String = ""
        Dim p As Integer = 0
        Dim K As Integer = 0
        do_sql.G_errmsg = ""
        do_sql.G_user_id = Session("user_id")
        user_id = Session("user_id")
        ViewState("userId") = user_id
        org_uid = Session("ORG_UID")
        tu_id = Session("TU_ID")
        eformid = Request("eformid")
        str_EFORMSN = Request("eformsn")
        read_only = Request("read_only") '1:查閱表單 2:表單批核 "":表單填寫
        hid_readonly.Text = read_only
        hid_eformsn.Text = str_EFORMSN

        'session被清空回首頁
        If user_id = "" And org_uid = "" And str_EFORMSN = "" Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.parent.location='../../index.aspx';")
            Response.Write(" </script>")
        Else
            '新增連線
            connstr = do_sql.G_conn_string
            '找出表單審核者
            If read_only = "2" Then
                but_exe.Text = "核准"
                If Session("user_id") = "" Then
                    do_sql.G_user_id = Page.User.Identity.Name.Substring(Page.User.Identity.Name.IndexOf("\") + 1)
                End If
                '新增連線
                connstr = do_sql.G_conn_string
                '開啟連線
                Dim db As New SqlConnection(connstr)
                db.Open()
                Dim strAgentCheck As New SqlCommand("SELECT empuid FROM flowctl WHERE eformsn = '" & str_EFORMSN & "' and hddate is null", db)
                Dim RdrAgentCheck = strAgentCheck.ExecuteReader()
                If RdrAgentCheck.Read() Then
                    AgentEmpuid = RdrAgentCheck.Item("empuid")
                End If
                RdrAgentCheck.Close()
                db.Close()
                backBtn.Visible = True
            End If

            If IsPostBack Then
                Exit Sub
            End If

            Dim dbQueryconn As New SqlConnection(do_sql.G_conn_string)
            Dim obQueryStatus As Object
            If read_only = "" Then 'read_only空值表示其為新增表單
                tranBtn.Visible = False '新增表單不顯示呈轉按鈕
                '-1:查無資料,表示未申請過;0:清除此sn ;1:新申請 ;2:已印但未回登 ;3:印列失敗 ;4:補登完畢 ;5:審核不通過(密等為2~5者才有)
                dbQueryconn.Open()
                Dim strQueryStatus As New SqlCommand("select top 1 [Status] from P_08 with(nolock) where PAIDNO = '" + user_id + "' order by LogTime Desc", dbQueryconn)
                obQueryStatus = strQueryStatus.ExecuteScalar()
                If Not obQueryStatus Is Nothing Then
                    hid_status.Text = obQueryStatus.ToString()
                End If
                dbQueryconn.Close()
            End If

            If read_only = "1" Or read_only = "2" Then '非新增表單,即僅只讀取表單
                If select_data() = False Then 'select_data() 讀取及設定資料
                    Exit Sub
                End If
                If read_only = "1" Then
                    dbQueryconn.Open()
                    'Status=5:審核不通過,則不顯示憑據寫入按鈕
                    Dim strQueryStatus As New SqlCommand("select * from p_08 with(nolock) where Status <> 5 AND WriteCard = 0 and EFORMSN ='" + str_EFORMSN + "' and PAIDNO='" + user_id + "' and agreeName is not null and AgreeTime is not null", dbQueryconn)
                    obQueryStatus = strQueryStatus.ExecuteScalar()
                    If Not obQueryStatus Is Nothing Then
                        btWrite.Visible = True
                    End If
                    dbQueryconn.Close()
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
                '找出同級單位以下全部單位
                Dim Org_Down As New C_Public
                stmt = "select * from Employee where ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ") order by emp_chinese_name"
                If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
                    Exit Sub
                End If
                If do_sql.G_table.Rows.Count > 0 Then
                    n_table = do_sql.G_table
                    p = 0
                    K = 0
                    DrDown_emp_chinese_name.Items.Clear()
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
                        Call DrDown_emp_chinese_name_SelectedIndexChanged(sender, e)
                    End If
                End If
            End If

            Call drow_txt()
        End If
    End Sub
    Protected Sub DrDown_emp_chinese_name_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrDown_emp_chinese_name.PreRender
        If read_only = "" Then
            str_app_id = DrDown_emp_chinese_name.SelectedValue
        End If
    End Sub

    Protected Sub DrDown_emp_chinese_name_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrDown_emp_chinese_name.SelectedIndexChanged
        If do_sql.select_urname(DrDown_emp_chinese_name.Items(DrDown_emp_chinese_name.SelectedIndex).Value) = False Then
            Lab_ORG_NAME_2.Text = ""
            Exit Sub
        End If
    End Sub
    Protected Sub but_exe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles but_exe.Click
        Dim FC As New C_FlowSend.C_FlowSend '宣告C_FlowSend物件 FC
        Dim SendVal As String = ""          '宣告字串變數 SendVal
        If AgentEmpuid = "" Then            '判斷是否為批核表單
            '若代理人ID為空值,即非代理人批核,SendVal則為: 表單種類ID,本人ID,表單ID,1,
            'SendVal = eformid & "," & do_sql.G_user_id & "," & str_EFORMSN & "," & "1" & ","
            SendVal = eformid & "," & DrDown_emp_chinese_name.SelectedValue & "," & str_EFORMSN & "," & "1" & ","
        Else '反之為代理人批核,SendVal則為: 表單種類ID,代理人ID,表單ID,1,
            SendVal = eformid & "," & AgentEmpuid & "," & str_EFORMSN & "," & "1" & ","
        End If
        Dim NextPer As Integer = 0                '上一級主管人數
        Dim db As New SqlConnection(connstr)      '開啟連線
        NextPer = FC.F_NextStep(SendVal, connstr) '關卡為上一級主管有多少人
        Dim Chknextstep As Integer = GetNextStepID(db, str_EFORMSN, user_id) '取得下一關卡ID

        Dim strgroup_id As String = "" '表單關卡群組ID
        If read_only <> "" Then        'read_only非空值,為非新申請表單,即審核or else
            db.Open()                  '判斷上一級主管
            Dim strUpComCheck As New SqlCommand("SELECT group_id FROM flow WHERE eformid = '" & eformid & "' and stepsid = '" & Chknextstep & "' and eformrole=1 ", db)
            Dim RdrUpCheck = strUpComCheck.ExecuteReader()
            If RdrUpCheck.Read() Then
                strgroup_id = RdrUpCheck.Item("group_id")
            End If
            RdrUpCheck.Close()
            db.Close()

        End If

        If NextPer = 0 And read_only = "" Then '若為新申請表單且無上一級主管,則return
            ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "MngMsg", "alert('無上一級主管');", True)
            Exit Sub
        End If

        do_sql.G_errmsg = ""
        '*************** Initial of If read_only <> "2" ***************
        If read_only <> "2" Then '2:送審表單(資料欄位為readOnly格式)
            Dim sStatus As String = -1
            Dim dbQueryconn As New SqlConnection(do_sql.G_conn_string)
            dbQueryconn.Open()
            Dim strQueryStatus As New SqlCommand("select top 1 [Status] from P_08 with(nolock) where PAIDNO = '" + user_id + "' order by LogTime Desc", dbQueryconn)
            Dim obQueryStatus As Object = strQueryStatus.ExecuteScalar()
            If Not obQueryStatus Is Nothing Then
                sStatus = obQueryStatus.ToString()
            End If
            dbQueryconn.Close()
            Dim sErrorMsg As String = String.Empty
            '-1:查無資料,表示未申請過;0:清除此sn ;1:新申請 ;2:已印但未回登 ;3:印列失敗 ;4:補登完畢 ;5:審核不通過(密等為2~5者才有) ;6:申請人取消
            Select Case sStatus
                Case "-1", "0", "3", "4", "5", "6" '可以再申請
                    Exit Select
                Case "1" '不可申請
                    bt_clear.Visible = True
                    Exit Sub
                Case "2"  '不可申請
                    Exit Sub
                Case Else
                    sErrorMsg = "目前系統忙碌中，請您稍候再試，或通知系統管理人員，謝謝！"
            End Select

            If sErrorMsg <> String.Empty Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('" + sErrorMsg + "');")
                Response.Write(" </script>")
                Exit Sub
            End If

            Dim stmt As String
            Dim str_nSTATUS As String = ""
            Dim str_nPROVEMENT As String = ""

            stmt = "select EFORMSN from P_08 where EFORMSN='" + str_EFORMSN + "'"
            If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
                Exit Sub
            End If

            If do_sql.G_table.Rows.Count > 0 Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('EFORMSN已重複,請重新load form');")
                Response.Write(" </script>")
                Exit Sub
            End If

            If tbFile_Name.Text.Trim().Length < 1 Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('[文件資料名稱]不可空白！');")
                Response.Write(" </script>")
                Exit Sub
            End If

            If tbUse_For.Text.Trim().Length < 1 Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('[用途]不可空白！');")
                Response.Write(" </script>")
                Exit Sub
            End If

            Dim out As Integer
            If txtOriginalSheets.Text.Trim().Length < 1 Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('[原件張數]不可空白！');")
                Response.Write(" </script>")
                Exit Sub
            ElseIf Not Integer.TryParse(txtOriginalSheets.Text.Trim(), out) Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('[原件張數]只可填入數字！');")
                Response.Write(" </script>")
                Exit Sub
            End If

            If txtCopySheets.Text.Trim().Length < 1 Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('[每份複印張數]不可空白！');")
                Response.Write(" </script>")
                Exit Sub
            ElseIf Not Integer.TryParse(txtCopySheets.Text.Trim(), out) Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('[每份複印張數]只可填入數字！');")
                Response.Write(" </script>")
                Exit Sub
            End If

            If txtTotalSheets.Text.Trim().Length < 1 Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('[合計複印張數]不可空白！');")
                Response.Write(" </script>")
                Exit Sub
            ElseIf Not Integer.TryParse(txtTotalSheets.Text.Trim(), out) Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('[合計複印張數]只可填入數字！');")
                Response.Write(" </script>")
                Exit Sub
            End If

            Dim SecurityItem As String = "0"
            If ddl_SecurityItem.Items.Count > 0 Then
                SecurityItem = ddl_SecurityItem.SelectedValue
            End If

            If ddl_Security_Status.SelectedValue <> "1" And SecurityItem = "0" Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('您的申請類別為密等級以上，請先選擇填寫的機密原始文件！');")
                Response.Write(" </script>")
                Exit Sub
            Else
                stmt = "insert into P_08 (EFORMSN,PAIDNO,TU_ID,Print_Name,Use_For," '表單序號,填表人身份證字號,填表人級職,填表人姓名,用途
                stmt += "ORG_UID,File_Name,Security_Status,Security_Guid,WriteCard," '填表人單位,申請影印類別(1~5),機密等級之申請單流水號,審核是否通過
                stmt += "Ori_sheet,Copy_sheet,Total_sheet)" '原件張數,每份複印張數,合計複印張數
                stmt += " values('" + str_EFORMSN + "','" + do_sql.G_user_id + "'," + tu_id + ",'" + Lab_emp_chinese_name.Text + "','" + tbUse_For.Text.Trim() + "',"
                stmt += org_uid + ",'" + tbFile_Name.Text.Trim() + "'," + ddl_Security_Status.SelectedValue + "," + SecurityItem + "," + IIf(ddl_Security_Status.SelectedValue = "1", "1", "0") + ","
                stmt += txtOriginalSheets.Text + "," + txtCopySheets.Text + "," + txtTotalSheets.Text + ")"

                If do_sql.db_exec(stmt, do_sql.G_conn_string) = False Then
                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('影印掃瞄申請失敗，請重新再試，或連絡系統管理人員！=" + do_sql.G_errmsg + "');")
                    Response.Write(" </script>")
                    Exit Sub
                End If

                '普通件一送件即寫TicketSN入卡久，密等件到審核通過了才寫TicketSN
                If ddl_Security_Status.SelectedValue = "1" Then
                    Dim js As String = "smfWriteTicket('" + str_EFORMSN + "');"
                    ClientScript.RegisterClientScriptBlock(Me.GetType(), "myscript", js, True)
                    Exit Sub
                End If
                '尋找此筆記錄流水號,寫入歷程表-1=新增
                Dim sLog_Guid As String = String.Empty
                db.Open()
                Dim querycnn As New SqlCommand("select Log_Guid from P_08 where EFORMSN='" + str_EFORMSN + "'", db)
                Dim LogGuidr = querycnn.ExecuteReader()
                If LogGuidr.Read() Then
                    sLog_Guid = LogGuidr("Log_Guid")
                End If
                LogGuidr.Close()
                db.Close()
                CP.ActionReWrite(IIf(IsNumeric(sLog_Guid), Integer.Parse(sLog_Guid), 0), user_id, 1, stmt)
            End If
        End If
        '*************** Terminal of If read_only <> "2" ***************

        If read_only = "2" Then
            Dim strAgentName As String = ""
            '找尋批核者姓名
            db.Open()
            Dim strPer As New SqlCommand("SELECT emp_chinese_name FROM EMPLOYEE WHERE employee_id = '" & user_id & "'", db)
            Dim RdPer = strPer.ExecuteReader()
            If RdPer.Read() Then
                strAgentName = RdPer("emp_chinese_name")
            End If
            RdPer.Close()
            db.Close()

            Dim bl_executeResult = False
            Dim conn As SqlConnection = New SqlConnection(connstr)
            Dim myTrans As SqlTransaction
            Try
                conn.Open()
                myTrans = conn.BeginTransaction()
                Dim myCommand As New SqlCommand()
                myCommand.Connection = conn
                myCommand.Transaction = myTrans
                myCommand.CommandText = "update P_08 set AgreeName ='" + strAgentName + "',AgreeORG_UID ='" + org_uid + "',AgreeTime=getdate() where EFORMSN='" + str_EFORMSN + "'"
                myCommand.ExecuteNonQuery()
                executeDBScript = myCommand.CommandText
                myCommand.CommandText = "update P_0804 set Update_Datetime= getdate(),Security_Status = 1 where Guid_ID=" + lbSecurity.Text
                myCommand.ExecuteNonQuery()
                executeDBScript += ";" + myCommand.CommandText
                myTrans.Commit()
                bl_executeResult = True
            Catch ex As Exception
                myTrans.Rollback()
                If Not myTrans.Connection Is Nothing Then
                    Console.WriteLine("An exception of type " & ex.GetType().ToString() & _
                                      " was encountered while attempting to roll back the transaction.")
                End If
            Finally
                conn.Close()
            End Try

            '寫入歷程表 2:審核通過=修改
            CP.ActionReWrite(IIf(String.IsNullOrEmpty(lbLog_Guid.Text), 0, Integer.Parse(lbLog_Guid.Text)), user_id, 2, executeDBScript)

            If bl_executeResult = False Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('審核失敗，請重新再試，或連絡系統管理人員！');")
                Response.Write(" </script>")
                Exit Sub
            ElseIf read_only = "2" Then '若為審核表單,則審核通過後,將母視窗資料重新整理更新
                Response.Write(" <script language='javascript'>")
                Response.Write(" window.dialogArguments.document.location.href = window.dialogArguments.document.location;")
                Response.Write(" </script>")
            End If
        End If

        Dim Val_P As String
        Val_P = ""

        '判斷下一關為上一級主管時人數是否超過一人
        If ((NextPer > 1 And read_only = "") Or (NextPer > 1 And strgroup_id = "Z860")) And ddl_Security_Status.SelectedValue <> "1" Then
            Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & str_EFORMSN & "&SelFlag=1&AppUID=" & DrDown_emp_chinese_name.SelectedValue)
        Else
            '表單審核
            Val_P = FCC.F_Send(SendVal, do_sql.G_conn_string)
            Dim PageUp As String = ""
            If read_only = "" Then
                PageUp = "New"
            End If
            If ddl_Security_Status.SelectedValue <> "1" Then
                Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & PageUp)
            Else
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('新增成功！');")
                Response.Write(" window.parent.parent.location='../../index.aspx';window.close();")
                Response.Write(" </script>")
            End If
        End If
    End Sub

    ''' <summary>
    ''' 取得下一關卡ID
    ''' </summary>
    ''' <param name="db">SqlConnection</param>
    ''' <param name="strEFORMSN">表單ID</param>
    ''' <param name="strUserID">申請人ID</param>
    ''' <returns>下一關卡ID</returns>
    ''' <remarks></remarks>
    Private Function GetNextStepID(ByRef db As SqlConnection, ByVal strEFORMSN As String, ByVal strUserID As String) As Integer
        Dim Chknextstep As Integer
        db.Open()                                 '判斷表單關卡
        Dim strComCheck As New SqlCommand("SELECT nextstep FROM flowctl WHERE eformsn = '" & strEFORMSN & "' and empuid = '" & strUserID & "' and hddate is null", db)
        Dim RdrCheck = strComCheck.ExecuteReader()
        If RdrCheck.Read() Then
            Chknextstep = RdrCheck.Item("nextstep")
        End If
        RdrCheck.Close()
        db.Close()
        Return Chknextstep
    End Function

    Function select_data() As Boolean
        select_data = False
        Dim stmt As String
        Call drow_txt()
        stmt = "select a.Log_Guid,a.status,A.Use_For,d.AD_Title,c.TU_Name,b.ORG_Name,a.LogTime,a.PAIDNO,"
        stmt += "a.Print_Name,a.File_Name,a.Security_Status,a.Security_Guid,isnull(AgreeName,'') as AgreeName,a.WriteCard "
        stmt += "from P_08 a left join ADMINGROUP b on a.ORG_UID=b.ORG_UID "
        stmt += "left join Titles_U c on a.TU_ID = c.TU_ID "
        stmt += "left join employee d on a.PAIDNO=d.employee_id "
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
            str_app_id = ""
        Else
            Lab_ORG_NAME_1.Text = do_sql.G_table.Rows(0).Item("ORG_Name").ToString
            Lab_emp_chinese_name.Text = do_sql.G_table.Rows(0).Item("Print_Name").ToString
            Lab_title_name_1.Text = do_sql.G_table.Rows(0).Item("AD_Title").ToString
            Lab_ORG_NAME_2.Text = do_sql.G_table.Rows(0).Item("ORG_Name").ToString
            str_app_id = do_sql.G_table.Rows(0).Item("PAIDNO").ToString
            Dim nSecurity_Status As String = do_sql.G_table.Rows(0).Item("Security_Status").ToString
            Dim nSecurity_Guid As String = do_sql.G_table.Rows(0).Item("Security_Guid").ToString
            ddl_Security_Status.SelectedValue = do_sql.G_table.Rows(0).Item("Security_Status").ToString
            ddl_Security_Status.Enabled = False
            '-1:查無資料,表示未申請過;0:清除此sn ;1:新申請 ;2:已印但未回登 ;3:印列失敗 ;4:補登完畢 ;5:審核不通過(密等為2~5者才有)
            Dim strStatusTemp As String = do_sql.G_table.Rows(0).Item("Status").ToString '表單狀態
            If nSecurity_Status <> "1" And nSecurity_Guid.Length > 0 Then
                lbSecurity.Text = nSecurity_Guid
                'read_only不為為空,即其為瀏覽表單,非新增表單,雖其為密等以上,亦不顯示Lable機密表單字樣
                If Not String.IsNullOrEmpty(read_only) Then
                    lb_Security.Visible = False
                Else
                    lb_Security.Visible = True
                End If

                'btnApplyForSecurity.Visible = True
                'HLSecurity.Visible = True '20120810 更改取消該檢視功能連結
                lbSecurity.Text = nSecurity_Guid
                'If read_only = "1" Then
                '    HLSecurity.NavigateUrl = "../08/MOA08010.aspx?Security_GuidID=" + nSecurity_Guid + "&Security_Level=" + nSecurity_Status + "&ViewStatusCode=1"
                'End If
                'If read_only = "2" Then
                HLSecurity.NavigateUrl = "../08/MOA08010.aspx?Security_GuidID=" + nSecurity_Guid + "&Security_Level=" + nSecurity_Status
                'End If
                TableSecurityDataSetting(nSecurity_Guid) '設定機密申請表單顯示資料
                If do_sql.G_table.Rows(0).Item("WriteCard").ToString = "0" And user_id = str_app_id Then
                    'btWrite.Visible = True
                End If
            End If

            DrDown_emp_chinese_name.Items.Clear()
            DrDown_emp_chinese_name.Items.Add(Trim(do_sql.G_table.Rows(0).Item("Print_Name").ToString))
            DrDown_emp_chinese_name.Items(0).Value = Trim(do_sql.G_table.Rows(0).Item("Print_Name").ToString)
            tbFile_Name.Text = do_sql.G_table.Rows(0).Item("File_Name").ToString
            tbFile_Name.Enabled = False
            tbUse_For.Text = do_sql.G_table.Rows(0).Item("Use_For").ToString
            tbUse_For.Enabled = False
            Lab_title_name_2.Text = do_sql.G_table.Rows(0).Item("AD_Title").ToString
            Lab_time.Text = CDate(do_sql.G_table.Rows(0).Item("LogTime").ToString).ToString("yyyy/MM/dd HH:mm:ss")
            lbLog_Guid.Text = do_sql.G_table.Rows(0).Item("Log_Guid").ToString

        End If

        select_data = True
    End Function

    Sub disabel_field()
        DrDown_emp_chinese_name.Enabled = False

        If read_only = "1" Then
            but_exe.Visible = False
            tranBtn.Visible = False
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
        p = 0
        p += 1
    End Sub

    Protected Sub backBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles backBtn.Click
        Try
            Dim db As New SqlConnection(connstr)
            Dim strAgentName As String = ""
            '找尋批核者姓名
            db.Open()
            Dim strPer As New SqlCommand("SELECT emp_chinese_name FROM EMPLOYEE WHERE employee_id = '" & user_id & "'", db)
            Dim RdPer = strPer.ExecuteReader()
            If RdPer.Read() Then
                strAgentName = RdPer("emp_chinese_name")
            End If
            RdPer.Close()
            db.Close()
            Dim bl_executeResult = False
            Dim conn As SqlConnection = New SqlConnection(connstr)
            Dim myTrans As SqlTransaction
            Try
                conn.Open()
                myTrans = conn.BeginTransaction()
                Dim myCommand As New SqlCommand()
                myCommand.Connection = conn
                myCommand.Transaction = myTrans
                myCommand.CommandText = "update P_08 set [status]=5,AgreeName ='" + strAgentName + "',AgreeORG_UID ='" + org_uid + "',AgreeTime=getdate() where EFORMSN='" + str_EFORMSN + "'"
                myCommand.ExecuteNonQuery()
                executeDBScript = myCommand.CommandText
                If ddl_Security_Status.SelectedValue <> "1" And lbSecurity.Text <> "" Then
                    myCommand.CommandText = "UPDATE P_0804 SET Security_Status=2,Update_Datetime= getdate() WHERE Guid_ID=" + lbSecurity.Text
                    myCommand.ExecuteNonQuery()
                    executeDBScript += ";" + myCommand.CommandText
                End If

                myTrans.Commit()
                bl_executeResult = True

            Catch ex As Exception
                myTrans.Rollback()
                If Not myTrans.Connection Is Nothing Then
                    Console.WriteLine("An exception of type " & ex.GetType().ToString() & _
                                      " was encountered while attempting to roll back the transaction.")
                End If
            Finally
                conn.Close()
            End Try

            '寫入歷程表內，2:駁回=修改
            CP.ActionReWrite(IIf(String.IsNullOrEmpty(lbLog_Guid.Text), 0, Integer.Parse(lbLog_Guid.Text)), user_id, 2, executeDBScript)

            If bl_executeResult = False Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('駁回失敗，請重新再試，或連絡系統管理人員！');")
                Response.Write(" </script>")
                Exit Sub
            End If

            Dim Val_P As String = ""
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
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('駁回表單失敗：" + ex.Message + "');")
            '重新整理頁面
            Response.Write(" window.dialogArguments.location='../00/MOA00010.aspx';")
            Response.Write(" window.close();")
            Response.Write(" </script>")
        End Try

    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete
        Try
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string
            '開啟連線
            Dim db As New SqlConnection(connstr)
            '單純讀取表單不可送件
            If read_only = "1" Then
                but_exe.Visible = False
                tranBtn.Visible = False
                backBtn.Visible = False

                '讀取表單可送件
            ElseIf read_only = "2" Then
                Dim PerCount As Integer = 0
                '判斷批核者是否為副主管,假如批核者為副主管則不可送件只可呈轉
                Dim ParentFlag As String = ""
                Dim ParentVal = user_id & "," & str_EFORMSN
                Dim FC As New C_FlowSend.C_FlowSend
                ParentFlag = FC.F_NextChief(ParentVal, connstr)
                If ParentFlag = "1" Then
                    '上一級主管多少人
                    db.Open()
                    Dim PerCountCom As New SqlCommand("select count(empuid) as PerCount from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = (select ORG_UID from EMPLOYEE where employee_id ='" & user_id & "'))", db)
                    Dim PerRdv = PerCountCom.ExecuteReader()
                    If PerRdv.Read() Then
                        PerCount = PerRdv("PerCount")
                    End If
                    PerRdv.Close()
                    db.Close()

                    '上一級沒人則呈現送件按鈕及呈轉按鈕
                    If PerCount = 0 Then
                        but_exe.Visible = True
                        tranBtn.Visible = True
                    Else
                        but_exe.Visible = True
                    End If
                End If
                '找出同級單位以下全部單位
                Dim Org_Down As New C_Public
                '沒有批核權限不可執行動作
                If (Org_Down.ApproveAuth(str_EFORMSN, user_id)) = "" Then
                    but_exe.Visible = False
                    tranBtn.Visible = False
                    backBtn.Visible = False
                End If
            Else
                backBtn.Visible = False
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

                stmt = "select * from Employee where ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ") order by emp_chinese_name"
                If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
                    Exit Sub
                End If

                n_table = do_sql.G_table
                P = 0
            Else
                Call disabel_field()
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

    Protected Sub ddl_Security_Status_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddl_Security_Status.SelectedIndexChanged
        lb_Security.Visible = False
        ddl_SecurityItem.Visible = False
        lbSecurity.Visible = False
        btnApplyForSecurity.Visible = False
        HLSecurity.Visible = False
        ddl_SecurityItem.Items.Clear()
        '當user點選了機密分類時，去截取已通過准許的機密申請單之主旨供選取
        If ddl_Security_Status.SelectedValue <> "1" Then
            lb_Security.Visible = True

            Dim dt As DataTable = New DataTable("securitydt")
            Try
                Using conn As New SqlConnection(connstr)
                    Dim da As New SqlDataAdapter("select Guid_ID,Subject from P_0804 where Print_UserID = '" + ViewState("userId").ToString() + "' and Security_Status = 0 and Security_Level=" + ddl_Security_Status.SelectedValue + "and Guid_ID not in (select Guid_ID from P_08 a join P_0804 b on b.Guid_ID=a.Security_Guid)", conn)
                    da.Fill(dt)
                End Using
            Catch ex As Exception
                dt = Nothing
            End Try
            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    ddl_SecurityItem.DataSource = dt
                    ddl_SecurityItem.DataTextField = "Subject"
                    ddl_SecurityItem.DataValueField = "Guid_ID"
                    ddl_SecurityItem.DataBind()
                    ddl_SecurityItem.Visible = True
                    'HLSecurity.Visible = True '20120810 更改取消該檢視功能連結
                    HLSecurity.NavigateUrl = "../08/MOA08010.aspx?Security_GuidID=" + ddl_SecurityItem.SelectedValue + "&Security_Level=" + ddl_Security_Status.SelectedValue
                    TableSecurityDataSetting(ddl_SecurityItem.SelectedValue)
                Else
                    lbSecurity.Visible = True
                    btnApplyForSecurity.Visible = True
                End If
            End If
        End If
    End Sub

    Private Sub TableSecurityDataSetting(guidId As String)
        Dim dt As DataTable = New DataTable()
        Try
            Using conn As New SqlConnection(connstr)
                Dim da As New SqlDataAdapter("select * from P_0804 where Guid_ID = " + guidId, conn)
                da.Fill(dt)
                If dt.Rows.Count > 0 Then
                    'Top1unitName 申請單位全銜
                    lbTop1unitName.Text = IIf(dt.Rows(0).Item(1) Is Nothing, "", dt.Rows(0).Item(1).ToString())
                    'Subject 主旨(頓號)簡由或(資料名稱)
                    lbSubject.Text = IIf(dt.Rows(0).Item(2) Is Nothing, "", dt.Rows(0).Item(2).ToString())
                    'SignDateTime 發文時間
                    lbSignDateTime.Text = IIf(dt.Rows(0).Item(3) Is Nothing, "", Convert.ToDateTime(dt.Rows(0).Item(3)).ToString("yyyy/MM/dd"))
                    'Security_No 發文字號
                    lbSecurity_No.Text = IIf(dt.Rows(0).Item(4) Is Nothing, "", dt.Rows(0).Item(4).ToString())
                    'Security_Level 機密等級 (2:密 3:機密 4:極機密 5:絕對機密)
                    Dim strSecurityLevelNo As String = IIf(dt.Rows(0).Item(5) Is Nothing, "", dt.Rows(0).Item(5).ToString())
                    Select Case strSecurityLevelNo
                        Case "2"
                            lbSecurity_Level.Text = "密"
                        Case "3"
                            lbSecurity_Level.Text = "機密"
                        Case "4"
                            lbSecurity_Level.Text = "極機密"
                        Case "5"
                            lbSecurity_Level.Text = "絕對機密"
                        Case Else
                            lbSecurity_Level.Text = ""
                    End Select
                    'Security_Type 機密屬性 (1:國家機密 2:軍事機密 3:國防秘密 4:國家機密亦屬軍事機密 5:國家機密亦屬國防秘密)
                    Dim strSecurityTypeNo As String = IIf(dt.Rows(0).Item(6) Is Nothing, "", dt.Rows(0).Item(6).ToString())
                    Select Case strSecurityTypeNo
                        Case "0"
                            lbSecurity_Type.Text = "一般公務機密"
                        Case "1"
                            lbSecurity_Type.Text = "國家機密"
                        Case "2"
                            lbSecurity_Type.Text = "軍事機密"
                        Case "3"
                            lbSecurity_Type.Text = "國防秘密"
                        Case "4"
                            lbSecurity_Type.Text = "國家機密亦屬軍事機密"
                        Case "5"
                            lbSecurity_Type.Text = "國家機密亦屬國防秘密"
                        Case Else
                            lbSecurity_Type.Text = ""
                    End Select
                    'Security_Range 保密期限/保密條件
                    lbSecurity_Range.Text = IIf(dt.Rows(0).Item(7) Is Nothing, "", dt.Rows(0).Item(7).ToString())
                    'ProduceUnit 產製單位 
                    lbProduceUnit.Text = IIf(dt.Rows(0).Item(22) Is Nothing, "", dt.Rows(0).Item(22).ToString())
                    'AgreeTimeOrNumber 同意複(影)印時間/文號
                    lbAgreeTimeOrNumber.Text = IIf(dt.Rows(0).Item(23) Is Nothing, "", dt.Rows(0).Item(23).ToString())
                    'AgreeSuperior 同意複(影)印權責長官級職姓名
                    lbAgreeSuperior.Text = IIf(dt.Rows(0).Item(24) Is Nothing, "", dt.Rows(0).Item(24).ToString())
                    'Purpose 用途 (1:呈閱 2:分會、辦 3:作業用 4:歸檔 5:隨文分發 6:會議分發 7:(其他)Purpose_Other)
                    Dim strPurposeNo As String = IIf(dt.Rows(0).Item(8) Is Nothing, "", dt.Rows(0).Item(8).ToString())
                    Select Case strPurposeNo
                        Case "1"
                            lbPurpose.Text = "呈閱"
                        Case "2"
                            lbPurpose.Text = "分會、辦"
                        Case "3"
                            lbPurpose.Text = "作業用"
                        Case "4"
                            lbPurpose.Text = "歸檔"
                        Case "5"
                            lbPurpose.Text = "隨文分發"
                        Case "6"
                            lbPurpose.Text = "會議分發"
                        Case "7"
                            lbPurpose.Text = "其他(" + IIf(dt.Rows(0).Item(9) Is Nothing, "", dt.Rows(0).Item(9).ToString()) + ")"
                        Case Else
                            lbPurpose.Text = ""
                    End Select
                    'Printer_Datetime 複印時間
                    lbPrinter_Datetime.Text = IIf(dt.Rows(0).Item(17) Is Nothing, "", Convert.ToDateTime(dt.Rows(0).Item(17)).ToString("yyyy/MM/dd"))
                    'Printer_Num 複(影)印機浮水印暗記編號
                    lbPrinter_Num.Text = IIf(dt.Rows(0).Item(10) Is Nothing, "", dt.Rows(0).Item(10).ToString())
                    'Ori_sheet 原件張數
                    lbOri_sheet.Text = IIf(dt.Rows(0).Item(18) Is Nothing, "", dt.Rows(0).Item(18).ToString())
                    'Copy_sheet 每份複印張數
                    lbCopy_sheet.Text = IIf(dt.Rows(0).Item(19) Is Nothing, "", dt.Rows(0).Item(19).ToString())
                    'Total_sheet 合計複印張數
                    lbTotal_sheet.Text = IIf(dt.Rows(0).Item(20) Is Nothing, "", dt.Rows(0).Item(20).ToString())
                    'Sheet_ID 複(影)印張數流水號
                    lbSheet_ID.Text = IIf(dt.Rows(0).Item(21) Is Nothing, "", dt.Rows(0).Item(21).ToString())
                    'Memo 附註
                    lbMemo.Text = IIf(dt.Rows(0).Item(15) Is Nothing, "", dt.Rows(0).Item(15).ToString())
                End If
            End Using
        Catch ex As Exception
            dt = Nothing
        End Try
    End Sub

    Protected Sub bt_clear_Click(sender As Object, e As System.EventArgs) Handles bt_clear.Click
        Dim strErrMsg As String = String.Empty
        'Try
        '    Dim SmfCom As SMFATLLib.SmfCom = New SMFATLLib.SmfCom
        '    Dim iSinit As Integer = SmfCom.SmfInit()
        '    If iSinit <> 0 Then
        '        strErrMsg = "讀取卡片失敗!!"
        '    Else
        '        Dim iCheck As Integer = SmfCom.SmfCleanTicket()
        '        If iCheck <> 0 Then
        '            strErrMsg = "清除卡片內Ticket失敗:" + iCheck.ToString()
        '        End If
        '    End If
        'Catch ex As Exception
        '    strErrMsg = "清除卡片內Ticket出現錯誤:" + ex.Message
        'End Try

        'If strErrMsg <> String.Empty Then
        '    Response.Write(" <script language='javascript'>")
        '    Response.Write(" alert('" + strErrMsg + "');")
        '    Response.Write(" </script>")
        '    Exit Sub
        'End If

        If str_EFORMSN <> "" And user_id <> "" Then
            Dim dt As DataTable = New DataTable("securitydt")
            Try
                Using Clearconn As New SqlConnection(connstr)
                    Dim da As New SqlDataAdapter("select top 1 * from P_08 where PAIDNO = '" + user_id + "' and Status = 1 order by LogTime Desc", Clearconn)
                    da.Fill(dt)
                End Using
            Catch ex As Exception
                dt = Nothing
            End Try

            If Not dt Is Nothing Then
                Dim bl_executeResult = False
                Dim sLog_Guid As String = dt.Rows(0).Item("Log_Guid").ToString()
                Dim Security_Guid As String = dt.Rows(0).Item("Security_Guid").ToString()
                Dim clearEFORMSN As String = dt.Rows(0).Item("EFORMSN").ToString()
                Dim conn As SqlConnection = New SqlConnection(connstr)
                Dim myTrans As SqlTransaction
                Try
                    conn.Open()
                    myTrans = conn.BeginTransaction()
                    Dim myCommand As New SqlCommand()
                    myCommand.Connection = conn
                    myCommand.Transaction = myTrans
                    '0:清除影印申請 1:申請完成未列印 2:已列印未回登資料 3:印列失敗 4:補登完畢 5:審核不通過 6:申請人取消
                    myCommand.CommandText = "update P_08 set Status = 6,UPdate_Date=getdate() where EFORMSN='" + clearEFORMSN + "'"
                    myCommand.ExecuteNonQuery()
                    executeDBScript = myCommand.CommandText
                    If Security_Guid.Length <> 0 Then   '如果本來申請單是密等級以上的，則清除該筆申請，其當初選的機密文件也會列為審核不通過，一併重新申請
                        '0:審核中或未送審 1:審核通過 2:審核不通過 3:申請人取消
                        myCommand.CommandText = "update P_0804 set Update_Datetime= getdate(),Security_Status = 3 where Guid_ID=" + Security_Guid
                        myCommand.ExecuteNonQuery()
                        executeDBScript += " " + myCommand.CommandText
                    End If

                    myTrans.Commit()
                    bl_executeResult = True

                Catch ex As Exception
                    myTrans.Rollback()
                    If Not myTrans.Connection Is Nothing Then
                        Console.WriteLine("An exception of type " & ex.GetType().ToString() & _
                                          " was encountered while attempting to roll back the transaction.")
                    End If
                Finally
                    conn.Close()
                End Try

                '寫入歷程表 2:清除記錄狀態=修改
                CP.ActionReWrite(IIf(IsNumeric(sLog_Guid), Integer.Parse(sLog_Guid), 0), user_id, 2, executeDBScript)

                If bl_executeResult = False Then
                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('清除已申請記錄失敗，請重新再試，或連絡系統管理人員！');")
                    Response.Write(" </script>")
                    Exit Sub
                Else
                    bt_clear.Visible = False
                    Response.Write(" <script language='javascript'>")
                    Response.Write("parent.location.reload();")
                    Response.Write(" alert('已清除申請中記錄，可重新做影印申請！');")
                    Response.Write(" </script>")
                End If
            Else
                bt_clear.Visible = False
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('您目前沒有已申請中的記錄!!');")
                Response.Write(" </script>")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub btWrite_Click(sender As Object, e As System.EventArgs) Handles btWrite.Click
        '密等級以上之申請通過上一級主管審核通過後，原申請人可透過此按鈕將TicketSN寫入卡片才能拿卡去printer做列印
        'Dim SmfCom As SMFATLLib.SmfCom = New SMFATLLib.SmfCom
        'Dim iSinit As Integer = SmfCom.SmfInit()
        'If iSinit <> 0 Then
        '    Response.Write(" <script language='javascript'>")
        '    Response.Write(" alert('Card Error!!');")
        '    Response.Write(" </script>")
        'Else
        'Dim iWriteResult = SmfCom.SmfWriteTicket(2, str_EFORMSN, 1, 1, 1)

        Dim js As String = "verifyWriteTicket('" + str_EFORMSN + "');"
        ClientScript.RegisterClientScriptBlock(Me.GetType(), "myscript", js, True)

        'If iWriteResult <> 0 Then
        '    Response.Write(" <script language='javascript'>")
        '    Response.Write(" alert('寫入卡片失敗，請重新再試!!');")
        '    Response.Write(" </script>")
        'Else
        '    Dim sUpdate As String = "Update P_08 set WriteCard = 1 where EFORMSN='" + str_EFORMSN + "'"
        '    If do_sql.db_sql(sUpdate, do_sql.G_conn_string) = True Then
        '        Exit Sub
        '    End If

        '    Response.Write(" <script language='javascript'>")
        '    Response.Write(" alert('完成申請已寫入卡片，您可持卡進行影印作業！');")
        '    '重新整理頁面
        '    Response.Write(" window.dialogArguments.location='../00/MOA00010.aspx';")
        '    Response.Write(" window.close();")
        '    Response.Write(" </script>")
        'End If
        'End If
    End Sub

    Protected Sub Button1_Click(sender As Object, e As System.EventArgs) Handles Button1.Click
        'Try
        '    Dim SmfCom As SMFATLLib.SmfCom = New SMFATLLib.SmfCom
        '    Dim iSinit As Integer = SmfCom.SmfInit()
        '    If iSinit <> 0 Then
        '        Response.Write(" <script language='javascript'>")
        '        Response.Write(" alert('Card Fail!!');")
        '        Response.Write(" </script>")
        '    Else
        '        Dim iCheck As Integer = SmfCom.SmfReadTicket()
        '        If iCheck = 0 Then
        '            Response.Write(" <script language='javascript'>")
        '            Response.Write(" alert('Sn:" + SmfCom.TicketSn + "');")
        '            Response.Write(" </script>")
        '        Else
        '            Response.Write(" <script language='javascript'>")
        '            Response.Write(" alert('SmfReadTicket Fail:" + iCheck.ToString() + "');")
        '            Response.Write(" </script>")
        '        End If
        '    End If
        'Catch ex As Exception
        '    Response.Write(" <script language='javascript'>")
        '    Response.Write(" alert('SmfReadTicket Error:" + ex.Message + "');")
        '    Response.Write(" </script>")
        'End Try

    End Sub


    Protected Sub btnApplyForSecurity_Click(sender As Object, e As System.EventArgs) Handles btnApplyForSecurity.Click
        Server.Transfer("MOA08008.aspx")
    End Sub

    Protected Sub tranBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tranBtn.Click
        Try
            If GetSuperiors(user_id).Rows.Count > 1 Then '若上一級主管為兩人以上,則切換至選擇欲呈轉主管之頁面
                Server.Transfer("../00/MOA00014.aspx?eformsn=" + str_EFORMSN)
            Else
                'If read_only = "2" Then
                '    '增加批核意見
                '    insComment(txtcomment.Text, str_EFORMSN, user_id)
                'End If
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
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('呈轉失敗:" + ex.Message + "');")
            Response.Write(" </script>")
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

End Class
