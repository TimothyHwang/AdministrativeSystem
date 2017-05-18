Imports System.Data.SqlClient
Imports System.Data
Imports WebUtilities.Functions

Partial Class M_Source_08_MOA08002
    Inherits System.Web.UI.Page
    Dim user_id, org_uid, sTU_ID, sLog_Guid, sLogStatus As String
    Dim bl_Check As Boolean = False
    Dim iPrintTatol As Int32
    Dim CCFun As New C_CheckFun
    Dim CF As New CFlowSend
    Dim CP As New C_Public
    Dim sql_function As New C_SQLFUN

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        sTU_ID = Session("TU_ID")
        'session被清空回首頁
        If user_id = "" Or org_uid = "" Or sTU_ID = "" Then
            bl_Check = True
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")
        Else
            '判斷登入者權限
            If CP.LoginCheck(user_id, "MOA08003") <> "" Then
                CP.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA08003.aspx")
                Response.End()
            End If
            Dim sourceUrl As String() = Request.UrlReferrer.ToString().Split("/") '取得來源頁面,返回上一頁按鈕使用
            Dim sourcePage As String = sourceUrl(sourceUrl.Length - 1)
            ViewState("sourcePage") = sourcePage
            '判斷導址值是否正確 Status=1:未印(申請中) 2:已印(可以補登) 3:印列失敗(跳出) 4:補登完畢(僅可檢視) 0:清除此sn(跳出)
            'sType = IIf(IsNothing(Request.QueryString("Type")), "", Request.QueryString("Type"))
            sLog_Guid = IIf(IsNothing(Request.QueryString("Log_Guid")), "", Request.QueryString("Log_Guid"))
            sLogStatus = IIf(IsNothing(Request.QueryString("Status")), "", Request.QueryString("Status"))
            If (sLogStatus.Trim().Length < 1 Or (Not CCFun.isNumeric(sLogStatus)) Or sLog_Guid.Trim().Length < 1 Or (Not CCFun.isNumeric(sLog_Guid))) Then
                bl_Check = True
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('影(複)印表單序號錯誤，請重新操作！');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")
            End If
        End If
        lbMsg.Text = ""
        If bl_Check = False And Not IsPostBack Then
           
            '查詢截取單一詳細資料
            Dim sQueryPrintLog As String = String.Empty
            Dim Printerdt As DataTable = DataRecords_DataRebinding(sLogStatus, sLog_Guid, sQueryPrintLog)
            If Not Printerdt Is Nothing And Printerdt.Rows.Count > 0 Then
                '一進入代表查詢即要Insert歷程入DB Table:[P_0802]-movement=3 查看 尚未動作(1:新增 2:修改 3:查看 4.另存新檔/列印)
                CP.ActionReWrite(IIf(IsNumeric(sLog_Guid), Integer.Parse(sLog_Guid), 0), user_id, 3, sQueryPrintLog)
                Dim PAIDNO As String = Printerdt.Rows(0)("PAIDNO").ToString()
                If sLogStatus = 2 And user_id = PAIDNO Then '1:未印 2:已印 3:印列失敗 4:補登完畢 5:申請單駁回 6:申請人取消 0:清除此sn
                    ibtnModify.Visible = True
                ElseIf sLogStatus = 4 Or sLogStatus = 5 Or sLogStatus = 6 Or user_id <> PAIDNO Then
                    ibtnModify.Visible = False
                End If

                lb_Log_Guid.Text = sLog_Guid
                lb_Print_Date.Text = Printerdt.Rows(0)("Print_Date").ToString()
                lb_Log_Time.Text = Printerdt.Rows(0)("LogTime").ToString()
                lb_Print1.Text = ShowPrint(Printerdt)(0)
                lb_Print2.Text = ShowPrint(Printerdt)(1)
                lb_Print3.Text = ShowPrint(Printerdt)(2)
                lb_PAIDNO.Text = Printerdt.Rows(0)("PAIDNO").ToString()
                Dim da_accountinfo As DataTable
                da_accountinfo = CF.F_AdminEmp(Printerdt.Rows(0)("PAIDNO").ToString(), sql_function.G_conn_string)
                If Not da_accountinfo Is Nothing And da_accountinfo.Rows.Count > 0 Then
                    lb_Print_Name.Text = da_accountinfo.Rows(0).Item("emp_chinese_name")
                    lb_ORG_NAME.Text = da_accountinfo.Rows(0).Item("ORG_NAME")
                    Dim sTU_Name As String = CF.getTU_Name(Printerdt.Rows(0)("TU_ID").ToString(), sql_function.G_conn_string)
                    If sTU_Name.Length > 0 And sTU_Name <> "error" Then
                        lb_TU_ID_Name.Text = sTU_Name
                    End If
                End If
                lb_PrintTotalCnt.Text = Printerdt.Rows(0)("PrintTotalCnt").ToString()
                lb_Print_Num.Text = Printerdt.Rows(0)("Print_Num").ToString()
                lb_File_Name.Text = Printerdt.Rows(0)("File_Name").ToString()
                lbOri_sheet.Text = IIf(IsDBNull(Printerdt.Rows(0)("Ori_sheet")), "", Printerdt.Rows(0)("Ori_sheet").ToString())
                lbCopy_sheet.Text = IIf(IsDBNull(Printerdt.Rows(0)("Copy_sheet")), "", Printerdt.Rows(0)("Copy_sheet").ToString())
                lbTotal_sheet.Text = IIf(IsDBNull(Printerdt.Rows(0)("Total_sheet")), "", Printerdt.Rows(0)("Total_sheet").ToString())
                tb_Useless.Text = Printerdt.Rows(0)("Useless").ToString()
                lbVerifyRequester.Text = IIf(IsDBNull(Printerdt.Rows(0)("VerifyRequester")), "", Printerdt.Rows(0)("VerifyRequester").ToString())
                lbApprovedBy.Text = IIf(IsDBNull(Printerdt.Rows(0)("ApprovedBy")), "", Printerdt.Rows(0)("ApprovedBy").ToString())
                tb_Use_For.Text = Printerdt.Rows(0)("Use_For").ToString()
                tb_memo.Text = Printerdt.Rows(0)("memo").ToString()

                '為了不影響Print_Num登記流水號，列印份數無法被修改,修改作廢張數也不影響登記流水號
                lb_Print_Num.Text = Printerdt.Rows(0)("Print_Num").ToString()
                Dim sChoseSecurity_Status As String = Printerdt.Rows(0)("Security_Status").ToString()
                ddl_Security_Status.SelectedIndex = Integer.Parse(sChoseSecurity_Status) - 1
                If (Printerdt.Rows(0)("Security_Status").ToString() <> "1") Then
                    HLSecurity.Visible = True
                    HLSecurity.NavigateUrl = "../08/MOA08010.aspx?Security_GuidID=" + Printerdt.Rows(0)("Security_Guid").ToString()
                End If
                If sLogStatus = "4" Or user_id <> Printerdt.Rows(0)("PAIDNO").ToString() Then '已補登完畢本人無法再做修改，非本人只能檢視
                    tb_Useless.Enabled = False
                    tb_memo.Enabled = False
                End If
            Else
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('該影(複)印資料表單查詢失敗，請重新操作！');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")
            End If
        End If
    End Sub
    '組合5種列印張數並回傳顯示給前台的字串
    Private Function ShowPrint(ByVal Printerdt As DataTable) As String()
        Dim _showPrint(3) As String
        Dim sA3C, sA4C, sA3M, sA4M, sScan As String
        sA3C = IIf(IsDBNull(Printerdt.Rows(0)("Copy_A3C")), "0", Printerdt.Rows(0)("Copy_A3C").ToString())
        sA4C = IIf(IsDBNull(Printerdt.Rows(0)("Copy_A4C")), "0", Printerdt.Rows(0)("Copy_A4C").ToString())
        sA3M = IIf(IsDBNull(Printerdt.Rows(0)("Copy_A3M")), "0", Printerdt.Rows(0)("Copy_A3M").ToString())
        sA4M = IIf(IsDBNull(Printerdt.Rows(0)("Copy_A4M")), "0", Printerdt.Rows(0)("Copy_A4M").ToString())
        sScan = IIf(IsDBNull(Printerdt.Rows(0)("SCan")), "0", Printerdt.Rows(0)("SCan").ToString())

        If Not CCFun.isNumeric(sA3C) Or Not CCFun.isNumeric(sA4C) Or Not CCFun.isNumeric(sA3M) Or Not CCFun.isNumeric(sA4M) Or Not CCFun.isNumeric(sScan) Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('該影(複)印資料表單查詢張數有錯誤，請重新操作或聯絡資訊人員！');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")
        Else
            Dim ayPrintTypeCnt As String() = {sA3M, sA4M, sA4C, sA3C, sScan}
            Dim ayPrintTypeName As String() = {"A3黑白", "A4黑白", "A4彩色", "A3彩色", "掃&nbsp;&nbsp;&nbsp;&nbsp;瞄"}
            For i As Int16 = 0 To ayPrintTypeCnt.Length - 1 Step 1
                If Int16.Parse(ayPrintTypeCnt(i).ToString()) <> 0 Then
                    If i = 0 Or i = 1 Then
                        _showPrint(0) += ayPrintTypeName(i).ToString() + ": " + ayPrintTypeCnt(i).ToString() + "&nbsp;張&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                    End If
                    If i = 2 Or i = 3 Then
                        _showPrint(1) += ayPrintTypeName(i).ToString() + ": " + ayPrintTypeCnt(i).ToString() + "&nbsp;張&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                    End If
                    If i = 4 Then
                        _showPrint(2) += ayPrintTypeName(i).ToString() + ": " + ayPrintTypeCnt(i).ToString() + "&nbsp;張"
                    End If
                    iPrintTatol += Int16.Parse(ayPrintTypeCnt(i).ToString())
                End If
            Next i
        End If

        'If iPrintTatol = 0 Then
        '    Response.Write(" <script language='javascript'>")
        '    Response.Write(" alert('該影(複)印資料表單查詢張數有錯誤(0)，請重新操作或聯絡資訊人員！');")
        '    Response.Write(" window.parent.location='../../index.aspx';")
        '    Response.Write(" </script>")
        'End If
        lb_Org_Num.Text = iPrintTatol.ToString()
        Return _showPrint
    End Function

    '查詢目前所有列印總筆數，組合登記流水號 ex:(掃)1~10
    Private Sub CountPrint_Num(ByVal iScanStatus As Integer)
        Dim CountPrint_Num As Integer = 0
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        Dim sQuery As String = "select isnull(sum(PrintTotalCnt),0) as PrintTotalCnt from dbo.P_0801 with(nolock)"

        command.CommandType = CommandType.Text
        command.CommandText = sQuery
        command.Parameters.Add(New SqlParameter("Log_Guid", SqlDbType.NVarChar, 6)).Value = sLog_Guid
        Try
            command.Connection.Open()
            Dim obj As Object = command.ExecuteScalar()
            If Not obj Is Nothing Then
                CountPrint_Num = Integer.Parse(obj.ToString())
            End If
        Catch ex As Exception
            CountPrint_Num = -1
            lbMsg.Text = "查詢目前所有列印總筆數錯誤:" + ex.Message
        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
            command.Dispose()
            command = Nothing
        End Try

        If (CountPrint_Num >= 0) Then
            iPrintTatol = Integer.Parse(lb_PrintTotalCnt.Text)
        Else
            lbMsg.Text = "查詢目前所有列印總筆數錯誤:" + CountPrint_Num.ToString()
            Return
        End If

        Dim iStartNum As Integer = CountPrint_Num + 1
        Dim iEndNum As Integer = CountPrint_Num + iPrintTatol
        If iScanStatus > 0 Then
            lb_Print_Num.Text = "(掃)" & iStartNum.ToString() & " ~ " & iEndNum.ToString()
        Else
            lb_Print_Num.Text = iStartNum.ToString() & " ~ " & iEndNum.ToString()
        End If
    End Sub

    '查詢該筆原始列印資料
    Private Function DataRecords_DataRebinding(ByVal ActionType As String, ByVal sLog_Guid As String, ByRef sQueryPrintLog As String) As DataTable
        DataRecords_DataRebinding = New DataTable("Detail")
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        'lblTitle.Text = "影印使用登記-[檢視]"
        lblTitle.Text = "影印使用申請-[檢視]"
        sQueryPrintLog = "select P.*, V.emp_chinese_name AS VerifyRequester, A.emp_chinese_name AS ApprovedBy from P_08 AS P with(nolock)"
        sQueryPrintLog = sQueryPrintLog + " LEFT JOIN EMPLOYEE AS V ON P.VerifyRequesterID = V.employee_id"
        sQueryPrintLog = sQueryPrintLog + " LEFT JOIN EMPLOYEE AS A ON P.ApprovedByID = A.employee_id where P.Log_Guid = @Log_Guid"
        If (ActionType = 2) Then
            'lblTitle.Text = "影印使用登記-[補登資料]"
            lblTitle.Text = "影印使用申請-[補登資料]"
        End If
        command.CommandType = CommandType.Text
        command.CommandText = sQueryPrintLog
        command.Parameters.Add(New SqlParameter("Log_Guid", SqlDbType.NVarChar, 6)).Value = sLog_Guid
        Try
            command.Connection.Open()
            Dim dr As SqlDataReader = command.ExecuteReader()
            DataRecords_DataRebinding.Load(dr)

        Catch ex As Exception
            DataRecords_DataRebinding = Nothing
        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
            command.Dispose()
            command = Nothing
        End Try
    End Function

    Protected Sub ibtnPrevious_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnPrevious.Click
        Dim sourcePage As String = ViewState("sourcePage")
        Dim parameters As New SortedList : With parameters
            .Add("Action", sourcePage)
        End With
        Response.Write(Utilities.SubmitFormGeneration(parameters))
        Response.End()
    End Sub

    Protected Sub ddl_Security_Status_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddl_Security_Status.SelectedIndexChanged
        If ddl_Security_Status.SelectedValue <> "1" Then
            HLSecurity.Visible = True
        Else
            HLSecurity.Visible = False
        End If
    End Sub

    '當作廢張數的改變不影鄉總列印張數
    Protected Sub tb_Useless_TextChanged(sender As Object, e As System.EventArgs) Handles tb_Useless.TextChanged
        bl_Check = False
        Dim sUseless As String = tb_Useless.Text.Trim()
        If Not CCFun.isNumeric(sUseless) Then
            lbMsg.Text = "作廢張數請輸入正整數字!"
        Else
            If Int16.Parse(sUseless) < 0 Then
                lbMsg.Text = "作廢張數請輸入>=0的正整數字!"
            Else
                iPrintTatol = Integer.Parse(lb_PrintTotalCnt.Text.Trim())
                If Int16.Parse(sUseless) > iPrintTatol Then
                    lbMsg.Text = "作廢張數無法大於總列印張數!"
                Else
                    bl_Check = True
                End If
            End If
        End If

        If bl_Check = False Then
            ibtnModify.Enabled = False
        Else
            ibtnModify.Enabled = True
        End If
    End Sub

    ''當列印份數一改變則變動總列印張數
    'Protected Sub tb_PrintSet_Cnt_TextChanged(sender As Object, e As System.EventArgs) Handles tb_PrintSet_Cnt.TextChanged
    '    bl_Check = False
    '    Dim sPrintSet_Cnt As String = tb_PrintSet_Cnt.Text.Trim()
    '    If Not CCFun.isNumeric(sPrintSet_Cnt) Then
    '        lbMsg.Text = "列印份數請輸入正整數字!"
    '    Else
    '        If Int16.Parse(sPrintSet_Cnt) < 1 Then
    '            lbMsg.Text = "列印份數請輸入>0的正整數字!"
    '        Else
    '            iPrintTatol = (Int16.Parse(lb_Org_Num.Text) * Int16.Parse(sPrintSet_Cnt))
    '            If iPrintTatol >= 0 Then '有可能印一張，那一張就剛好作廢，所以有可能總列印張數=0
    '                lb_PrintTotalCnt.Text = iPrintTatol
    '                bl_Check = True
    '            Else
    '                lbMsg.Text = "計算總列印張數錯誤，請重新再試或聯絡資訊人員！" + iPrintTatol.ToString()
    '            End If
    '        End If
    '    End If

    '    If bl_Check = False Then
    '        ibtnAdd.Enabled = False
    '        ibtnModify.Enabled = False
    '    Else
    '        ibtnAdd.Enabled = True
    '        ibtnModify.Enabled = True
    '    End If
    'End Sub

    Protected Sub ibtnModify_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ibtnModify.Click
        DBAction()
    End Sub

    '按鈕動作:更新登記記錄
    Private Sub DBAction()
        Dim bl_checkData As Boolean = True

        '驗証複印[作廢張數]
        If tb_Useless.Text.Length = 0 Or Not CCFun.isNumeric(tb_Useless.Text.Trim()) Then
            lbMsg.Text = "請再次確認您的[作廢張數]資料是否正確~"
            bl_checkData = False
        End If

        ''驗証極機密或絕對機密的[申請單流水號]
        'Dim iSecurity_Guid As Integer = 0
        'If ddl_Security_Status.SelectedValue <> "1" And tb_Security_Guid.Visible = True Then
        '    If tb_Security_Guid.Text.Trim() = "" Then
        '        lbMsg.Text = "普通等級以上的列印文件請先查詢您事前申請單建立的流水號回來這兒回填哦~"
        '        bl_checkData = False
        '    ElseIf Not CCFun.isNumeric(tb_Security_Guid.Text.Trim) Then
        '        lbMsg.Text = "普通等級以上之文件申請單建立的流水號必需是正整數字!!~"
        '        bl_checkData = False
        '    Else
        '        iSecurity_Guid = Integer.Parse(tb_Security_Guid.Text.Trim)
        '        If (iSecurity_Guid <= 0) Then
        '            lbMsg.Text = "普通等級以上之文件需填寫的申請單流水號必需為>0的正整數哦~"
        '            bl_checkData = False
        '        End If
        '    End If
        'End If

        '處理[登記流水號]
        Dim sPrint_Num As String = lb_Print_Num.Text.Trim()
        If sPrint_Num.Length = 0 Then
            lbMsg.Text = "[登記流水號]有誤，請重新操作!!~"
            bl_checkData = False
        End If

        Dim ComDB As New SqlCommand
        ComDB.Connection = New SqlConnection(sql_function.G_conn_string)
        If bl_checkData Then 'Update Table P_08
            Try
                Dim sUpdateStr As String = "update P_08 set Useless=@Useless,Print_Num=@Print_Num,"
                sUpdateStr += "Memo=@Memo,Update_Date=getdate(),[Status]=4 where Log_Guid = @Log_Guid and [Status]=2"
                ComDB.CommandText = sUpdateStr
                ComDB.Parameters.Clear()
                ComDB.Parameters.Add(New SqlParameter("@Print_Num", SqlDbType.NVarChar, 30)).Value = sPrint_Num
                ComDB.Parameters.Add(New SqlParameter("@Useless", SqlDbType.Int)).Value = Int32.Parse(tb_Useless.Text.Trim())
                ComDB.Parameters.Add(New SqlParameter("@Memo", SqlDbType.NVarChar, 255)).Value = tb_memo.Text.Trim()
                ComDB.Parameters.Add(New SqlParameter("@Log_Guid", SqlDbType.Int)).Value = Integer.Parse(lb_Log_Guid.Text.Trim())
                ComDB.Connection.Open()
                ComDB.ExecuteNonQuery()
                ComDB.Connection.Close()
                ComDB.Dispose()
                ComDB = Nothing
                '寫入歷程 2:補登列印資料=修改
                CP.ActionReWrite(Integer.Parse(lb_Log_Guid.Text.Trim()), user_id, 2, sUpdateStr)
            Catch ex As Exception
                bl_checkData = False
                lbMsg.Text = "更新P_08登記資料失敗，請重新操作或聯絡資訊人員!!!!" + ex.Message
            End Try
            If bl_checkData Then
                CCFun.AlertSussTranscation("補登影印資料成功!!", "MOA08003.aspx")
            End If
        End If
    End Sub

    'Protected Sub tb_Security_Guid_TextChanged(sender As Object, e As System.EventArgs) Handles tb_Security_Guid.TextChanged
    '    If (tb_Security_Guid.Text.Trim <> "") Then
    '        '判斷輸入的機密資訊申請單是否為同一人所寫所印by PAIDNO
    '        Dim Security_DateTime As String = String.Empty
    '        Dim Security_Level As Integer = 0
    '        Dim bl_CheckPerson As Boolean = QuerySecuritybyPID(Int32.Parse(tb_Security_Guid.Text.Trim()), Security_Level, Security_DateTime)
    '        If (bl_CheckPerson <> True) Then
    '            lbMsg.Text = "此份機密資料非您本人身份証ID填寫，請由原送單人補登!!"
    '            tb_Security_Guid.Text = ""
    '        End If

    '        If (Security_Level <> 0) Then
    '            Dim choseSecurity_Level As Integer = Integer.Parse(ddl_Security_Status.SelectedValue)
    '            If (choseSecurity_Level <> Security_Level) Then
    '                lbMsg.Text = "此份機密資料原申請機密等級與您現行所選不符!"
    '                tb_Security_Guid.Text = ""
    '            End If
    '        End If

    '        If (Security_DateTime <> String.Empty) Then
    '            Try
    '                Dim dt_Security_DateTime As DateTime = DateTime.Parse(Security_DateTime)
    '                Dim dt_Print_Date As DateTime = DateTime.Parse(lb_Print_Date.Text.Trim())
    '                If (dt_Security_DateTime > dt_Print_Date) Then
    '                    lbMsg.Text = "機密資料原申請日期時間晚於您欲回登的此筆資料列印時間!!"
    '                    tb_Security_Guid.Text = ""
    '                End If
    '            Catch ex As Exception
    '                lbMsg.Text = "查詢此份機密資料原申請日期時間有誤!!"
    '                tb_Security_Guid.Text = ""
    '            End Try
    '        End If

    '    End If
    'End Sub

    '查詢該份機密資訊申請單是否為同一人所寫所印by PID
    Private Function QuerySecuritybyPID(ByVal Security_Guid As Integer, ByRef Security_Level As Integer, ByRef Security_DateTime As String) As Boolean
        QuerySecuritybyPID = False
        Security_DateTime = String.Empty
        Security_Level = 0
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        'lblTitle.Text = "影印使用登記-[補登影印資料]"
        lblTitle.Text = "影印使用申請-[補登影印資料]"
        Dim sQueryString As String = "select Security_Level,Security_DateTime from P_0804 with(nolock) where Print_UserID = @Print_UserID and Guid_ID=@Security_Guid"

        command.CommandType = CommandType.Text
        command.CommandText = sQueryString
        command.Parameters.Add(New SqlParameter("Print_UserID", SqlDbType.VarChar, 10)).Value = user_id
        command.Parameters.Add(New SqlParameter("Security_Guid", SqlDbType.Int)).Value = Security_Guid
        Try
            command.Connection.Open()
            Dim dr As SqlDataReader = command.ExecuteReader()
            If (dr.HasRows) Then
                While (dr.Read)
                    QuerySecuritybyPID = True
                    Security_Level = Integer.Parse(dr.GetValue(0).ToString())
                    Security_DateTime = dr.GetValue(1).ToString()
                End While
            End If
            dr.Close()
        Catch ex As Exception
            QuerySecuritybyPID = False
        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
            command.Dispose()
            command = Nothing
        End Try
    End Function
End Class
