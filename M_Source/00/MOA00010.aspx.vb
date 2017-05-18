Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.Common
Imports System.Web.UI
Imports System.Collections.Generic

Partial Class Source_00_MOA00010
    Inherits Page

    Public Shared SelAll As String      '是否全選
    Dim conn As New C_SQLFUN
    Dim connstr As String = conn.G_conn_string
    Dim db As New SqlConnection(connstr)
    Dim user_id, org_uid As String
    'Dim printRecordsReportID As String = GetEFormId("影印紀錄呈核單") '取得影印紀錄呈核單ID
    Dim doorAndMeetingControlID As String = GetEFormId("門禁會議管制申請單") '取得門禁會議管制申請單ID

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As EventArgs) Handles Me.PreRenderComplete
        If Not IsPostBack Then
            Dim tLItm As New ListItem("請選擇", "")
            DropDownList1.Items.Insert(0, tLItm)
            DropDownList1.Items(1).Selected = False
        End If
    End Sub

    Protected Sub AppBtn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles AppBtn.Click
        '開啟連線
        Dim db As New SqlConnection(connstr)

        '找出二級單位
        Dim Org_Down As New C_Public

        Dim SelNum As Integer
        Dim SelCount As Integer = 0
        Dim intApply1 As Integer = 0 '請假申請單數量
        Dim intApply2 As Integer = 0 '會議室申請單數量
        Dim intApply3 As Integer = 0 '派車申請單數量
        Dim intApply4 As Integer = 0 '房舍水電申請單(新)數量
        Dim intApply5 As Integer = 0 '完工報告單數量
        Dim intApply6 As Integer = 0 '會客洽公申請單數量
        Dim intApply7 As Integer = 0 '資訊設備媒體攜出入申請單數量
        Dim intApply8 As Integer = 0 '報修申請單數量
        Dim intApply9 As Integer = 0 '影印使用申請單數量
        Dim intApply10 As Integer = 0 '呈轉會客洽公申請單數量
        Dim intApply11 As Integer = 0 '上呈主管超過2人申請單數量
        Dim intApply12 As Integer = 0 '其他
        Dim intApply13 As Integer = 0 '影印記錄呈核單數量
        Dim intApply14 As Integer = 0 '門禁會議管制申請單數量
        Dim intApply15 As Integer = 0 '資訊設備維修申請單數量
        '判斷使用者選擇要批核多少張表單
        For SelNum = 0 To GridView1.Rows.Count - 1
            If CType(GridView1.Rows(SelNum).Cells(0).FindControl("selchk"), CheckBox).Checked = True Then
                SelCount = SelCount + 1
            End If
        Next

        If SelCount <= 0 Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('請先勾選您欲核准的表單!!');")
            Response.Write(" </script>")
        Else
            '表單審核
            ''Dim FC As New C_FlowSend.C_FlowSend
            Dim FC As New CFlowSend

            ''簽核訊息
            Dim strQuickShow As String = ""

            Dim i As Integer
            For i = 0 To GridView1.Rows.Count - 1
                If CType(GridView1.Rows(i).Cells(0).FindControl("selchk"), CheckBox).Checked = True Then

                    Dim eformid, employee_id, eformsn, eformrole As String

                    eformid = GridView1.Rows(i).Cells(6).Text
                    employee_id = GridView1.Rows(i).Cells(7).Text
                    eformsn = GridView1.Rows(i).Cells(8).Text
                    eformrole = GridView1.Rows(i).Cells(9).Text

                    Dim Chknextstep As Integer

                    '判斷表單關卡
                    db.Open()
                    Dim strComCheck As New SqlCommand("SELECT nextstep FROM flowctl WHERE eformsn = '" & eformsn & "' and empuid = '" & employee_id & "' and hddate is null", db)
                    Dim RdrCheck As SqlDataReader = strComCheck.ExecuteReader()
                    If RdrCheck.Read() Then
                        Chknextstep = CType(RdrCheck.Item("nextstep"), Integer)
                    End If
                    db.Close()

                    Dim strgroup_id As String = ""
                    Dim strgroup_name As String = ""

                    '判斷上一級主管
                    db.Open()
                    Dim strUpComCheck As New SqlCommand("SELECT a.group_id,b.object_name as group_name FROM flow a left join systemobj b on a.group_id=b.object_uid WHERE eformid = '" & eformid & "' and stepsid = '" & Chknextstep & "' and eformrole=1 ", db)
                    Dim RdrUpCheck As SqlDataReader = strUpComCheck.ExecuteReader()
                    If RdrUpCheck.Read() Then
                        strgroup_id = CType(RdrUpCheck.Item("group_id"), String)
                        strgroup_name = CType(RdrUpCheck.Item("group_name"), String)
                    End If
                    db.Close()

                    Dim NextPer As Integer

                    Dim SendUPVal As String = eformid & "," & employee_id & "," & eformsn & "," & eformrole & ","

                    '關卡為上一級主管有多少人
                    NextPer = CType(FC.F_NextStep(SendUPVal, connstr), Integer)

                    If (NextPer > 1 And strgroup_id = "Z860") Then ''下一關卡中應為不論是不是上一級主管，只要有一筆超過2人簽核，就不可以使用批次簽核
                        '上一級主管兩人以上統計
                        intApply11 += 1
                    Else
                        '批核統計
                        If eformid = "YAqBTxRP8P" Then       '請假申請單
                            intApply1 += 1
                        ElseIf eformid = "4rM2YFP73N" Then   '會議室申請單
                            intApply2 += 1
                        ElseIf eformid = "j2mvKYe3l9" Then   '派車申請單
                            intApply3 += 1
                        ElseIf eformid = "F9MBD7O97G" Then   '房舍水電申請單(新)
                            intApply4 += 1
                        ElseIf eformid = "4ZXNVRV8B6" Then   '完工報告單
                            intApply5 += 1
                        ElseIf eformid = "U28r13D6EA" Then   '會客洽公申請單
                            intApply6 += 1
                        ElseIf eformid = "D6Y95Y5XSU" Then   '資訊設備媒體攜出入申請單
                            intApply7 += 1
                        ElseIf eformid = "9JKSDRR5V3" Then   '報修申請單
                            intApply8 += 1
                        ElseIf eformid = "74BN58683M" Then   '影印使用申請單
                            intApply9 += 1
                        ElseIf eformid = "" Then
                            intApply12 += 1
                            'ElseIf eformid = printRecordsReportID Then '影印記錄呈核單
                            'intApply13 += 1
                        ElseIf eformid = doorAndMeetingControlID Then '門禁會議管制申請單
                            intApply14 += 1
                        ElseIf eformid = "BL7U2QP3IG" Then   '資訊設備維修申請單
                            intApply15 += 1
                        End If

                        '判斷是否為代理人批核
                        If UCase(user_id) <> UCase(employee_id) Then
                            Dim strAgentName As String = ""
                            Dim strComment As String
                            '找尋批核者姓名
                            db.Open()
                            Dim strPer As New SqlCommand("SELECT emp_chinese_name FROM EMPLOYEE WHERE employee_id = '" & user_id & "'", db)
                            Dim RdPer As SqlDataReader = strPer.ExecuteReader()
                            If RdPer.Read() Then
                                strAgentName = CType(RdPer("emp_chinese_name"), String)
                            End If
                            db.Close()
                            strComment = "(此表單已由" & strAgentName & "代理批核)"
                            '增加批核意見
                            insComment(strComment, eformsn, employee_id)
                        End If

                        If eformid = "YAqBTxRP8P" Then '請假申請單
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
                            Dim strPerAgent As New SqlCommand("SELECT PWNAME,PANAME,nAGENT1,nAGENT2,nAGENT3,nSTARTTIME,nSTHOUR,nENDTIME,nETHOUR,PWIDNO,PAIDNO FROM P_01 WHERE EFORMSN = '" & eformsn & "'", db)
                            Dim RdPerAgent As SqlDataReader = strPerAgent.ExecuteReader()
                            If RdPerAgent.Read() Then
                                strPWNAME = CType(RdPerAgent("PWNAME"), String)
                                strPANAME = CType(RdPerAgent("PANAME"), String)
                                strnAGENT1 = CType(RdPerAgent("nAGENT1"), String)
                                strnAGENT2 = CType(RdPerAgent("nAGENT2"), String)
                                strnAGENT3 = CType(RdPerAgent("nAGENT3"), String)
                                strnSTARTTIME = CType(RdPerAgent("nSTARTTIME"), String)
                                strnSTHOUR = CType(RdPerAgent("nSTHOUR"), String)
                                strnENDTIME = CType(RdPerAgent("nENDTIME"), String)
                                strnETHOUR = CType(RdPerAgent("nETHOUR"), String)
                                strPWIDNO = CType(RdPerAgent("PWIDNO"), String)
                                strPAIDNO = CType(RdPerAgent("PAIDNO"), String)
                            End If
                            db.Close()

                            If strnAGENT1 <> "" Then
                                AgentMail(strnAGENT1, eformsn, strPANAME, strnSTARTTIME, strnSTHOUR, strnENDTIME, strnETHOUR)
                            End If
                            If strnAGENT2 <> "" Then
                                AgentMail(strnAGENT2, eformsn, strPANAME, strnSTARTTIME, strnSTHOUR, strnENDTIME, strnETHOUR)
                            End If
                            If strnAGENT3 <> "" Then
                                AgentMail(strnAGENT3, eformsn, strPANAME, strnSTARTTIME, strnSTHOUR, strnENDTIME, strnETHOUR)
                            End If

                            '送MAIL給被代理人
                            If UCase(strPWIDNO) <> UCase(strPAIDNO) Then

                                Dim MOAServer As String = ""
                                Dim SmtpHost As String = ""
                                Dim SystemMail As String = ""
                                Dim MailYN As String = ""
                                MOAServer = CType(FC.F_MailBase("MOAServer", connstr), String)
                                SmtpHost = CType(FC.F_MailBase("SmtpHost", connstr), String)
                                SystemMail = CType(FC.F_MailBase("SystemMail", connstr), String)
                                MailYN = CType(FC.F_MailBase("Mail_Flag", connstr), String)

                                Dim empemail As String = ""

                                '申請者mail
                                db.Open()
                                Dim strPer As New SqlCommand("SELECT empemail FROM EMPLOYEE WHERE employee_id = '" & strPAIDNO & "'", db)
                                Dim RdPer As SqlDataReader = strPer.ExecuteReader()
                                If RdPer.Read() Then
                                    empemail = CType(RdPer("empemail"), String)
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

                            '批核表單
                            Dim Val_P As String

                            Dim SendVal As String = eformid & "," & employee_id & "," & eformsn & "," & eformrole & ","

                            Val_P = CType(FC.F_Send(SendVal, ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString), String)

                            If SelCount = 1 Then
                                Server.Transfer("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=New")
                            End If

                        ElseIf eformid = "U28r13D6EA" Then '會客洽公申請單

                            '判斷批核者是否為副主管,假如批核者為副主管則不可快速送件只可呈轉
                            Dim ParentFlag As String
                            Dim ParentVal As String = employee_id & "," & eformsn
                            Dim PerCount As Integer = 0

                            ParentFlag = CType(FC.F_NextChief(ParentVal, connstr), String)

                            If ParentFlag = "1" Then

                                '上一級主管多少人
                                db.Open()
                                Dim PerCountCom As New SqlCommand("select count(empuid) as PerCount from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = (select ORG_UID from EMPLOYEE where employee_id ='" & user_id & "'))", db)
                                Dim PerRdv As SqlDataReader = PerCountCom.ExecuteReader()
                                If PerRdv.Read() Then
                                    PerCount = CType(PerRdv("PerCount"), Integer)
                                End If
                                db.Close()

                                '上一級沒人則呈現送件按鈕
                                If PerCount = 0 Then

                                    '批核表單
                                    Dim Val_P As String = ""

                                    Dim SendVal As String = eformid & "," & employee_id & "," & eformsn & "," & eformrole & ","

                                    Val_P = CType(FC.F_Send(SendVal, ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString), String)

                                    If SelCount = 1 Then
                                        Server.Transfer("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=New")
                                    End If

                                Else

                                    '登入者為二級單位主管則可核准
                                    Dim strParentOrg As String
                                    strParentOrg = Org_Down.getUporg(org_uid, 2)

                                    If UCase(strParentOrg) = UCase(org_uid) Then

                                        '批核表單
                                        Dim Val_P As String

                                        Dim SendVal As String = eformid & "," & employee_id & "," & eformsn & "," & eformrole & ","

                                        Val_P = CType(FC.F_Send(SendVal, ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString), String)

                                        If SelCount = 1 Then
                                            Server.Transfer("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=New")
                                        End If

                                    Else

                                        Dim strOrgUid As String = ""

                                        '判斷被代理人單位
                                        db.Open()
                                        Dim UpPerCom As New SqlCommand("select ORG_UID from EMPLOYEE where employee_id ='" & employee_id & "'", db)
                                        Dim UpPerRdv As SqlDataReader = UpPerCom.ExecuteReader()
                                        If UpPerRdv.Read() Then
                                            strOrgUid = UpPerRdv("ORG_UID")
                                        End If
                                        db.Close()

                                        '被代理人為二級單位主管則可以送件
                                        If UCase(strParentOrg) = UCase(strOrgUid) Then

                                            '批核表單
                                            Dim Val_P As String = ""

                                            Dim SendVal As String = eformid & "," & employee_id & "," & eformsn & "," & eformrole & ","

                                            Val_P = CType(FC.F_Send(SendVal, ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString), String)

                                            If SelCount = 1 Then
                                                Server.Transfer("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=New")
                                            End If

                                        Else

                                            intApply6 -= 1
                                            intApply10 += 1

                                            '呈轉表單
                                            Dim SendVal As String = ""
                                            SendVal = eformid & "," & employee_id & "," & eformsn & "," & "1"

                                            Dim Val_P As String = FC.F_Transfer(SendVal, connstr)
                                        End If
                                    End If
                                End If
                            Else
                                ''匯出資料至外部資料庫
                                Try
                                    If Chknextstep.ToString.Equals("-1") Then
                                        Using connA As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString2").ConnectionString)
                                            connA.Open()
                                            Dim trans As SqlTransaction
                                            trans = connA.BeginTransaction
                                            ''會客洽公申請單
                                            Dim P_05AData As New P_05A(eformsn)
                                            P_05AData.Insert(trans, connA)
                                            ' ''會客人員明細資料表
                                            'Dim P_0501AData As New P_0501A(eformsn)
                                            'P_0501AData.Insert(trans, connA)
                                            trans.Commit()
                                            trans.Dispose()
                                        End Using
                                        ''更新門禁欄位已匯出時間
                                        Using connA As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
                                            connA.Open()
                                            Dim tran As SqlTransaction
                                            tran = connA.BeginTransaction
                                            Dim cmd As New SqlCommand("UPDATE P_05 SET nCheckDT=GetDate() WHERE EFORMSN='" & eformsn & "'", connA, tran)
                                            cmd.ExecuteNonQuery()
                                            tran.Commit()
                                            tran.Dispose()
                                        End Using
                                    End If
                                Catch sqlex As SqlException
                                    If (Not "-1".Equals(sqlex.Message.IndexOf("無法開啟至 SQL Server 的連接", StringComparison.Ordinal).ToString) Or "-1".Equals(sqlex.Message.IndexOf("登入失敗", StringComparison.Ordinal).ToString)) Then
                                        MessageBox.Show("批核完成\n會客洽公申請單寫入門禁系統連線錯誤，請連繫管理員")
                                    Else
                                        MessageBox.Show("批核完成\n會客洽公申請單寫入門禁系統錯誤\n" & sqlex.Message)
                                    End If
                                Catch ex As Exception
                                    MessageBox.Show(ex.Message)
                                End Try
                                '批核表單
                                Dim SendVal As String = eformid & "," & employee_id & "," & eformsn & "," & eformrole & ","
                                Dim Val_P As String = CType(FC.F_Send(SendVal, ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString), String)

                                If SelCount = 1 Then
                                    Server.Transfer("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=New")
                                End If

                            End If
                        ElseIf eformid = "74BN58683M" Then '影印使用申請單
                            Dim SendVal As String = ""
                            Dim printerEmpuid As String = GridView1.Rows(i).Cells(7).Text '該申請單欲發送之審核者ID
                            Dim printerEformsn As String = GridView1.Rows(i).Cells(8).Text '該申請單之SN

                            SendVal = eformid & "," & printerEmpuid & "," & printerEformsn & "," & "1" & ","
                            '關卡為上一級主管有多少人
                            Dim printerNextPer As Integer = CType(FC.F_NextStep(SendVal, connstr), Integer)
                            Dim printerChknextstep As Integer
                            '判斷表單關卡
                            db.Open()
                            Dim printerComCheck As New SqlCommand("SELECT nextstep FROM flowctl WHERE eformsn = '" & printerEformsn & "' and empuid = '" & printerEmpuid & "' and hddate is null", db)
                            Dim printerRdrCheck = strComCheck.ExecuteReader()
                            If printerRdrCheck.Read() Then
                                printerChknextstep = printerRdrCheck.Item("nextstep")
                            End If
                            RdrCheck.Close()
                            db.Close()

                            'Dim strgroup_id As String = ""


                            'do_sql.G_errmsg = ""

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

                            Dim executeDBScript As String = String.Empty
                            Dim bl_executeResult = False
                            Dim myConn As SqlConnection = New SqlConnection(connstr)
                            Dim myTrans As SqlTransaction
                            Try
                                'Setp 1:先更新log
                                myConn.Open()
                                myTrans = myConn.BeginTransaction()
                                Dim myCommand As New SqlCommand()
                                myCommand.Connection = myConn
                                myCommand.Transaction = myTrans
                                myCommand.CommandText = "update P_08 set AgreeName ='" + strAgentName + "',AgreeORG_UID ='" + org_uid + "',AgreeTime=getdate() where EFORMSN='" + printerEformsn + "'"
                                myCommand.ExecuteNonQuery()
                                executeDBScript = myCommand.CommandText
                                'myCommand.CommandText = "update P_0804 set Update_Datetime= getdate(),Security_Status = 1 where Guid_ID=" + lbSecurity.Text
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
                                myConn.Close()
                            End Try

                            Dim CP As New C_Public
                            '寫入歷程表 2:審核通過=修改
                            'CP.ActionReWrite(IIf(String.IsNullOrEmpty(lbLog_Guid.Text), 0, Integer.Parse(lbLog_Guid.Text)), user_id, 2, executeDBScript)

                            If bl_executeResult = False Then
                                Response.Write(" <script language='javascript'>")
                                Response.Write(" alert('審核失敗，請重新再試，或連絡系統管理人員！');")
                                Response.Write(" </script>")
                                Exit Sub
                            End If

                            Response.Write(" <script language='javascript'>")
                            Response.Write(" alert('" + printerChknextstep.ToString() + "');")
                            Response.Write(" </script>")
                            'ElseIf eformid = printRecordsReportID Then '影印記錄呈核單
                            '    Dim approveResult As String = updatePrintRecordsReportApprove(GridView1.Rows(i).Cells(8).Text, Session("user_id").ToString())
                            '    If Not approveResult.Equals("OK") Then
                            '        Response.Write(" <script language='javascript'>")
                            '        Response.Write(" alert('審核失敗，請重新再試，或連絡系統管理人員！');")
                            '        Response.Write(" </script>")
                            '        Exit Sub
                            '    End If
                        ElseIf eformid = doorAndMeetingControlID Then '門禁會議管制申請單
                            Dim approveResult As String = ApproveP_09(GridView1.Rows(i).Cells(8).Text, Session("user_id").ToString())
                            If Not approveResult.Equals("OK") Then
                                Response.Write(" <script language='javascript'>")
                                Response.Write(" alert('審核失敗，請重新再試，或連絡系統管理人員！');")
                                Response.Write(" </script>")
                                Exit Sub
                            End If
                        ElseIf eformid = "BL7U2QP3IG" Then   '資訊設備維修申請單
                            If NextPer > 1 Then
                                strQuickShow += "簽核表單第" + (i + 1) + "筆中有上呈下一關卡簽核人超過2人表單，請點選詳細資料批核\n"
                            Else
                                '批核表單
                                Dim Val_P As String = ""

                                Dim SendVal As String = eformid & "," & employee_id & "," & eformsn & "," & eformrole & ","

                                Val_P = FC.F_Send(SendVal, ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

                                If SelCount = 1 Then
                                    Server.Transfer("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=New")
                                End If
                            End If
                        Else

                            '批核表單
                            Dim Val_P As String = ""

                            Dim SendVal As String = eformid & "," & employee_id & "," & eformsn & "," & eformrole & ","

                            Val_P = FC.F_Send(SendVal, ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

                            If SelCount = 1 Then
                                Server.Transfer("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=New")
                            End If

                        End If

                    End If

                End If
            Next


            If intApply1 <> 0 Then
                strQuickShow += "核准請假申請單 " & intApply1 & " 筆\n"
            End If

            If intApply2 <> 0 Then
                strQuickShow += "核准會議室申請單 " & intApply2 & " 筆\n"
            End If

            If intApply3 <> 0 Then
                strQuickShow += "核准派車申請單 " & intApply3 & " 筆\n"
            End If

            If intApply4 <> 0 Then
                strQuickShow += "核准房舍水電申請單 " & intApply4 & " 筆\n"
            End If

            If intApply5 <> 0 Then
                strQuickShow += "核准完工報告單 " & intApply5 & " 筆\n"
            End If

            If intApply6 <> 0 Then
                strQuickShow += "核准會客洽公申請單 " & intApply6 & " 筆\n"
            End If

            If intApply10 <> 0 Then
                strQuickShow += "呈轉會客洽公申請單 " & intApply10 & " 筆\n"
            End If

            If intApply9 <> 0 Then
                strQuickShow += "核准銷假申請單 " & intApply9 & " 筆"
            End If

            If intApply13 <> 0 Then
                strQuickShow += "影印記錄呈核單" & intApply13 & " 筆"
            End If

            If intApply14 <> 0 Then
                strQuickShow += "門禁會議管制申請單" & intApply14 & " 筆"
            End If

            If intApply15 <> 0 Then
                strQuickShow += "資訊設備維修申請單" & intApply15 & " 筆"
            End If

            If intApply11 <> 0 Then
                strQuickShow += "上呈主管超過2人共" & intApply11 & " 筆，請點選詳細資料批核\n"
            End If

            If strQuickShow <> "" Then

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('" & strQuickShow & "');")
                Response.Write(" </script>")

            End If

            '清空是否全選
            SelAll = ""

            Server.Transfer("../00/MOA00010.aspx")
        End If
    End Sub

    '核准呈核單
    Private Function updatePrintRecordsReportApprove(ByVal EFORMSN As String, ByVal verifyBy As String) As String
        Dim resultScript As String = ""
        Dim trans As SqlTransaction = Nothing
        Try
            db.Open()
            Dim strSQL = "SELECT Log_Guid FROM ReportP_08Mapping WHERE EFORMSN = '" + EFORMSN + "'"
            Dim dt As New DataTable
            Dim ds As New DataSet
            Dim da As SqlDataAdapter = New SqlDataAdapter(strSQL, db)
            da.Fill(ds)
            dt = ds.Tables(0)
            trans = db.BeginTransaction()
            If dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows '依照對應表內該呈核單所屬Log_Guid Update P_08 影印紀錄資料
                    Dim p08Result As String = UpdateP_08ApprovedByID(db, trans, dr("Log_Guid").ToString(), verifyBy)
                    If Not p08Result.Equals("OK") Then
                        Throw New Exception(p08Result)
                    End If
                Next
                '更新呈核單資料,狀態: 1:待批核 2:核准 3:駁回
                Dim updateSQL = "UPDATE PrintRecordsReport SET Status = 2 ,verifyDate = GETDATE() ,"
                updateSQL = updateSQL + "verifyBy = '" + Session("user_id") + "' WHERE EFORMSN = '" + EFORMSN + "'"
                Dim updateComm As New SqlCommand(updateSQL, db, trans)
                updateComm.ExecuteNonQuery()
                '更新審核流程資料為核准狀態
                Dim flowSQL = "UPDATE flowctl SET hddate = GETDATE(),gonogo = 'E'"
                flowSQL = flowSQL + " WHERE eformsn = '" + EFORMSN + "' AND gonogo = '?' AND nextstep = '-1'"
                Dim flowComm As New SqlCommand(flowSQL, db, trans)
                flowComm.ExecuteNonQuery()
            Else
                Throw New Exception("Get PrintRecordsReport Fail")
            End If
            trans.Commit()
            resultScript = "OK"
        Catch ex As Exception
            trans.Rollback()
            resultScript = "alert(""資料核准失敗:" + ex.Message + """);"
        Finally
            db.Close()
        End Try
        updatePrintRecordsReportApprove = resultScript
    End Function

    '更改影印紀錄資料ApprovedByID欄位,即批核者
    Private Function UpdateP_08ApprovedByID(ByRef db As SqlConnection, ByRef trans As SqlTransaction, ByVal log_Guid As String, ByVal ApprovedByID As String) As String
        Dim strResult As String = "OK"
        Try
            Dim strSQL = "UPDATE P_08 SET ApprovedByID = '" + ApprovedByID + "' WHERE Log_Guid = '" + log_Guid + "'"
            Dim comm As New SqlCommand(strSQL, db, trans)
            comm.ExecuteNonQuery()
        Catch ex As Exception
            strResult = ex.Message
        End Try
        UpdateP_08ApprovedByID = strResult
    End Function

    '核准門禁會議管制申請單
    Private Function ApproveP_09(ByVal EFORMSN As String, ByVal verifyBy As String) As String
        Dim result = "OK"
        Dim conn As SqlConnection = New SqlConnection(connstr)
        Try
            conn.Open()
            '表單狀態 [ 0：未登管 1：已登管 2：退件 3：刪除 ]
            Dim myCommand As New SqlCommand("UPDATE P_09 SET Status = 1, ModifyBy = @ModifyBy, ModifyDate = GETDATE() WHERE EFORMSN = @EFORMSN", conn)
            myCommand.Parameters.Add("@ModifyBy", SqlDbType.VarChar, 10).Value = verifyBy
            myCommand.Parameters.Add("@EFORMSN", SqlDbType.VarChar, 16).Value = EFORMSN
            myCommand.ExecuteNonQuery()
        Catch ex As Exception
            result = ex.Message
        Finally
            conn.Close()
        End Try

        '表單審核
        Dim FCC As New CFlowSend
        Dim SendVal As String = doorAndMeetingControlID & "," & verifyBy & "," & EFORMSN & "," & "1" & ","
        Dim Val_P As String = FCC.F_Send(SendVal, connstr)
        Return result
    End Function

    '駁回呈核單
    Private Function updatePrintRecordsReportReject(ByVal EFORMSN As String, ByVal verifyBy As String) As String
        Dim resultScript As String = ""
        Dim trans As SqlTransaction = Nothing
        Try
            db.Open()
            Dim strSQL = "SELECT Log_Guid FROM ReportP_08Mapping WHERE EFORMSN = '" + EFORMSN + "'"
            Dim dt As New DataTable
            Dim ds As New DataSet
            Dim da As SqlDataAdapter = New SqlDataAdapter(strSQL, db)
            da.Fill(ds)
            dt = ds.Tables(0)
            trans = db.BeginTransaction()
            If dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows '依照對應表內該呈核單所屬Log_Guid Update P_08 影印紀錄資料
                    Dim p08Result As String = ClearP_08VerifyRequesterID(db, trans, dr("Log_Guid").ToString())
                    If Not p08Result.Equals("OK") Then
                        Throw New Exception(p08Result)
                    End If
                Next
                '呈核單狀態: 1:待批核 2:核准 3:駁回
                Dim updateSQL = "UPDATE PrintRecordsReport SET Status = 3 ,verifyDate = GETDATE() ,"
                updateSQL = updateSQL + "verifyBy = '" + verifyBy + "' WHERE EFORMSN = '" + EFORMSN + "'"
                Dim updateComm As New SqlCommand(updateSQL, db, trans)
                updateComm.ExecuteNonQuery()
                '更新審核流程資料為駁回狀態
                Dim flowSQL = "UPDATE flowctl SET hddate = GETDATE(),gonogo = '0'"
                flowSQL = flowSQL + " WHERE eformsn = '" + EFORMSN + "' AND gonogo = '?' AND nextstep = '-1'"
                Dim flowComm As New SqlCommand(flowSQL, db, trans)
                flowComm.ExecuteNonQuery()
            Else
                Throw New Exception("Get PrintRecordsReport Fail")
            End If
            trans.Commit()
            resultScript = "OK"
        Catch ex As Exception
            trans.Rollback()
            resultScript = ex.Message
        Finally
            db.Close()
        End Try
        updatePrintRecordsReportReject = resultScript
    End Function

    '清除影印紀錄資料VerifyRequesterID欄位,即呈核申請者,亦即將該筆資料設定為未呈核狀態
    Private Function ClearP_08VerifyRequesterID(ByRef db As SqlConnection, ByRef trans As SqlTransaction, ByVal log_Guid As String) As String
        Dim strResult As String = "OK"
        Try
            Dim strSQL = "UPDATE P_08 SET VerifyRequesterID = '' WHERE Log_Guid = '" + log_Guid + "'"
            Dim comm As New SqlCommand(strSQL, db, trans)
            comm.ExecuteNonQuery()
        Catch ex As Exception
            strResult = ex.Message
        End Try
        ClearP_08VerifyRequesterID = strResult
    End Function

    '駁回門禁會議管制申請單
    Private Function RejectP_09(ByVal EFORMSN As String, ByVal verifyBy As String) As String
        Dim result As String = "OK"
        Try
            Dim conn As SqlConnection = New SqlConnection(connstr)
            conn.Open()
            '表單狀態 [ 0：未登管 1：已登管 2：退件 3：刪除 ]
            Dim myCommand As New SqlCommand("UPDATE P_09 SET Status = 2, ModifyBy = @ModifyBy, ModifyDate = GETDATE() WHERE EFORMSN = @EFORMSN", conn)
            myCommand.Parameters.Add("@ModifyBy", SqlDbType.VarChar, 10).Value = verifyBy
            myCommand.Parameters.Add("@EFORMSN", SqlDbType.VarChar, 16).Value = EFORMSN
            myCommand.ExecuteNonQuery()
            conn.Close()

            Dim Val_P As String = ""
            '表單駁回
            Dim FC As New C_FlowSend.C_FlowSend
            Dim SendVal As String = doorAndMeetingControlID & "," & verifyBy & "," & EFORMSN & "," & "1"
            Val_P = FC.F_Back(SendVal, connstr)
        Catch ex As Exception
            result = ex.Message
        End Try
        Return result
    End Function

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        '分頁
        Dim strOrd As String
        strOrd = " ORDER BY appdate DESC"
        SqlDataSource2.SelectCommand = SQLALL(user_id) & strOrd
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated

        If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then
            '隱藏eformid,employee_id,eformsn,eformrole
            e.Row.Cells(6).Visible = False
            e.Row.Cells(7).Visible = False
            e.Row.Cells(8).Visible = False
            e.Row.Cells(9).Visible = False
        End If
    End Sub


    Protected Sub AllChk_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AllChk.CheckedChanged

        '是否全選
        Dim i As Integer = 0
        If AllChk.Checked = False Then
            For i = 0 To GridView1.Rows.Count - 1
                CType(GridView1.Rows(i).Cells(0).FindControl("selchk"), CheckBox).Checked = False
            Next
            '沒全選
            SelAll = ""
        Else
            For i = 0 To GridView1.Rows.Count - 1
                If CType(GridView1.Rows(i).Cells(0).FindControl("HiddenField1"), HiddenField).Value <> "F9MBD7O97G" Then
                    CType(GridView1.Rows(i).Cells(0).FindControl("selchk"), CheckBox).Checked = True
                End If
            Next
            '全選
            SelAll = "1"
        End If

    End Sub

    Protected Sub DropDownList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged
        '判斷要審核哪張表單
        Dim strsql As String = SQLALL(user_id)
        'Dim strsql As String = "SELECT flowctl.eformid, flowctl.eformrole, flowctl.appdate, flowctl.empuid,(SELECT emp_chinese_name FROM flowctl AS f WHERE (eformsn = flowctl.eformsn) AND (stepsid = 1)) AS emp_chinese_name, EFORMS.frm_chinese_name, flowctl.eformsn, flowctl.hddate, V_EformShow.ShowContent FROM flowctl INNER JOIN EFORMS ON flowctl.eformid = EFORMS.eformid INNER JOIN V_EformShow ON flowctl.eformsn = V_EformShow.EFORMSN WHERE (flowctl.hddate IS NULL) and (flowctl.gonogo='?' OR flowctl.gonogo='R') and flowctl.empuid = '" & user_id & "'"

        If DropDownList1.SelectedIndex = 0 Then
            strsql = strsql
        ElseIf DropDownList1.SelectedIndex <> 0 Then
            strsql += " and flowctl.eformid = '" & DropDownList1.SelectedValue & "'"
        End If
        SqlDataSource2.SelectCommand = strsql & " ORDER BY appdate DESC"
    End Sub

    Protected Sub BackBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BackBtn.Click
        '將申請的會議室撤回
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string
        '開啟連線
        Dim db As New SqlConnection(connstr)
        '表單駁回
        Dim FC As New C_FlowSend.C_FlowSend
        Dim SelCount As Int32 = 0

        '判斷使用者選擇要駁回多少張表單
        For SelNum As Int32 = 0 To GridView1.Rows.Count - 1
            If CType(GridView1.Rows(SelNum).Cells(0).FindControl("selchk"), CheckBox).Checked = True Then
                SelCount = SelCount + 1
            End If
        Next

        If SelCount <= 0 Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('請先勾選您欲駁回的表單!!');")
            Response.Write(" </script>")
        Else
            Dim i As Integer = 0
            For i = 0 To GridView1.Rows.Count - 1
                If CType(GridView1.Rows(i).Cells(0).FindControl("selchk"), CheckBox).Checked = True Then

                    Dim eformid, employee_id, eformsn, eformrole As String

                    eformid = GridView1.Rows(i).Cells(6).Text
                    employee_id = GridView1.Rows(i).Cells(7).Text
                    eformsn = GridView1.Rows(i).Cells(8).Text
                    eformrole = GridView1.Rows(i).Cells(9).Text

                    Dim Val_P As String = ""

                    Dim SendVal As String = eformid & "," & employee_id & "," & eformsn & "," & eformrole

                    Val_P = FC.F_Back(SendVal, ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

                    '駁回會議室申請時刪除會議室申請時段
                    If eformid = "4rM2YFP73N" Then

                        '將會議室申請時段放到移除資料表
                        db.Open()
                        Dim insCom As New SqlCommand("INSERT INTO P_0205 (MTT_Num, EFORMSN, MeetSn, MeetTime, MeetHour, DelUser) SELECT MTT_Num, EFORMSN, MeetSn, MeetTime, MeetHour, '" & employee_id & "' FROM P_0204 WHERE EFORMSN='" & eformsn & "'", db)
                        insCom.ExecuteNonQuery()
                        db.Close()

                        '刪除申請的會議室時段
                        db.Open()
                        Dim delCom As New SqlCommand("DELETE FROM P_0204 WHERE (EFORMSN='" & eformsn & "')", db)
                        delCom.ExecuteNonQuery()
                        db.Close()
                        'ElseIf eformid = printRecordsReportID Then '影印記錄呈核單
                        '    Dim rejectResult As String = updatePrintRecordsReportReject(GridView1.Rows(i).Cells(8).Text, Session("user_id").ToString())
                        '    If Not rejectResult.Equals("OK") Then
                        '        Response.Write(" <script language='javascript'>")
                        '        Response.Write(" alert('審核失敗，請重新再試，或連絡系統管理人員！');")
                        '        Response.Write(" </script>")
                        '        Exit Sub
                        '    End If
                    ElseIf eformid = doorAndMeetingControlID Then '門禁會議管制申請單
                        Dim rejectResult As String = RejectP_09(GridView1.Rows(i).Cells(8).Text, Session("user_id").ToString())
                        If Not rejectResult.Equals("OK") Then
                            Response.Write(" <script language='javascript'>")
                            Response.Write(" alert('審核失敗，請重新再試，或連絡系統管理人員！');")
                            Response.Write(" </script>")
                            Exit Sub
                        End If
                    End If

                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('表單駁回完成');")
                    Response.Write(" </script>")

                End If
            Next
            '清空是否全選
            SelAll = ""
            Server.Transfer("../00/MOA00010.aspx")
        End If
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType.Equals(DataControlRowType.DataRow) Then
            Dim selchk As CheckBox
            Dim HiddenField1 As HiddenField

            selchk = e.Row.Cells(0).FindControl("selchk")
            HiddenField1 = e.Row.Cells(0).FindControl("HiddenField1") ''表單編號

            If HiddenField1.Value = "F9MBD7O97G" Then
                '房舍水電報修單無法直接做核准和駁回
                selchk.Visible = False
            Else
                '如果有非房舍水電報修單者才會有勾選處理的功能，有勾選才能有做核准和駁回的動作
                AllChk.Enabled = True
            End If

            Select Case HiddenField1.Value
                Case "BL7U2QP3IG"
                    ''限制表單內容不超過總字數25個字，以免影響版面編排
                    If e.Row.Cells(3).Text.Length > 25 Then
                        e.Row.Cells(3).Text = e.Row.Cells(3).Text.Substring(0, 25) + "..."
                    End If
            End Select
        End If
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Dim streformid, streformsn As String
        Dim strPath As String = ""

        '顯示選取的表單資料
        streformid = GridView1.Rows(GridView1.SelectedIndex.ToString()).Cells(6).Text
        streformsn = GridView1.Rows(GridView1.SelectedIndex.ToString()).Cells(8).Text

        '表單資料夾

        If streformid = "YAqBTxRP8P" Then       '請假申請單
            strPath = "MOA00020.aspx?x=MOA01001&y=" & streformid & "&Read_Only=2&EFORMSN=" & streformsn
        ElseIf streformid = "4rM2YFP73N" Then   '會議室申請單
            strPath = "MOA00020.aspx?x=MOA02001&y=" & streformid & "&Read_Only=2&EFORMSN=" & streformsn
        ElseIf streformid = "j2mvKYe3l9" Then   '派車申請單
            strPath = "MOA00020.aspx?x=MOA03001&y=" & streformid & "&Read_Only=2&EFORMSN=" & streformsn
        ElseIf streformid = "61TY3LELYT" Then   '房舍水電申請單
            strPath = "MOA00020.aspx?x=MOA04001&y=" & streformid & "&Read_Only=2&EFORMSN=" & streformsn
        ElseIf streformid = "F9MBD7O97G" Then   '房舍水電申請單(新) peter
            strPath = "MOA00020.aspx?x=MOA04100&y=" & streformid & "&Read_Only=2&EFORMSN=" & streformsn
        ElseIf streformid = "4ZXNVRV8B6" Then   '完工報告單
            strPath = "MOA00020.aspx?x=MOA04003&y=" & streformid & "&Read_Only=2&EFORMSN=" & streformsn
        ElseIf streformid = "U28r13D6EA" Then   '會客洽公申請單
            strPath = "MOA00020.aspx?x=MOA05001&y=" & streformid & "&Read_Only=2&EFORMSN=" & streformsn
        ElseIf streformid = "D6Y95Y5XSU" Then   '資訊設備媒體攜出入申請單
            strPath = "MOA00020.aspx?x=MOA06001&y=" & streformid & "&Read_Only=2&EFORMSN=" & streformsn
        ElseIf streformid = "9JKSDRR5V3" Then   '報修申請單
            strPath = "MOA00020.aspx?x=MOA07001&y=" & streformid & "&Read_Only=2&EFORMSN=" & streformsn
        ElseIf streformid = "5D82872F5L" Then   '銷假申請單
            strPath = "MOA00020.aspx?x=MOA01003&y=" & streformid & "&Read_Only=2&EFORMSN=" & streformsn
        ElseIf streformid = "74BN58683M" Then   '影印使用申請單
            strPath = "MOA00020.aspx?x=MOA08001&y=" & streformid & "&Read_Only=2&EFORMSN=" & streformsn
            'ElseIf streformid = printRecordsReportID Then   '影印紀錄呈核單
            'strPath = "MOA00020.aspx?x=MOA08014&y=" & streformid & "&Read_Only=2&EFORMSN=" & streformsn
        ElseIf streformid = doorAndMeetingControlID Then   '門禁會議管制申請單
            strPath = "MOA00020.aspx?x=MOA09001&y=" & streformid & "&Read_Only=2&EFORMSN=" & streformsn
        ElseIf streformid = "BL7U2QP3IG" Then   '資訊設備維修申請單
            strPath = "MOA00020.aspx?x=MOA11001&y=" & streformid & "&Read_Only=2&EFORMSN=" & streformsn
        End If

        'Response.Write(strPath)
        'Return
        Response.Write(" <script language='javascript'>")
        Response.Write(" sPath = '" & strPath & "';")
        Response.Write(" strFeatures = 'dialogWidth=900px;dialogHeight=750px;help=no;status=no;resizable=yes;scroll=no;dialogTop=100;dialogLeft=100';")
        Response.Write(" showModalDialog(sPath,self,strFeatures);")
        Response.Write(" </script>")

        GridView1.DataSourceID = Nothing
        GridView1.DataBind()
        GridView1.DataSourceID = SqlDataSource2.ID
        GridView1.DataBind()

    End Sub

    '取得表單種類ID
    Private Function GetEFormId(ByVal eformName As String) As String
        db.Open()
        Dim dr As SqlDataReader = New SqlCommand("SELECT eformid FROM EFORMS WHERE frm_chinese_name = '" + eformName + "'", db).ExecuteReader
        If dr.Read Then
            GetEFormId = dr("eformid")
        Else
            GetEFormId = ""
        End If
        'GetEFormId = sqlcomm.ExecuteScalar()
        db.Close()
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        'lbMsg.Text = ""
        'session被清空回首頁
        If user_id = "" Or org_uid = "" Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")
        Else

            '是否由未批核表單頁面傳送過來
            Dim strHyper As String = ""
            strHyper = Request.QueryString("strHyper")
            Try

                If IsPostBack = False Then
                    If strHyper <> "" Then

                        If strHyper = "YAqBTxRP8P" Or strHyper = "4rM2YFP73N" Or strHyper = "j2mvKYe3l9" _
                            Or strHyper = "61TY3LELYT" Or strHyper = "U28r13D6EA" Or strHyper = "5D82872F5L" _
                            Or strHyper = "4ZXNVRV8B6" Or strHyper = "F9MBD7O97G" Or strHyper = "S9QR2W8X6U" _
                            Or strHyper = "BL7U2QP3IG" Then '加新表單
                            '判斷要審核哪張表單
                            Dim strsql As String = SQLALL(user_id)
                            strsql += " and flowctl.eformid = '" & strHyper & "'"
                            SqlDataSource2.SelectCommand = strsql & " ORDER BY appdate DESC"
                        Else

                            '判斷登入者權限
                            Dim LoginCheck As New C_Public

                            '禁止無帳號者竄入
                            LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00010.aspx")
                            Response.End()

                        End If
                        

                    Else

                        '判斷登入者權限
                        Dim LoginCheck As New C_Public

                        If LoginCheck.LoginCheck(user_id, "MOA00010") <> "" Then
                            LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00010.aspx")
                            Response.End()
                        End If

                        ''判斷要審核哪張表單
                        Dim strsql As String = SQLALL(user_id)
                        SqlDataSource2.SelectCommand = strsql & " ORDER BY appdate DESC"
                    End If
                End If
            Catch ex As Exception
                'lbMsg.Text = ex.Message
            End Try
            '如果查詢結果都沒有任何記錄，則勾選，核準，駁回3個按鈕皆無法點用
            If GridView1.Rows.Count < 1 Then
                AppBtn.Enabled = False
                BackBtn.Enabled = False
                AllChk.Enabled = False
            End If
        End If

    End Sub

    Public Function SQLALL(ByVal UserSel As String)

        Dim strsql As String = ""

        '找出代理表單
        Dim AgentFlag As New C_Public

        If AgentFlag.AgentAll(user_id, Now.Date) = "" Then

            '整合SQL搜尋字串
            strsql = "SELECT flowctl.eformid, flowctl.eformrole, flowctl.appdate, flowctl.empuid, EFORMS.frm_chinese_name, flowctl.eformsn, flowctl.hddate, V_EformShow.ShowContent, V_EformShow.PWNAME AS emp_chinese_name FROM flowctl INNER JOIN EFORMS ON flowctl.eformid = EFORMS.eformid INNER JOIN V_EformShow ON flowctl.eformsn = V_EformShow.EFORMSN WHERE (flowctl.hddate IS NULL) and (flowctl.gonogo='?' OR flowctl.gonogo='R') and flowctl.empuid = '" & UserSel & "'"

        Else

            '整合SQL搜尋字串
            strsql = "SELECT flowctl.eformid, flowctl.eformrole, flowctl.appdate, flowctl.empuid, EFORMS.frm_chinese_name, flowctl.eformsn, flowctl.hddate, V_EformShow.ShowContent, V_EformShow.PWNAME AS emp_chinese_name FROM flowctl INNER JOIN EFORMS ON flowctl.eformid = EFORMS.eformid INNER JOIN V_EformShow ON flowctl.eformsn = V_EformShow.EFORMSN WHERE (flowctl.hddate IS NULL) and (flowctl.gonogo='?' OR flowctl.gonogo='R') and (flowctl.empuid = '" & UserSel & "' OR flowctl.empuid IN (" & AgentFlag.AgentAll(user_id, Now.Date) & "))"

        End If

        SQLALL = strsql

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
            'lbMsg.Text = ex.Message
        End Try

        insComment = ""

    End Function

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

            '找出同級單位以下全部單位
            Dim Org_Down As New C_Public

            Dim empemail As String = ""

            '找尋代理人mail
            db.Open()
            Dim strPer As New SqlCommand("SELECT empemail FROM EMPLOYEE WHERE ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ") AND emp_chinese_name = '" & strnAGENT & "'", db)
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
            'lbMsg.Text = ex.Message
        End Try

        AgentMail = ""

    End Function

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted
        Dim strOrd As String
        '排序條件
        strOrd = " ORDER BY " & GridView1.SortExpression.ToString()
        SqlDataSource2.SelectCommand = SQLALL(user_id) & strOrd
    End Sub


    Public Function getFlowStatus(ByVal _streformsn As String) As String
        '流程狀態
        Dim strFlowStatus As String = ""
        Dim _strstepsid As String = ""
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        db.Open()
        Dim strstepChk As New SqlCommand("select stepsid from dbo.flowctl where eformsn='" & _streformsn & "' and gonogo ='?'", db)
        Dim RdrstepChk = strstepChk.ExecuteReader()

        If RdrstepChk.read() Then
            _strstepsid = RdrstepChk.item("stepsid")
        End If
        db.Close()

        Select Case _strstepsid
            Case "29468"
                strFlowStatus = "1" '新送件
            Case "19420", "21079", "21027", "1032"
                strFlowStatus = "2" '處理中
            Case "21080"
                strFlowStatus = "3" '待料中
            Case "1182", "1135"
                strFlowStatus = "4" '完工
        End Select
        Return strFlowStatus

    End Function

    Protected Sub DropDownList1_DataBound(sender As Object, e As System.EventArgs) Handles DropDownList1.DataBound
        If Not IsNothing(DropDownList1.Items.FindByValue(Request("strHyper"))) Then
            DropDownList1.SelectedValue = Request("strHyper")
        End If
    End Sub
End Class
''' <summary>  
'''Summary description for MessageBox.  
''' </summary>  
Public Class MessageBox

    Private Shared m_executingPages As New Hashtable()
    Private Sub MessageBox()

    End Sub

    ''' <summary>  
    ''' MessageBox訊息窗  
    ''' </summary>  
    ''' <param name="sMessage">要顯示的訊息</param>  
    Public Shared Sub Show(ByVal sMessage As String)

        '' If this is the first time a page has called this method then  
        If (Not m_executingPages.Contains(HttpContext.Current.Handler)) Then

            '' Attempt to cast HttpHandler as a Page.  
            Dim executingPage As Page = CType(HttpContext.Current.Handler, Page)
            If (executingPage IsNot Nothing) Then

                '' Create a Queue to hold one or more messages.  
                Dim messageQueue As Queue = New Queue()
                '' Add our message to the Queue  
                messageQueue.Enqueue(sMessage)

                '' Add our message queue to the hash table. Use our page reference  
                '' (IHttpHandler) as the key.  
                m_executingPages.Add(HttpContext.Current.Handler, messageQueue)
                '' Wire up Unload event so that we can inject some JavaScript for the alerts.  
                ''executingPage.Unload += New EventHandler(ExecutingPage_Unload)
                AddHandler executingPage.Unload, AddressOf ExecutingPage_Unload
            End If
        Else

            '' If were here then the method has allready been called from the executing Page.  
            '' We have allready created a message queue and stored a reference to it in our hastable.   
            Dim queue As Queue = CType(m_executingPages(HttpContext.Current.Handler), Queue)

            '' Add our message to the Queue  
            queue.Enqueue(sMessage)
        End If
    End Sub
    Public Shared Sub RedirectShow(ByVal sMessage As String)
        '' If this is the first time a page has called this method then  
        If (Not m_executingPages.Contains(HttpContext.Current.Handler)) Then

            '' Attempt to cast HttpHandler as a Page.  
            Dim executingPage As Page = CType(HttpContext.Current.Handler, Page)
            If (executingPage IsNot Nothing) Then
                '' Create a Queue to hold one or more messages.  
                Dim messageQueue As New Queue()
                '' Add our message to the Queue  
                messageQueue.Enqueue(sMessage)

                '' Add our message queue to the hash table. Use our page reference  
                '' (IHttpHandler) as the key.  
                m_executingPages.Add(HttpContext.Current.Handler, messageQueue)
                '' Wire up Unload event so that we can inject some JavaScript for the alerts.  
                ''executingPage.Unload += new EventHandler(ExecutingPageRedirect_Unload);
                AddHandler executingPage.Unload, AddressOf ExecutingPageRedirect_Unload
            End If

        Else

            '' If were here then the method has allready been called from the executing Page.  
            '' We have allready created a message queue and stored a reference to it in our hastable.   
            Dim queue As Queue = CType(m_executingPages(HttpContext.Current.Handler), Queue)

            '' Add our message to the Queue  
            queue.Enqueue(sMessage)
        End If
    End Sub
    '' Our page has finished rendering so lets output the JavaScript to produce the alert's  
    Private Shared Sub ExecutingPage_Unload(sender As Object, e As EventArgs)

        '' Get our message queue from the hashtable  
        Dim queue As Queue = CType(m_executingPages(HttpContext.Current.Handler), Queue)

        If (queue IsNot Nothing) Then

            Dim sb As New StringBuilder()
            '' How many messages have been registered?  
            Dim iMsgCount As Integer = queue.Count
            '' Use StringBuilder to build up our client slide JavaScript.  
            sb.Append("<script language='javascript'>")
            '' Loop round registered messages  
            Dim sMsg As String
            While (iMsgCount > 0)

                sMsg = CType(queue.Dequeue(), String)
                ''sMsg = sMsg.Replace( "\n", "\\n" ); //這部分是我mark掉的  
                sMsg = sMsg.Replace("""", "'")
                ''W3c建議要避開的危險字元  
                ''&;`'\"|*?~<>^()[]{}$\n\r  
                sMsg = sMsg.Replace("\n", "_")
                sMsg = sMsg.Replace("\r", "_")
                sb.Append("alert( """ + sMsg + """ );")
                iMsgCount -= 1
            End While
            '' Close our JS  
            sb.Append("</script>")
            '' Were done, so remove our page reference from the hashtable  
            m_executingPages.Remove(HttpContext.Current.Handler)
            '' Write the JavaScript to the end of the response stream.  
            HttpContext.Current.Response.Write(sb.ToString())
        End If
    End Sub
    Private Shared Sub ExecutingPageRedirect_Unload(sender As Object, e As EventArgs)

        '' Get our message queue from the hashtable  
        Dim queue As Queue = CType(m_executingPages(HttpContext.Current.Handler), Queue)

        If (queue IsNot Nothing) Then
            Dim sb As New StringBuilder()
            '' How many messages have been registered?  
            Dim iMsgCount As Integer = queue.Count
            '' Use StringBuilder to build up our client slide JavaScript.  
            sb.Append("<script language='javascript'>")
            '' Loop round registered messages  
            Dim sMsg As String
            While (iMsgCount > 0)
                sMsg = CType(queue.Dequeue(), String)
                ''sMsg = sMsg.Replace( "\n", "\\n" ); //這部分是我mark掉的  
                sMsg = sMsg.Replace("""", "'")
                ''W3c建議要避開的危險字元  
                ''&;`'\"|*?~<>^()[]{}$\n\r  
                sMsg = sMsg.Replace("\n", "_")
                sMsg = sMsg.Replace("\r", "_")
                sb.Append("alert( """ + sMsg + """ );")
                sb.Append("top.location.href=\index.aspx\;\n")
                iMsgCount -= 1
            End While
            '' Close our JS  
            sb.Append("</script>")
            '' Were done, so remove our page reference from the hashtable  
            m_executingPages.Remove(HttpContext.Current.Handler)
            '' Write the JavaScript to the end of the response stream.  
            HttpContext.Current.Response.Write(sb.ToString())
        End If
    End Sub
End Class
