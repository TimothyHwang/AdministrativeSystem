Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System
Imports Microsoft.VisualBasic

Public Class CFlowSend
    Public Sub New()

    End Sub

    ''' <summary>
    ''' 找尋組織人員資料
    ''' </summary>
    ''' <param name="employee_id"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function F_AdminEmp(ByVal employee_id As String, ByVal conn As String) As Object
        Try
            F_AdminEmp = ""

            '開啟連線
            Dim db As New SqlConnection(conn)

            '人員資料
            db.Open()
            Dim StepSQL As String = "select e.emp_chinese_name,e.empemail,e.member_uid,e.ORG_UID,a.ORG_NAME from ADMINGROUP a,EMPLOYEE e where e.ORG_UID = a.ORG_UID and e.employee_id = '" & employee_id & "'"
            Dim ds As New DataSet
            Call New SqlDataAdapter(StepSQL, db).Fill(ds)
            F_AdminEmp = ds.Tables.Item(0)
            db.Close()

            'Dim da As SqlDataAdapter = New SqlDataAdapter(StepSQL, db)
            'da.Fill(ds)
            'db.Close()
            'Return ds.Tables(0)

        Catch ex As Exception
            F_AdminEmp = ex.Message
        End Try
        Return F_AdminEmp
    End Function

    ''' <summary>
    ''' 重新分派
    ''' </summary>
    ''' <param name="FormArr"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function F_Assign(ByVal FormArr As String, ByVal conn As String) As Object
        F_Assign = ""

        Dim FormList As Object = Split(FormArr, ",", -1, CompareMethod.Binary)

        Dim strTran As String = ""
        Dim D_eformid As String = ""
        Dim D_eformrole As String = ""
        Dim D_eformsn As String = ""
        Dim D_group_name As String = ""
        Dim D_stepsid As String = ""
        Dim D_steps As String = ""
        Dim D_empname As String = ""
        Dim D_nextstep As String = ""
        Dim D_org_uid As String = ""

        Try
            '分派者id
            Dim employee_id As String = FormList(0).ToString()
            Dim flowsn As String = FormList(1).ToString()

            '開啟連線
            Dim db As New SqlConnection(conn)

            '取得該筆表單資料
            db.Open()
            Dim RdTranDetail As Object = New SqlCommand("SELECT * FROM flowctl WHERE (flowsn = '" & flowsn & "')", db).ExecuteReader
            If RdTranDetail.Read() Then
                D_eformid = RdTranDetail.Item("eformid").ToString()
                D_eformrole = RdTranDetail.Item("eformrole").ToString()
                D_eformsn = RdTranDetail.Item("eformsn").ToString()
                D_group_name = RdTranDetail.Item("group_name").ToString()
                D_stepsid = RdTranDetail.Item("stepsid").ToString()
                D_steps = RdTranDetail.Item("steps").ToString()
                D_nextstep = RdTranDetail.Item("nextstep").ToString()
            End If
            db.Close()

            '表單變更成重新分派狀態
            db.Open()
            Call New SqlCommand("UPDATE flowctl SET hddate = GETDATE(),gonogo = 'R' WHERE (flowsn = '" & flowsn & "') ", db).ExecuteNonQuery()
            db.Close()

            '取得分派者資料
            db.Open()
            Dim Rdv As Object = New SqlCommand("select emp_chinese_name,org_uid from employee where employee_id = '" & employee_id & "'", db).ExecuteReader()
            If Rdv.Read() Then
                D_org_uid = Rdv.Item("org_uid").ToString()
                D_empname = Rdv.Item("emp_chinese_name").ToString()
            End If
            db.Close()

            '表單重新分派
            db.Open()
            strTran = "INSERT INTO flowctl(eformid,eformrole,eformsn,stepsid,steps,empuid,emp_chinese_name,group_name,gonogo,nextstep,important,recdate,appdate,deptcode,createdate) "
            strTran += " VALUES ('" & D_eformid & "','" & D_eformrole & "','" & D_eformsn & "','" & D_stepsid & "','" & D_steps & "','" & employee_id & "','" & D_empname & "','" & D_group_name & "','?', '" & D_nextstep & "','1',GETDATE(),GETDATE(),'" & D_org_uid & "',GETDATE()) "
            Dim strTranUpd As New SqlCommand(strTran, db)
            strTranUpd.ExecuteNonQuery()
            db.Close()

            F_Assign = "表單重新分派完成"

        Catch ex As Exception
            F_Assign = ex.Message
        End Try
        Return F_Assign
    End Function

    ''' <summary>
    ''' 駁回
    ''' </summary>
    ''' <param name="FormArr"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_Back(ByVal FormArr As String, ByVal conn As String) As Object
        F_Back = ""

        'eformid,employee_id,eformsn,eformrole
        Dim FormList As Object = Split(FormArr, ",", -1, CompareMethod.Binary)
        Dim strPath As String = ""
        Dim streformName As String = ""

        Try
            Dim eformid As String = FormList(0).ToString()
            Dim employee_id As String = FormList(1).ToString()
            Dim eformsn As String = FormList(2).ToString()
            Dim eformrole As String = FormList(3).ToString()

            If eformid = "YAqBTxRP8P" Then       '請假申請單
                streformName = "請假申請單"
                strPath = "MOA00020.aspx?x=MOA01001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "4rM2YFP73N" Then   '會議室申請單
                streformName = "會議室申請單"
                strPath = "MOA00020.aspx?x=MOA02001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "j2mvKYe3l9" Then   '派車申請單
                streformName = "派車申請單"
                strPath = "MOA00020.aspx?x=MOA03001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "61TY3LELYT" Then   '房舍水電申請單
                streformName = "房舍水電申請單"
                strPath = "MOA00020.aspx?x=MOA04100&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn '表單路徑
            ElseIf eformid = "F9MBD7O97G" Then   '房舍水電申請單(新) peter
                streformName = "房舍水電申請單(博)"
                strPath = "MOA00020.aspx?x=MOA04100&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn '表單路徑
            ElseIf eformid = "4ZXNVRV8B6" Then   '完工報告單
                streformName = "完工報告單"
                strPath = "MOA00020.aspx?x=MOA04003&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "U28r13D6EA" Then   '會客洽公申請單
                streformName = "會客洽公申請單"
                strPath = "MOA00020.aspx?x=MOA05001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "D6Y95Y5XSU" Then   '資訊設備媒體攜出入申請單
                streformName = "資訊設備媒體攜出入申請單"
                strPath = "MOA00020.aspx?x=MOA06001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "9JKSDRR5V3" Then   '報修申請單
                streformName = "報修申請單"
                strPath = "MOA00020.aspx?x=MOA07001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "5D82872F5L" Then   '銷假申請單
                streformName = "銷假申請單"
                strPath = "MOA00020.aspx?x=MOA01003&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "74BN58683M" Then   '影印使用申請單
                streformName = "影印使用申請單"
                strPath = "MOA00020.aspx?x=MOA08001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
                'ElseIf eformid = "BL7U2QP3IG" Then   '資訊設備維修申請單
                '    streformName = "資訊設備維修申請單"
                '    strPath = "MOA00020.aspx?x=MOA11001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑           
            End If

            '開啟連線
            Dim db As New SqlConnection(conn)

            Dim NowFlow As New flowctl(eformid, eformsn, "?")

            ''刪除此表單的所有批核結果
            ''只留下第一和第二個關卡資料
            'db.Open()
            'Dim delCom As New SqlCommand("DELETE FROM flowctl WHERE (eformsn = '" & eformsn & "') AND (stepsid <> '" & nextstep & "' AND stepsid <> '1')", db)
            'delCom.ExecuteNonQuery()
            'db.Close()

            ''變更成表單為送出狀態
            'db.Open()
            'Dim strComUpd As New SqlCommand("UPDATE flowctl SET hddate = getdate(),gonogo = '0' WHERE (eformsn = '" & eformsn & "') AND stepsid = '" & nextstep & "'", db)
            'strComUpd.ExecuteNonQuery()
            'db.Close()

            '變更成表單狀態
            db.Open()
            Dim strComUpd As New SqlCommand("UPDATE flowctl SET hddate = getdate(),gonogo = '0' WHERE (eformsn = '" & eformsn & "') AND (hddate IS NULL) ", db)
            strComUpd.ExecuteNonQuery()
            db.Close()

            '找出原本表單Table
            Dim EformTable As String = F_EformTable(eformid, conn).ToString

            '填入各個表單流程結束欄位
            db.Open()
            Dim strEformUpd As New SqlCommand("UPDATE " & EformTable & " SET PENDFLAG = '0' where EFORMSN ='" & eformsn & "'", db)
            strEformUpd.ExecuteNonQuery()
            db.Close()

            Dim chinese_name As String = ""
            Dim empemail As String = ""
            Dim empuid As String = ""

            '找出填表者id
            db.Open()
            Dim RdvId As Object = New SqlCommand("select empuid from flowctl where eformsn = '" & eformsn & "' AND eformid = '" & eformid & "' AND eformrole = 1 AND gonogo='-'", db).ExecuteReader()
            If RdvId.Read() Then
                empuid = RdvId.Item("empuid").ToString()
            End If
            db.Close()

            '找出填表者資料
            db.Open()
            Dim Rdv As Object = New SqlCommand("select emp_chinese_name,empemail from employee where employee_id = '" & empuid & "'", db).ExecuteReader()
            If Rdv.Read() Then
                chinese_name = Rdv.Item("emp_chinese_name").ToString()
                empemail = Rdv.Item("empemail").ToString()
            End If
            db.Close()

            '找出發送MAIL基本資訊
            '網站伺服器位置
            Dim MOAServer As String = F_MailBase("MOAServer", conn).ToString()

            'SmtpHost
            Dim SmtpHost As String = F_MailBase("SmtpHost", conn).ToString()

            '系統郵件發送者
            Dim SystemMail As String = F_MailBase("SystemMail", conn).ToString()

            '是否寄送EMail
            Dim MailYN As String = F_MailBase("Mail_Flag", conn).ToString()

            '發送Mail給申請人
            Dim MailBody As String = ""
            MailBody += "您的" & streformName & "已經被駁回" & "<br>"
            MailBody += " <a href='" & MOAServer & strPath & "'>" & streformName & "</a>"

            '判斷是否寄送Mail
            If MailYN = "Y" Then
                'F_MailGO(SystemMail, "系統通知", SmtpHost, empemail, "表單駁回通知", MailBody)
                If eformid = "BL7U2QP3IG" Then  ''需寄信給申請人與資訊官
                    'Dim empemail As String = NowFlow.SignerWithGroupInFlow("單位資訊官", 2).empemail
                    If empemail.Length > 0 Then
                        F_MailGO(SystemMail, "系統通知", SmtpHost, NowFlow.SignerWithGroupInFlow("單位資訊官", 2).empemail, "表單駁回通知", MailBody)
                    End If
                End If
            End If

        Catch ex As Exception
            F_Back = ex.Message
        End Try
        Return F_Back
    End Function

    ''' <summary>
    ''' 駁回前幾關
    ''' </summary>
    ''' <param name="FormArr"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_BackM(ByVal FormArr As String, ByVal conn As String, ByVal num As Int32)

        F_BackM = ""

        'eformid,employee_id,eformsn,eformrole
        Dim FormList = Split(FormArr, ",")
        Dim eformid, employee_id, eformsn, eformrole As String
        Dim backNum As Integer = num
        Dim backi As Integer = 1
        Dim strBack As String = ""

        Try
            eformid = FormList(0)
            employee_id = FormList(1)
            eformsn = FormList(2)
            eformrole = FormList(3)
            'backNum = CType(FormList(4), Integer)
            backNum = num

            Do While (backi <= backNum)

                If backi = 1 Then
                    strBack = CType(F_BackSel(eformid & "," & employee_id & "," & eformsn & "," & eformrole, conn), String)
                Else
                    strBack = CType(F_BackSel(strBack, conn), String)
                End If
                backi = backi + 1
            Loop

            F_BackM = strBack

        Catch ex As Exception
            F_BackM = ex.Message
        End Try

    End Function

    ''' <summary>
    ''' 駁回前幾關
    ''' </summary>
    ''' <param name="FormArr"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_BackM(ByVal FormArr As String, ByVal conn As String) As String

        F_BackM = ""

        'eformid,employee_id,eformsn,eformrole
        Dim FormList = Split(FormArr, ",")
        Dim eformid, employee_id, eformsn, eformrole As String
        Dim backNum As Integer = 1
        Dim backi As Integer = 1
        Dim strBack As String = ""

        Try
            eformid = FormList(0)
            employee_id = FormList(1)
            eformsn = FormList(2)
            eformrole = FormList(3)
            backNum = CType(FormList(4), Integer)

            Do While (backi <= backNum)

                If backi = 1 Then
                    strBack = CType(F_BackSel(eformid & "," & employee_id & "," & eformsn & "," & eformrole, conn), String)
                Else
                    strBack = CType(F_BackSel(strBack, conn), String)
                End If

            Loop

            F_BackM = strBack

        Catch ex As Exception
            F_BackM = ex.Message
        End Try

    End Function

    ''' <summary>
    ''' 駁回關卡
    ''' </summary>
    ''' <param name="FormArr"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_BackSel(ByVal FormArr As String, ByVal conn As String)

        F_BackSel = ""

        'eformid,employee_id,eformsn,eformrole
        Dim FormList = Split(FormArr, ",")
        Dim eformid, employee_id, eformsn, eformrole As String
        Dim upstepid As Integer = 0
        Dim upflowsnOne As Integer = 0
        Dim upempuidOne As String = ""
        Dim strSql As String
        Dim stepid As String

        Try
            eformid = FormList(0)
            employee_id = FormList(1)
            eformsn = FormList(2)
            eformrole = FormList(3)

            '開啟連線
            Dim db As New SqlConnection(conn)

            '找出現在哪一關

            db.Open()
            Dim stepsidCom As New SqlCommand("select stepsid from flowctl where eformid = '" & eformid & "' AND eformsn = '" & eformsn & "' AND eformrole = '" & eformrole & "' AND empuid = '" & employee_id & "' and hddate IS NULL", db)
            Dim RdvstepID = stepsidCom.ExecuteReader()
            If RdvstepID.Read() Then
                upstepid = CType(RdvstepID.Item("stepsid"), Integer)
            End If
            db.Close()

            '變更成表單狀態
            stepid = getstepsid(eformsn, employee_id, conn)
            db.Open()
            If stepid = "21027" Then
                strSql = "UPDATE flowctl SET hddate = getdate(),comment='不修繕',gonogo = 'N' WHERE (eformsn = '" & eformsn & "') AND (hddate IS NULL) "
            ElseIf stepid = "21079" Or stepid = "1182" Then
                strSql = "UPDATE flowctl SET hddate = getdate(),gonogo = '0' WHERE (eformsn = '" & eformsn & "') AND (hddate IS NULL) "
            ElseIf eformid = "F9MBD7O97G" Then
                strSql = "UPDATE flowctl SET hddate = getdate(),comment='駁回',gonogo = '0' WHERE (eformsn = '" & eformsn & "') AND (hddate IS NULL) "
            Else
                strSql = "UPDATE flowctl SET hddate = getdate(),comment='退件，'+comment,gonogo = '0' WHERE (eformsn = '" & eformsn & "') AND (hddate IS NULL) "
            End If
            Dim strComUpdM As New SqlCommand(strSql, db)
            strComUpdM.ExecuteNonQuery()
            db.Close()

            '找出上一關

            db.Open()
            Dim strupCom As New SqlCommand("select flowsn,empuid from flowctl where eformid = '" & eformid & "' AND eformsn = '" & eformsn & "' AND eformrole = '" & eformrole & "' AND nextstep = '" & upstepid & "'", db)
            Dim RdvupID = strupCom.ExecuteReader()
            If RdvupID.Read() Then
                upflowsnOne = CType(RdvupID.Item("flowsn"), Integer)
                upempuidOne = CType(RdvupID.Item("empuid"), String)
            End If
            db.Close()

            '新增FlowCTL關聯基本資料
            db.Open()
            Dim insComNext As New SqlCommand("INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,gonogo,nextstep,goback,nextuser,stepsid,orgname,important,recdate,is_testmode,appdate,deptcode,createdate) Select eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,'?',nextstep,0,'',stepsid,orgname,important,recdate,0,getdate(),deptcode,getdate() FROM flowctl AS flowctl_1 WHERE (flowctl_1.flowsn = " & upflowsnOne & ")", db)
            insComNext.ExecuteNonQuery()
            db.Close()

            Dim BackFlowctl As New flowctl(upflowsnOne)
            'F_BackSel = eformid & "," & upempuidOne & "," & eformsn & "," & eformrole
            If (Not IsNothing(BackFlowctl)) Then F_BackSel = eformid & "," & eformsn & "," & "1," & upstepid & "," & BackFlowctl.stepsid & "," & "表單成功送給-" & BackFlowctl.emp_chinese_name & "," & eformrole
        Catch ex As Exception
            F_BackSel = ex.Message
        End Try

    End Function

    ''' <summary>
    ''' 撤銷
    ''' </summary>
    ''' <param name="FormArr"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_DrawBack(ByVal FormArr As String, ByVal conn As String)

        F_DrawBack = ""

        'eformid,eformsn
        Dim FormList = Split(FormArr, ",")
        Dim eformid, eformsn As String
        Dim strPath As String = ""
        Dim streformName As String = ""

        Try
            eformid = FormList(0)
            eformsn = FormList(1)

            If eformid = "YAqBTxRP8P" Then       '請假申請單
                streformName = "請假申請單"
                strPath = "MOA00020.aspx?x=MOA01001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "4rM2YFP73N" Then   '會議室申請單
                streformName = "會議室申請單"
                strPath = "MOA00020.aspx?x=MOA02001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "j2mvKYe3l9" Then   '派車申請單
                streformName = "派車申請單"
                strPath = "MOA00020.aspx?x=MOA03001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "61TY3LELYT" Then   '房舍水電申請單
                streformName = "房舍水電申請單"
                strPath = "MOA00020.aspx?x=MOA04001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "F9MBD7O97G" Then   '房舍水電申請單(博)
                streformName = "房舍水電申請單(博)"
                strPath = "MOA00020.aspx?x=MOA04001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "4ZXNVRV8B6" Then   '完工報告單
                streformName = "完工報告單"
                strPath = "MOA00020.aspx?x=MOA04003&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "U28r13D6EA" Then   '會客洽公申請單
                streformName = "會客洽公申請單"
                strPath = "MOA00020.aspx?x=MOA05001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "D6Y95Y5XSU" Then   '資訊設備媒體攜出入申請單
                streformName = "資訊設備媒體攜出入申請單"
                strPath = "MOA00020.aspx?x=MOA06001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "9JKSDRR5V3" Then   '報修申請單
                streformName = "報修申請單"
                strPath = "MOA00020.aspx?x=MOA07001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "5D82872F5L" Then   '銷假申請單
                streformName = "銷假申請單"
                strPath = "MOA00020.aspx?x=MOA01003&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "74BN58683M" Then   '影印使用申請單
                streformName = "影印使用申請單"
                strPath = "MOA00020.aspx?x=MOA08001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
                'ElseIf eformid = "BL7U2QP3IG" Then   '資訊設備維修申請單
                '    streformName = "資訊設備維修申請單"
                '    strPath = "MOA00020.aspx?x=MOA11001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑     
            End If

            '開啟連線
            Dim db As New SqlConnection(conn)

            '判斷是否表單有批核過
            Dim FlowCount As Integer = 0
            db.Open()
            Dim flowCountCom As New SqlCommand("select count(*) as FlowCount from flowctl where eformsn = '" & eformsn & "' AND eformid = '" & eformid & "' AND eformrole = 1 ", db)
            Dim RdvflowCount = flowCountCom.ExecuteReader()
            If RdvflowCount.Read() Then
                FlowCount = CInt(RdvflowCount.Item("FlowCount"))
            End If
            db.Close()

            '已批核過不可撤銷
            If FlowCount > 2 Then

                F_DrawBack = "1"        '批核過不可撤銷

            Else

                '變更成表單為撤銷狀態
                db.Open()
                Dim strComUpd As New SqlCommand("UPDATE flowctl SET hddate = GETDATE(),gonogo = 'B' WHERE (eformsn = '" & eformsn & "')  AND (hddate IS NULL) ", db)
                strComUpd.ExecuteNonQuery()
                db.Close()

                ''20140217 新增log紀錄
                'If (eformid = "S9QR2W8X6U" Or eformid = "U28r13D6EA") Then
                '    Call WriteAgentRecord(eformid, HttpContext.Current.Session("user_id").ToString(), db)
                '    Call WriteFlowRecord(eformid, "UPDATE flowctl SET hddate = GETDATE(),gonogo = 'B' WHERE (eformsn = '" & eformsn & "')  AND (hddate IS NULL) ")
                'End If


                '找出原本表單Table
                Dim EformTable As String
                EformTable = CType(F_EformTable(eformid, conn), String)

                '填入各個表單流程結束欄位
                db.Open()
                Dim strEformUpd As New SqlCommand("UPDATE " & EformTable & " SET PENDFLAG = 'B' where EFORMSN ='" & eformsn & "'", db)
                strEformUpd.ExecuteNonQuery()
                db.Close()

                Dim chinese_name As String = ""
                Dim empemail As String = ""
                Dim empuid As String = ""

                '找出填表者id
                db.Open()
                Dim peridCom As New SqlCommand("select empuid from flowctl where eformsn = '" & eformsn & "' AND eformid = '" & eformid & "' AND eformrole = 1 AND gonogo='-'", db)
                Dim RdvId = peridCom.ExecuteReader()
                If RdvId.Read() Then
                    empuid = CType(RdvId.Item("empuid"), String)
                End If
                db.Close()

                '找出填表者資料
                db.Open()
                Dim perCom As New SqlCommand("select emp_chinese_name,empemail from employee where employee_id = '" & empuid & "'", db)
                Dim Rdv = perCom.ExecuteReader()
                If Rdv.Read() Then
                    chinese_name = CType(Rdv.Item("emp_chinese_name"), String)
                    empemail = CType(Rdv.Item("empemail"), String)
                End If
                db.Close()

                '找出發送MAIL基本資訊
                '網站伺服器位置
                Dim MOAServer As String = ""
                MOAServer = CType(F_MailBase("MOAServer", conn), String)

                'SmtpHost
                Dim SmtpHost As String = ""
                SmtpHost = CType(F_MailBase("SmtpHost", conn), String)

                '系統郵件發送者
                Dim SystemMail As String = ""
                SystemMail = CType(F_MailBase("SystemMail", conn), String)

                '是否寄送EMail
                Dim MailYN As String = ""
                MailYN = CType(F_MailBase("Mail_Flag", conn), String)

                '發送Mail給申請人
                Dim MailBody As String = ""
                MailBody += "您的" & streformName & "已經被撤銷" & "<br>"
                MailBody += " <a href='" & MOAServer & strPath & "'>" & streformName & "</a>"

                '判斷是否寄送Mail
                If MailYN = "Y" Then
                    F_MailGO(SystemMail, "系統通知", SmtpHost, empemail, "表單撤銷通知", MailBody)
                End If

                F_DrawBack = "表單撤銷完成"

            End If



        Catch ex As Exception
            F_DrawBack = ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 找尋表單資料表
    ''' </summary>
    ''' <param name="eformid"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_EformTable(ByVal eformid As String, ByVal conn As String)
        Try
            F_EformTable = ""

            '開啟連線
            Dim db As New SqlConnection(conn)

            '找出表單資料表
            db.Open()
            Dim stepcom As New SqlCommand("select PRIMARY_TABLE from EFORMS where eformid = '" & eformid & "'", db)
            Dim steprd = stepcom.ExecuteReader()
            If steprd.Read() Then
                F_EformTable = steprd.Item("PRIMARY_TABLE")          '表單資料表
            End If
            db.Close()

        Catch ex As Exception
            F_EformTable = ex.Message
        End Try

    End Function

    ''' <summary>
    ''' 找尋關卡群組人員
    ''' </summary>
    ''' <param name="group_id"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_GroupEmp(ByVal group_id As String, ByVal conn As String) As Object
        Try
            '開啟連線
            Dim db As New SqlConnection(conn)

            '判斷該關卡有何人
            db.Open()
            Dim StepSQL As String = "SELECT employee_id,object_name,object_type FROM SYSTEMOBJ s ,SYSTEMOBJUSE u where s.object_uid = u.object_uid and u.object_uid = '" & group_id & "'"
            Dim ds As New DataSet
            Dim da As SqlDataAdapter = New SqlDataAdapter(StepSQL, db)
            da.Fill(ds)
            db.Close()
            Return ds.Tables(0)

        Catch ex As Exception
            'F_GroupEmp = ex.Message
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' 找尋關卡群組人員特定人員
    ''' </summary>
    ''' <param name="group_id"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_GroupEmp(ByVal group_id As String, ByVal conn As String, ByVal superiorid As String) As Object
        Try
            '開啟連線
            Dim db As New SqlConnection(conn)

            '判斷該關卡有特定群組人
            db.Open()
            Dim StepSQL As String = "SELECT employee_id,object_name,object_type FROM SYSTEMOBJ s ,SYSTEMOBJUSE u where s.object_uid = u.object_uid and u.object_uid = '" & group_id & "' and u.employee_id='" & superiorid & "'"
            Dim ds As New DataSet
            Dim da As SqlDataAdapter = New SqlDataAdapter(StepSQL, db)
            da.Fill(ds)
            db.Close()
            Return ds.Tables(0)

        Catch ex As Exception
            'F_GroupEmp = ex.Message
            Throw New Exception(ex.Message)
        End Try
    End Function
    ''' <summary>
    ''' 跳簽
    ''' </summary>
    ''' <param name="FormArr"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_Jump(ByVal FormArr As String, ByVal conn As String)
        F_Jump = conn
    End Function

    ''' <summary>
    ''' 寄送MAIL基本資訊
    ''' </summary>
    ''' <param name="CONFIG_VAR"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_MailBase(ByVal CONFIG_VAR As String, ByVal conn As String)
        Try
            F_MailBase = ""

            '開啟連線
            Dim db As New SqlConnection(conn)

            '找出MAIL基本資訊
            db.Open()
            Dim MailCom As New SqlCommand("select CONFIG_VALUE from SYSCONFIG WHERE CONFIG_VAR = '" & CONFIG_VAR & "'", db)
            Dim MailRdv = MailCom.ExecuteReader()
            If MailRdv.Read() Then
                F_MailBase = MailRdv("CONFIG_VALUE")
            End If
            db.Close()

        Catch ex As Exception
            F_MailBase = ex.Message
        End Try

    End Function

    ''' <summary>
    ''' 寄送MAIL給相關人員
    ''' </summary>
    ''' <param name="StrsendFrom"></param>
    ''' <param name="StrsendFromWhom"></param>
    ''' <param name="Strhost"></param>
    ''' <param name="Struser"></param>
    ''' <param name="Strsubject"></param>
    ''' <param name="Strbody"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_MailGO(ByVal StrsendFrom As String, ByVal StrsendFromWhom As String, ByVal Strhost As String, ByVal Struser As String, ByVal Strsubject As String, ByVal Strbody As String)

        '寄送Mail
        Dim oMail As New C_Mail.C_Mail

        '寄件人()
        oMail.sendFrom = StrsendFrom '寄件人Mail
        oMail.sendFromWhom = StrsendFromWhom '寄件者名稱
        oMail.host = Strhost 'SMTPHost
        oMail.to.Add(Struser) '收件人Mail
        oMail.subject = Strsubject '主旨
        oMail.body = Strbody '內容
        oMail.send()

        F_MailGO = ""

    End Function

    Function F_MailGO(ByVal StrsendFrom As String, ByVal StrsendFromWhom As String, ByVal Strhost As String, ByVal Struser As String(), ByVal Strsubject As String, ByVal Strbody As String)

        '寄送Mail
        Dim oMail As New C_Mail.C_Mail

        '寄件人()
        oMail.sendFrom = StrsendFrom '寄件人Mail
        oMail.sendFromWhom = StrsendFromWhom '寄件者名稱
        oMail.host = Strhost 'SMTPHost
        oMail.to.AddRange(Struser) '收件人Mail
        oMail.subject = Strsubject '主旨
        oMail.body = Strbody '內容
        oMail.send()

        F_MailGO = ""

    End Function
    ''' <summary>
    ''' 使用者是否為表單呈送者的上一級主管
    ''' </summary>
    ''' <param name="FormArr"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_NextChief(ByVal FormArr As String, ByVal conn As String)

        Dim FFT As New C_FlowFT.C_FlowFT
        F_NextChief = FFT.F_NextChief(FormArr, conn)

    End Function

    ''' <summary>
    ''' 使用者下一關批核長官多少人
    ''' </summary>
    ''' <param name="FormArr"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_NextStep(ByVal FormArr As String, ByVal conn As String)
        F_NextStep = ""
        'eformid,employee_id,eformsn,eformrole
        Dim FormList = Split(FormArr, ",")
        Dim eformid, employee_id, eformsn, eformrole As String
        Dim NewEform As Integer = 0
        Dim Nextobject_uid As String = ""
        Dim NextStepName As String = ""
        Dim Nextobject_type As Integer
        Dim PerCount As Integer = -1
        Dim Nowdeptcode As String = ""      '批核長官單位
        Dim Nowgroup_name As String = ""    '批核長官現在關卡
        Dim NowGroupID As String = ""      '批核長官簽核群組
        Dim Nownextstep As String = ""      '下一關關卡號碼
        Dim NowNextobject_name As String = ""   '下一關關卡名稱
        Dim NowNextobject_id As String = ""  '下一關關卡ID

        Try

            eformid = FormList(0)
            employee_id = FormList(1)
            eformsn = FormList(2)
            eformrole = FormList(3)

            '開啟連線
            Dim db As New SqlConnection(conn)

            '判斷是否為新表單
            db.Open()
            Dim NeweformCom As New SqlCommand("select flowsn from flowctl where eformsn = '" & eformsn & "'", db)
            Dim Rdv = NeweformCom.ExecuteReader()
            If Rdv.Read() Then
                NewEform = 1
            End If
            db.Close()
            Select Case NewEform
                Case 0
                    '新表單
                    db.Open()
                    Dim NewStepRdv = New SqlCommand("select object_uid,object_name,object_type from flow,SYSTEMOBJ where group_id=object_uid and eformid = '" & eformid & "' and eformrole = '" & eformrole & "' ORDER BY flow.y DESC ", db).ExecuteReader()
                    If NewStepRdv.Read() Then
                        Nextobject_uid = NewStepRdv("object_uid").ToString()
                        NextStepName = NewStepRdv("object_name").ToString()
                        Nextobject_type = CType(NewStepRdv("object_type"), Integer)
                    End If
                    db.Close()

                    Select Case Nextobject_type
                        Case 2 '上一級主管
                            '上一級主管多少人
                            db.Open()
                            Dim PerCountCom As New SqlCommand("select count(empuid) as PerCount from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = (select ORG_UID from EMPLOYEE where employee_id ='" & employee_id & "'))", db)
                            Dim PerRdv = PerCountCom.ExecuteReader()
                            If PerRdv.Read() Then
                                PerCount = CType(PerRdv("PerCount"), Integer)
                            End If
                            db.Close()
                        Case 1, 3, 4 '1:全關卡,3:同單位,4:指定群組
                            PerCount = 1
                        Case Else
                            '其他批核長官多少人
                            'db.Open()
                            'Dim PerCountCom As New SqlCommand("select count(object_num) as PerCount from SYSTEMOBJUSE where object_uid ='" & Nextobject_uid & "'", db)
                            'Dim PerRdv = PerCountCom.ExecuteReader()
                            'If PerRdv.read() Then
                            '    PerCount = PerRdv("PerCount")
                            'End If
                            'db.Close()
                    End Select
                Case 1
                    '審核中表單
                    db.Open()
                    Dim NewStepRdv = New SqlCommand("select deptcode,group_name,nextstep from flowctl where eformsn = '" & eformsn & "' and empuid = '" & employee_id & "' and hddate IS NULL ", db).ExecuteReader()
                    If NewStepRdv.Read() Then
                        Nowdeptcode = CType(NewStepRdv("deptcode"), String)
                        Nowgroup_name = CType(NewStepRdv("group_name"), String)
                        Nownextstep = CType(NewStepRdv("nextstep"), String)
                    End If
                    db.Close()

                    '判斷審核的下一關關卡是否為上一級主管
                    db.Open()
                    Dim NextStepCom As New SqlCommand("select s.object_name,s.object_uid from flow f,SYSTEMOBJ s where f.group_id = s.object_uid and steps = '" & Nownextstep & "'", db)
                    Dim NextStepRdv = NextStepCom.ExecuteReader()
                    If NextStepRdv.Read() Then
                        NowNextobject_name = CType(NextStepRdv("object_name"), String)
                        NowNextobject_id = CType(NextStepRdv("object_uid"), String)
                    End If
                    db.Close()

                    'If Nowgroup_name = "上一級主管" Or Nowobject_name = "上一級主管" Then
                    If NowNextobject_name = "上一級主管" Then
                        '上一級主管多少人
                        db.Open()
                        Dim PerCountCom As New SqlCommand("select count(empuid) as PerCount from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = '" & Nowdeptcode & "')", db)
                        Dim PerRdv = PerCountCom.ExecuteReader()
                        If PerRdv.Read() Then
                            PerCount = CType(PerRdv("PerCount"), Integer)
                        End If
                        db.Close()
                    Else
                        If eformid = "BL7U2QP3IG" Then
                            ''資訊報修管制
                            Dim strSql As String = "select group_id from flow,SYSTEMOBJ where group_id=object_uid and object_uid='" + NowNextobject_id + "'and eformid = '" & eformid & "' and eformrole = '1' ORDER BY flow.y DESC "
                            Dim DC As New SQLDBControl
                            Dim DR As SqlDataReader

                            DR = DC.CreateReader(strSql)
                            If DR.HasRows Then
                                If DR.Read() Then
                                    '下一關簽核內簽核主管多少人
                                    db.Open()
                                    Dim PerCountCom As New SqlCommand(CType(("SELECT COUNT(*) AS PerCount FROM SYSTEMOBJUSE WHERE OBJECT_UID='" & DR("group_id") & "'"), String), db)
                                    Dim PerRdv = PerCountCom.ExecuteReader()
                                    If PerRdv.Read() Then
                                        PerCount = CType(IIf("".Equals(PerRdv("PerCount")), 1, PerRdv("PerCount")), Integer)
                                    End If
                                    db.Close()
                                End If
                            End If
                            DC.Dispose()
                        ElseIf eformid = "U28r13D6EA" Or eformid = "S9QR2W8X6U" Then
                            Dim strSql As String = "select group_id from flow,SYSTEMOBJ where group_id=object_uid and eformid = '" & eformid & "' and eformrole = '1' ORDER BY flow.y DESC "
                            db.Open()
                            Dim dt As New DataTable
                            dt.Load(New SqlCommand(strSql, db).ExecuteReader)
                            db.Close()
                            For Each row In dt.Rows
                                ''全關卡判斷管制單位簽核人，通常管制單位只有1人簽核

                                '管制單位內主管多少人
                                db.Open()
                                Dim PerCountCom As New SqlCommand(CType(("SELECT COUNT(*) AS PerCount FROM SYSTEMOBJUSE WHERE OBJECT_UID='" & row("group_id") & "'"), String), db)
                                Dim PerRdv = PerCountCom.ExecuteReader()
                                If PerRdv.Read() Then
                                    PerCount = CType(IIf("".Equals(PerRdv("PerCount")), 1, PerRdv("PerCount")), Integer)
                                End If
                                db.Close()

                            Next

                        Else
                            PerCount = 1
                        End If
                    End If
            End Select
            F_NextStep = PerCount
        Catch ex As Exception
            F_NextStep = ex.Message
        End Try

    End Function
    ''' <summary>
    ''' 使用者下一關批核長官多少人
    ''' </summary>
    ''' <param name="FormArr"></param>
    ''' <param name="conn"></param>
    ''' <param name="strgroup_id"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_NextStep(ByVal FormArr As String, ByVal conn As String, ByRef strgroup_id As String) As Object
        F_NextStep = ""
        'eformid,employee_id,eformsn,eformrole
        Dim FormList = Split(FormArr, ",")
        Dim eformid, employee_id, eformsn, eformrole As String
        Dim NewEform As Integer = 0
        Dim Nextobject_uid As String = ""
        Dim NextStepName As String = ""
        Dim Nextobject_type As Integer
        Dim PerCount As Integer = -1
        Dim Nowdeptcode As String = ""      '批核長官單位
        Dim Nowgroup_name As String = ""    '批核長官現在關卡
        Dim NowGroupID As String = ""      '批核長官簽核群組
        Dim Nownextstep As String = ""      '下一關關卡號碼
        Dim Nowobject_name As String = ""   '下一關關卡名稱

        Try

            eformid = FormList(0)
            employee_id = FormList(1)
            eformsn = FormList(2)
            eformrole = FormList(3)

            '開啟連線
            Dim db As New SqlConnection(conn)

            '判斷是否為新表單
            db.Open()
            Dim NeweformCom As New SqlCommand("select flowsn from flowctl where eformsn = '" & eformsn & "'", db)
            Dim Rdv = NeweformCom.ExecuteReader()
            If Rdv.Read() Then
                NewEform = 1
            End If
            db.Close()
            Select Case NewEform
                Case 0
                    '新表單
                    db.Open()
                    Dim NewStepRdv = New SqlCommand("select object_uid,object_name,object_type from flow,SYSTEMOBJ where group_id=object_uid and eformid = '" & eformid & "' and eformrole = '" & eformrole & "' ORDER BY flow.y DESC ", db).ExecuteReader()
                    If NewStepRdv.Read() Then
                        Nextobject_uid = NewStepRdv("object_uid").ToString()
                        NextStepName = NewStepRdv("object_name").ToString()
                        Nextobject_type = CType(NewStepRdv("object_type"), Integer)
                    End If
                    db.Close()

                    Select Case Nextobject_type
                        Case 2 '上一級主管
                            '上一級組織主管多少人
                            db.Open()
                            Dim PerCountCom As New SqlCommand("select count(empuid) as PerCount from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = (select ORG_UID from EMPLOYEE where employee_id ='" & employee_id & "'))", db)
                            Dim PerRdv = PerCountCom.ExecuteReader()
                            If PerRdv.Read() Then
                                PerCount = CType(PerRdv("PerCount"), Integer)
                            End If
                            db.Close()
                        Case 1, 3, 4 '1:全關卡,3:同單位,4:指定群組
                            PerCount = 1
                        Case Else
                            '其他批核長官多少人
                            'db.Open()
                            'Dim PerCountCom As New SqlCommand("select count(object_num) as PerCount from SYSTEMOBJUSE where object_uid ='" & Nextobject_uid & "'", db)
                            'Dim PerRdv = PerCountCom.ExecuteReader()
                            'If PerRdv.read() Then
                            '    PerCount = PerRdv("PerCount")
                            'End If
                            'db.Close()
                    End Select
                Case 1
                    '審核中表單
                    db.Open()
                    Dim NewStepRdv = New SqlCommand("SELECT A.DEPTCODE,A.GROUP_NAME,A.NEXTSTEP,B.OBJECT_UID AS GROUP_ID FROM FLOWCTL A,SYSTEMOBJ B WHERE A.EFORMSN = '" & eformsn & "' AND A.EMPUID = '" & employee_id & "' AND A.HDDATE IS NULL AND A.GROUP_NAME=B.OBJECT_NAME AND B.OBJECT_TYPE='" & Nextobject_type & "'", db).ExecuteReader()
                    If NewStepRdv.Read() Then
                        Nowdeptcode = CType(NewStepRdv("deptcode"), String)
                        NowGroupID = CType(NewStepRdv("object_uid"), String)
                        Nowgroup_name = CType(NewStepRdv("group_name"), String)
                        Nownextstep = CType(NewStepRdv("nextstep"), String)
                    End If
                    db.Close()

                    '判斷審核的下一關關卡是否為上一級主管
                    db.Open()
                    Dim NextStepCom As New SqlCommand("select s.object_name from flow f,SYSTEMOBJ s where f.group_id = s.object_uid and steps = '" & Nownextstep & "'", db)
                    Dim NextStepRdv = NextStepCom.ExecuteReader()
                    If NextStepRdv.Read() Then
                        Nowobject_name = CType(NextStepRdv("object_name"), String)
                    End If
                    db.Close()

                    If Nowgroup_name = "上一級主管" Or Nowobject_name = "上一級主管" Then
                        '上一級組織主管多少人
                        db.Open()
                        Dim PerCountCom As New SqlCommand("select count(empuid) as PerCount from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = '" & Nowdeptcode & "')", db)
                        Dim PerRdv = PerCountCom.ExecuteReader()
                        If PerRdv.Read() Then
                            PerCount = CType(PerRdv("PerCount"), Integer)
                        End If
                        db.Close()
                    Else
                        If eformid = "U28r13D6EA" Or eformid = "S9QR2W8X6U" Then
                            Dim strSql As String = "select group_id from flow,SYSTEMOBJ where group_id=object_uid and eformid = '" & eformid & "' and eformrole = '1' ORDER BY flow.y DESC "
                            db.Open()
                            Dim dt As New DataTable
                            dt.Load(New SqlCommand(strSql, db).ExecuteReader)
                            db.Close()
                            For Each row In dt.Rows
                                ''全關卡判斷管制單位簽核人，通常管制單位只有1人簽核

                                '審核關卡內多少人
                                db.Open()
                                Dim PerCountCom As New SqlCommand(CType(("SELECT COUNT(*) AS PerCount FROM SYSTEMOBJUSE WHERE OBJECT_UID='" & row("group_id") & "' AND EMPLOYEE_ID IS NOT NULL"), String), db)
                                Dim PerRdv = PerCountCom.ExecuteReader()
                                If PerRdv.Read() Then
                                    PerCount = CType(IIf("".Equals(PerRdv("PerCount")), 1, PerRdv("PerCount")), Integer)
                                End If
                                db.Close()

                            Next

                        Else
                            PerCount = 1
                        End If
                    End If
            End Select
            F_NextStep = PerCount
        Catch ex As Exception
            F_NextStep = ex.Message
        End Try

    End Function
    ''' <summary>
    ''' 送件
    ''' </summary>
    ''' <param name="FormArr"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_Send(ByVal FormArr As String, ByVal conn As String) As Object
        Dim doorAndMeetingControlID As String = GetEFormId("門禁會議管制申請單") '取得門禁會議管制申請單ID
        F_Send = ""

        'eformid,employee_id,eformsn,eformrole,senduser_id
        Dim FormList = Split(FormArr, ",")
        Dim eformid, employee_id, eformsn, eformrole, senduser_id As String
        Dim eformFlag As String = ""
        Dim chinese_name, member_uid, org_uid, title_id, empemail As String
        Dim streformName As String = ""
        Dim strPath As String = ""
        Dim strPathRead As String = ""
        Dim stepid As String
        Dim strSql As String

        Try
            eformid = FormList(0)
            employee_id = FormList(1)
            eformsn = FormList(2)
            eformrole = FormList(3)
            senduser_id = FormList(4)

            Dim tool As New C_Public

            chinese_name = ""
            member_uid = ""
            org_uid = ""
            title_id = ""
            empemail = ""

            If eformid = "YAqBTxRP8P" Then       '請假申請單
                streformName = "請假申請單"
                strPath = "MOA00020.aspx?x=MOA01001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
                strPathRead = "MOA00020.aspx?x=MOA01001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "4rM2YFP73N" Then   '會議室申請單
                streformName = "會議室申請單"
                strPath = "MOA00020.aspx?x=MOA02001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
                strPathRead = "MOA00020.aspx?x=MOA02001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "j2mvKYe3l9" Then   '派車申請單
                streformName = "派車申請單"
                strPath = "MOA00020.aspx?x=MOA03001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
                strPathRead = "MOA00020.aspx?x=MOA03001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "61TY3LELYT" Then   '房舍水電申請單
                streformName = "房舍水電申請單"
                strPath = "MOA00020.aspx?x=MOA04001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
                strPathRead = "MOA00020.aspx?x=MOA04001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "F9MBD7O97G" Then   '新房舍水電申請單
                streformName = "房舍水電申請單"
                strPath = "MOA00020.aspx?x=MOA04100&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
                strPathRead = "MOA00020.aspx?x=MOA04100&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "4ZXNVRV8B6" Then   '完工報告單
                streformName = "完工報告單"
                strPath = "MOA00020.aspx?x=MOA04003&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
                strPathRead = "MOA00020.aspx?x=MOA04003&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "U28r13D6EA" Then   '會客洽公申請單
                streformName = "會客洽公申請單"
                strPath = "MOA00020.aspx?x=MOA05001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
                strPathRead = "MOA00020.aspx?x=MOA05001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "D6Y95Y5XSU" Then   '資訊設備媒體攜出入申請單
                streformName = "資訊設備媒體攜出入申請單"
                strPath = "MOA00020.aspx?x=MOA06001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
                strPathRead = "MOA00020.aspx?x=MOA06001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "9JKSDRR5V3" Then   '報修申請單
                streformName = "報修申請單"
                strPath = "MOA00020.aspx?x=MOA07001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
                strPathRead = "MOA00020.aspx?x=MOA07001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "5D82872F5L" Then   '銷假申請單
                streformName = "銷假申請單"
                strPath = "MOA00020.aspx?x=MOA01003&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
                strPathRead = "MOA00020.aspx?x=MOA01003&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "74BN58683M" Then   '影印使用申請單
                streformName = "影印使用申請單"
                strPath = "MOA00020.aspx?x=MOA08001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
                strPathRead = "MOA00020.aspx?x=MOA08001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = doorAndMeetingControlID Then   '門禁會議管制申請單
                streformName = "門禁會議管制申請單"
                strPath = "MOA00020.aspx?x=MOA09001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
                strPathRead = "MOA00020.aspx?x=MOA09001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
                'ElseIf eformid = "BL7U2QP3IG" Then   '資訊設備維修申請單
                '    streformName = "資訊設備維修申請單"
                '    strPath = "MOA00020.aspx?x=MOA11001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
                '    strPathRead = "MOA00020.aspx?x=MOA11001&y=" & eformid & "&Read_Only=1&EFORMSN=" & eformsn   '表單路徑
            Else
                Exit Function
            End If

            '開啟連線
            Dim db As New SqlConnection(conn)

            '找出發送MAIL基本資訊
            '網站伺服器位置
            Dim MOAServer As String = F_MailBase("MOAServer", conn).ToString()

            'SmtpHost
            Dim SmtpHost As String = F_MailBase("SmtpHost", conn).ToString()

            '系統郵件發送者
            Dim SystemMail As String = F_MailBase("SystemMail", conn).ToString()

            '是否寄送EMail
            Dim MailYN As String = F_MailBase("Mail_Flag", conn).ToString()

            '找出填表者資料
            db.Open()
            Dim perCom As New SqlCommand("select emp_chinese_name,member_uid,org_uid,title_id,empemail from employee where employee_id = '" & employee_id & "'", db)
            Dim Rdv = perCom.ExecuteReader()
            If Rdv.Read() Then
                chinese_name = Rdv.Item("emp_chinese_name").ToString()
                member_uid = Rdv.Item("member_uid").ToString()
                org_uid = Rdv.Item("org_uid").ToString()
                'title_id = Rdv.Item("title_id")
                empemail = Rdv.Item("empemail").ToString()
            End If
            db.Close()

            '判斷是否為新表單
            db.Open()
            Dim strCom As New SqlCommand("SELECT * FROM flowctl WHERE eformsn = '" & eformsn & "'", db)
            Dim Rdr = strCom.ExecuteReader()
            If Rdr.Read() Then
                eformFlag = "1"
            End If
            db.Close()

            'eformFlag = ""　則為新表單
            If eformFlag = "" Then

                '找尋關卡
                Dim da_step As DataTable
                da_step = CType(F_Step(eformid, eformrole, "1", conn), DataTable)

                '串簽
                If da_step.Rows.Count = "1" Then

                    Dim steps, nextstep, group_id, nextnextstep As String

                    steps = da_step.Rows(0).Item(0).ToString()          '下一關stepsid
                    nextstep = da_step.Rows(0).Item(1).ToString()       '下一關steps
                    group_id = da_step.Rows(0).Item(2).ToString()       '下一關group_id
                    nextnextstep = da_step.Rows(0).Item(3).ToString()   '下下一關nextstep

                    '新增FlowCTL基本資料
                    db.Open()
                    Dim insCom As New SqlCommand("INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,hddate,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','0','" & employee_id & "',N'" & chinese_name & "','申請人',GETDATE(),'-','" & nextstep & "','1','1',GETDATE(),GETDATE(),'" & org_uid & "',GETDATE())", db)
                    insCom.ExecuteNonQuery()
                    db.Close()

                    ' ''20130822 新增log紀錄
                    'If (eformid = "S9QR2W8X6U" Or eformid = "U28r13D6EA") Then
                    '    Call WriteAgentRecord(eformid, employee_id, db)
                    '    Call WriteFlowRecord(eformid, "INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,hddate,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','0','" & employee_id & "',N'" & chinese_name & "','申請人',GETDATE(),'-','" & nextstep & "','1','1',GETDATE(),GETDATE(),'" & org_uid & "',GETDATE())")
                    'End If

                    '找出群組資料
                    Dim da_group As DataTable
                    If eformid = "BL7U2QP3IG" Then ''資訊設備維修申請單
                        Dim DC As New SQLDBControl
                        strSql = "SELECT C.EMPLOYEE_ID,B.OBJECT_NAME,B.OBJECT_TYPE FROM SYSTEMOBJUSE A LEFT JOIN SYSTEMOBJ B ON A.OBJECT_UID=B.OBJECT_UID LEFT JOIN EMPLOYEE C ON A.EMPLOYEE_ID =C.EMPLOYEE_ID WHERE B.OBJECT_UID='" & tool.GetInformationIDByName("單位資訊官") & "' AND C.EMPLOYEE_ID IN ('" + senduser_id + "')"
                        da_group = DC.CreateDataTable(strSql)
                        DC.Dispose()
                    Else
                        da_group = CType(F_GroupEmp(group_id, conn), DataTable)
                    End If


                    Dim FullName As String = GetEmployeeName(senduser_id)     '審核者的中文姓名及帳號

                    '簽核人數可能超過兩人以上
                    Dim check_per As Integer = 0

                    For check_per = 0 To da_group.Rows.Count - 1

                        Dim group_emp, object_name, object_type As String
                        object_name = da_group.Rows(check_per).Item(1).ToString()      '群組名稱
                        object_type = da_group.Rows(check_per).Item(2).ToString()      '群組類型
                        group_emp = ""                                                 '群組人員

                        Dim SendFlag, Org_One1, Org_One2 As String
                        SendFlag = ""                                       '送件Flag
                        Org_One1 = ""                                       '審核者單位
                        Org_One2 = ""                                       '關卡人員單位

                        'object_type=1(全關卡),object_type=2(上一級主管),object_type=3(同單位),object_type=4(指定群組)

                        If object_type = "1" Then

                            If senduser_id = "" Then
                                group_emp = da_group.Rows(check_per).Item(0).ToString()
                            Else
                                group_emp = senduser_id
                            End If

                            SendFlag = "1"

                        ElseIf object_type = "2" Then

                            '影印使用申請單若其申請類別為普通，則不需要走下一關卡經過長官審核
                            Dim bl_PrinterJumpNext = False '影印使用申請單使用
                            If eformid = "74BN58683M" Then
                                db.Open()
                                Dim SecurityStatus As New SqlCommand("SELECT Security_Status FROM P_08 WHERE eformsn = '" & eformsn & "'", db)
                                Dim RdrSecurityStatus = SecurityStatus.ExecuteReader()
                                If RdrSecurityStatus.Read() Then
                                    Dim ChkSecurity As String = RdrSecurityStatus.Item("Security_Status").ToString()
                                    If ChkSecurity = "1" Then
                                        bl_PrinterJumpNext = True
                                        'Chknextstep = "-1"
                                    End If
                                End If
                                db.Close()
                            End If

                            If bl_PrinterJumpNext Then
                                '為了影印使用申請普通單，要再新增一筆假裝已完成的FlowCTL資料
                                db.Open()
                                Dim insCom2 As New SqlCommand("INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,hddate,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','0','" & employee_id & "',N'" & chinese_name & "','申請人',GETDATE(),'E','-1','1','1',GETDATE(),GETDATE(),'" & org_uid & "',GETDATE())", db)
                                insCom2.ExecuteNonQuery()
                                db.Close()
                            End If

                            ''20130822 新增log紀錄
                            'If (eformid = "S9QR2W8X6U" Or eformid = "U28r13D6EA") Then
                            '    Call WriteAgentRecord(eformid, employee_id, db)
                            '    Call WriteFlowRecord(eformid, "INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,hddate,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','0','" & employee_id & "',N'" & chinese_name & "','申請人',GETDATE(),'E','-1','1','1',GETDATE(),GETDATE(),'" & org_uid & "',GETDATE())")
                            'End If

                            If bl_PrinterJumpNext = False Then
                                If senduser_id = "" Then

                                    '上一級主管則往上一個層級長官送件
                                    '找出填表者上一級單位
                                    db.Open()
                                    Dim strTopOrg As New SqlCommand("SELECT employee_id FROM EMPLOYEE WHERE ORG_UID IN (SELECT PARENT_ORG_UID FROM ADMINGROUP WHERE ORG_UID = '" & org_uid & "')", db)
                                    Dim RdTopOrg = strTopOrg.ExecuteReader()
                                    If RdTopOrg.Read() Then
                                        group_emp = RdTopOrg.Item("employee_id").ToString()
                                    End If
                                    db.Close()

                                Else
                                    '指定上一級主管
                                    group_emp = senduser_id
                                End If

                                SendFlag = "1"
                            End If



                        ElseIf object_type = "3" Then

                            If senduser_id = "" Then

                                '審核者單位
                                db.Open()
                                Dim strEmpOrg1 As New SqlCommand("SELECT ORG_UID FROM EMPLOYEE WHERE employee_id = '" & employee_id & "'", db)
                                Dim RdTopOrg1 = strEmpOrg1.ExecuteReader()
                                If RdTopOrg1.Read() Then
                                    Org_One1 = RdTopOrg1.Item("ORG_UID").ToString()
                                End If
                                db.Close()

                                '關卡人員單位
                                db.Open()
                                Dim strEmpOrg2 As New SqlCommand("SELECT ORG_UID FROM EMPLOYEE WHERE employee_id = '" & da_group.Rows(check_per).Item(0).ToString() & "'", db)
                                Dim RdTopOrg2 = strEmpOrg2.ExecuteReader()
                                If RdTopOrg2.Read() Then
                                    Org_One2 = RdTopOrg2.Item("ORG_UID").ToString()
                                End If
                                db.Close()

                                If Org_One1 = Org_One2 Then
                                    group_emp = da_group.Rows(check_per).Item(0).ToString()
                                    SendFlag = "1"
                                End If

                            Else
                                '指定送給單位行政官
                                If check_per = 0 Then
                                    group_emp = senduser_id
                                    SendFlag = "1"
                                End If
                            End If


                        ElseIf object_type = "4" Then
                            If eformid = "4rM2YFP73N" Then
                                '會議室指定群組單位人員
                                db.Open()
                                Dim strTopOrg As New SqlCommand("SELECT Owner FROM P_0201,P_0204 WHERE P_0201.MeetSn = P_0204.MeetSn AND EFORMSN = '" & eformsn & "'", db)
                                Dim RdTopOrg = strTopOrg.ExecuteReader()
                                If RdTopOrg.Read() Then
                                    group_emp = RdTopOrg.Item("Owner").ToString()
                                End If
                                db.Close()
                                '此人員為該會議室管理員
                                If group_emp = da_group.Rows(check_per).Item(0) Then
                                    SendFlag = "1"
                                End If
                            ElseIf eformid = doorAndMeetingControlID Then '門禁管制指定群組單位人員
                                group_emp = senduser_id
                                If group_emp = da_group.Rows(check_per).Item(0) Then
                                    SendFlag = "1"
                                End If
                            End If
                        ElseIf object_type = "5" Then

                            If senduser_id = "" Then
                                '關卡為申請者
                                db.Open()
                                Dim strTopOrg As New SqlCommand("SELECT empuid FROM flowctl WHERE eformsn = '" & eformsn & "' AND gonogo='-'", db)
                                Dim RdTopOrg = strTopOrg.ExecuteReader()
                                If RdTopOrg.Read() Then
                                    group_emp = RdTopOrg.Item("empuid").ToString()
                                End If
                                db.Close()
                            Else
                                group_emp = senduser_id
                            End If

                            SendFlag = "1"

                        End If

                        '寫入流程
                        If SendFlag = "1" Then

                            '找出群組人員的詳細資料
                            Dim da_admin As DataTable
                            da_admin = CType(F_AdminEmp(group_emp, conn), DataTable)

                            Dim admin_chinese_name, admin_empemail, admin_member_uid, admin_org_uid, admin_GROUP_NAME As String
                            Dim admin_agent As BasicEmployee
                            admin_chinese_name = da_admin.Rows(0).Item(0).ToString()     '中文姓名
                            admin_empemail = da_admin.Rows(0).Item(1).ToString()         'Email
                            admin_member_uid = da_admin.Rows(0).Item(2).ToString()       '身份證字號
                            admin_org_uid = da_admin.Rows(0).Item(3).ToString()          '單位代號
                            admin_GROUP_NAME = da_admin.Rows(0).Item(4).ToString()       '單位名稱
                            admin_agent = New BasicEmployee(tool.GetAgentID(admin_member_uid, DateTime.Now.ToShortDateString()))
                            FullName = admin_chinese_name & "(" & group_emp & ")"

                            '釋放DataTable
                            da_admin.Dispose()

                            '新增FlowCTL關聯基本資料
                            db.Open()
                            Dim insComNext As New SqlCommand("INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','" & nextstep & "','" & group_emp & "',N'" & admin_chinese_name & "','" & object_name & "','?', '" & nextnextstep & "','" & steps & "','1',GETDATE(),GETDATE(),'" & admin_org_uid & "',GETDATE())", db)
                            insComNext.ExecuteNonQuery()
                            db.Close()

                            ' ''20130822 新增log紀錄
                            'If (eformid = "S9QR2W8X6U" Or eformid = "U28r13D6EA") Then
                            '    Call WriteAgentRecord(eformid, employee_id, db)
                            '    Call WriteFlowRecord(eformid, "INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','" & nextstep & "','" & group_emp & "',N'" & admin_chinese_name & "','" & object_name & "','?', '" & nextnextstep & "','" & steps & "','1',GETDATE(),GETDATE(),'" & admin_org_uid & "',GETDATE())")
                            'End If

                            '發送Mail給批核長官
                            Dim MailBody As String = ""
                            MailBody += "申請人:" & chinese_name & "<br>"
                            MailBody += " <a href='" & MOAServer & strPath & "'>" & streformName & "</a>"

                            '判斷是否寄送Mail
                            If MailYN = "Y" Then
                                If admin_agent.employee_id.Length > 0 Then
                                    F_MailGO(SystemMail, "系統通知", SmtpHost, admin_empemail + "," + admin_agent.empemail, streformName, MailBody)
                                Else
                                    F_MailGO(SystemMail, "系統通知", SmtpHost, admin_empemail, streformName, MailBody)
                                End If
                            End If

                        End If
                    Next

                    '回傳值給如影隨行
                    'eformid,eformsn,gonogo,stepsid,goto_steps,sShowMsg,eformrole
                    F_Send = eformid & "," & eformsn & "," & "1," & "1," & steps & "," & "表單成功送給-" & FullName & "," & eformrole

                    '釋放DataTable
                    da_group.Dispose()

                Else
                    '會簽
                    Dim check_group As Integer = 0
                    Dim steps, nextstep, group_id, nextnextstep As String

                    For check_group = 0 To da_step.Rows.Count - 1

                        steps = da_step.Rows(check_group).Item(0).ToString()          '下一關stepsid
                        nextstep = da_step.Rows(check_group).Item(1).ToString()       '下一關steps
                        group_id = da_step.Rows(check_group).Item(2).ToString()       '下一關group_id
                        nextnextstep = da_step.Rows(check_group).Item(3).ToString()   '下下一關nextstep

                        If check_group = 0 Then
                            '新增FlowCTL基本資料
                            db.Open()
                            Dim insCom As New SqlCommand("INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,hddate,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','0','" & employee_id & "',N'" & chinese_name & "','申請人',GETDATE(),'-','" & nextstep & "','1','1',GETDATE(),GETDATE(),'" & org_uid & "',GETDATE())", db)
                            insCom.ExecuteNonQuery()
                            db.Close()
                        End If

                        ''20130822 新增log紀錄
                        'If (eformid = "S9QR2W8X6U" Or eformid = "U28r13D6EA") Then
                        '    Call WriteAgentRecord(eformid, employee_id, db)
                        '    Call WriteFlowRecord(eformid, "INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,hddate,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','0','" & employee_id & "',N'" & chinese_name & "','申請人',GETDATE(),'-','" & nextstep & "','1','1',GETDATE(),GETDATE(),'" & org_uid & "',GETDATE())")
                        'End If

                        '找出群組資料
                        Dim da_group As DataTable
                        da_group = CType(F_GroupEmp(group_id, conn), DataTable)

                        Dim FullName As String = GetEmployeeName(senduser_id)     '審核者的中文姓名及帳號

                        '簽核人數可能超過兩人以上
                        Dim check_per As Integer = 0
                        For check_per = 0 To da_group.Rows.Count - 1

                            Dim group_emp, object_name, object_type As String
                            object_name = da_group.Rows(check_per).Item(1).ToString()      '群組名稱
                            object_type = da_group.Rows(check_per).Item(2).ToString()      '群組類型
                            group_emp = ""                                      '群組人員

                            Dim SendFlag, Org_One1, Org_One2 As String
                            SendFlag = ""                                       '送件Flag
                            Org_One1 = ""                                       '審核者單位
                            Org_One2 = ""                                       '關卡人員單位

                            'object_type=1(全關卡),object_type=2(上一級主管),object_type=3(同單位),object_type=4(指定群組)

                            If object_type = "1" Then

                                If senduser_id = "" Then
                                    group_emp = da_group.Rows(check_per).Item(0).ToString()
                                Else
                                    group_emp = senduser_id
                                End If

                                SendFlag = "1"

                            ElseIf object_type = "2" Then

                                If senduser_id = "" Then

                                    '上一級主管則往上一個層級長官送件
                                    '找出填表者上一級單位
                                    db.Open()
                                    Dim strTopOrg As New SqlCommand("SELECT employee_id FROM EMPLOYEE WHERE ORG_UID IN (SELECT PARENT_ORG_UID FROM ADMINGROUP WHERE ORG_UID = '" & org_uid & "')", db)
                                    Dim RdTopOrg = strTopOrg.ExecuteReader()
                                    If RdTopOrg.Read() Then
                                        group_emp = RdTopOrg.Item("employee_id").ToString()
                                    End If
                                    db.Close()

                                Else
                                    '指定上一級主管
                                    group_emp = senduser_id
                                End If

                                SendFlag = "1"

                            ElseIf object_type = "3" Then

                                If senduser_id = "" Then

                                    '審核者單位
                                    db.Open()
                                    Dim strEmpOrg1 As New SqlCommand("SELECT ORG_UID FROM EMPLOYEE WHERE employee_id = '" & employee_id & "'", db)
                                    Dim RdTopOrg1 = strEmpOrg1.ExecuteReader()
                                    If RdTopOrg1.Read() Then
                                        Org_One1 = RdTopOrg1.Item("ORG_UID").ToString()
                                    End If
                                    db.Close()

                                    '關卡人員單位
                                    db.Open()
                                    Dim strEmpOrg2 As New SqlCommand("SELECT ORG_UID FROM EMPLOYEE WHERE employee_id = '" & da_group.Rows(check_per).Item(0).ToString() & "'", db)
                                    Dim RdTopOrg2 = strEmpOrg2.ExecuteReader()
                                    If RdTopOrg2.Read() Then
                                        Org_One2 = RdTopOrg2.Item("ORG_UID").ToString()
                                    End If
                                    db.Close()

                                    If Org_One1 = Org_One2 Then
                                        group_emp = da_group.Rows(check_per).Item(0).ToString()
                                        SendFlag = "1"
                                    End If

                                Else
                                    '指定送給單位行政官
                                    If check_per = 0 Then
                                        group_emp = senduser_id
                                        SendFlag = "1"
                                    End If
                                End If

                            ElseIf object_type = "4" Then

                                '會議室指定群組單位人員
                                db.Open()
                                Dim strTopOrg As New SqlCommand("SELECT Owner FROM P_0201,P_0204 WHERE P_0201.MeetSn = P_0204.MeetSn AND EFORMSN = '" & eformsn & "'", db)
                                Dim RdTopOrg = strTopOrg.ExecuteReader()
                                If RdTopOrg.Read() Then
                                    group_emp = RdTopOrg.Item("Owner").ToString()
                                End If
                                db.Close()

                                '此人員為該會議室管理員
                                If group_emp = da_group.Rows(check_per).Item(0) Then
                                    SendFlag = "1"
                                End If

                            ElseIf object_type = "5" Then

                                If senduser_id = "" Then
                                    '關卡為申請者
                                    db.Open()
                                    Dim strTopOrg As New SqlCommand("SELECT empuid FROM flowctl WHERE eformsn = '" & eformsn & "' AND gonogo='-'", db)
                                    Dim RdTopOrg = strTopOrg.ExecuteReader()
                                    If RdTopOrg.Read() Then
                                        group_emp = RdTopOrg.Item("empuid").ToString()
                                    End If
                                    db.Close()
                                Else
                                    group_emp = senduser_id
                                End If

                                SendFlag = "1"

                            End If

                            '寫入流程
                            If SendFlag = "1" Then

                                '找出群組人員的詳細資料
                                Dim da_admin As DataTable
                                da_admin = CType(F_AdminEmp(group_emp, conn), DataTable)

                                Dim admin_chinese_name, admin_empemail, admin_member_uid, admin_org_uid, admin_GROUP_NAME As String
                                Dim admin_agent As BasicEmployee
                                admin_chinese_name = da_admin.Rows(0).Item(0).ToString()     '中文姓名
                                admin_empemail = da_admin.Rows(0).Item(1).ToString()         'Email
                                admin_member_uid = da_admin.Rows(0).Item(2).ToString()       '身份證字號
                                admin_org_uid = da_admin.Rows(0).Item(3).ToString()          '單位代號
                                admin_GROUP_NAME = da_admin.Rows(0).Item(4).ToString()       '單位名稱
                                admin_agent = New BasicEmployee(tool.GetAgentID(admin_member_uid, DateTime.Now.ToShortDateString())) '代理人資料
                                FullName = admin_chinese_name & "(" & group_emp & ")"

                                'HttpContext.Current.Response.Write(admin_agent.employee_id)
                                'HttpContext.Current.Response.End()

                                '釋放DataTable
                                da_admin.Dispose()

                                '新增FlowCTL關聯基本資料
                                db.Open()
                                Dim insComNext As New SqlCommand("INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','" & nextstep & "','" & group_emp & "',N'" & admin_chinese_name & "','" & object_name & "','?', '" & nextnextstep & "','" & steps & "','1',GETDATE(),GETDATE(),'" & admin_org_uid & "',GETDATE())", db)
                                insComNext.ExecuteNonQuery()
                                db.Close()

                                ''20130822 新增log紀錄
                                'If (eformid = "S9QR2W8X6U" Or eformid = "U28r13D6EA") Then
                                '    Call WriteAgentRecord(eformid, employee_id, db)
                                '    Call WriteFlowRecord(eformid, "INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','" & nextstep & "','" & group_emp & "',N'" & admin_chinese_name & "','" & object_name & "','?', '" & nextnextstep & "','" & steps & "','1',GETDATE(),GETDATE(),'" & admin_org_uid & "',GETDATE())")
                                'End If

                                '發送Mail給批核長官
                                Dim MailBody As String = ""
                                MailBody += "申請人:" & chinese_name & "<br>"
                                MailBody += " <a href='" & MOAServer & strPath & "'>" & streformName & "</a>"

                                '判斷是否寄送Mail
                                If MailYN = "Y" Then
                                    If admin_agent.employee_id.Length > 0 Then
                                        F_MailGO(SystemMail, "系統通知", SmtpHost, admin_empemail + "," + admin_agent.empemail, streformName, MailBody)
                                    Else
                                        F_MailGO(SystemMail, "系統通知", SmtpHost, admin_empemail, streformName, MailBody)
                                    End If
                                End If

                            End If

                        Next

                        '回傳值給如影隨行
                        'eformid,eformsn,gonogo,stepsid,goto_steps,sShowMsg,eformrole
                        F_Send = eformid & "," & eformsn & "," & "1," & "1," & steps & "," & "表單成功送給-" & FullName & "," & eformrole

                        '釋放DataTable
                        da_group.Dispose()

                    Next

                End If

                '釋放DataTable
                da_step.Dispose()
            Else
                '舊表單
                '判斷審核者
                Dim Chkflowsn, Chkstepsid, Chksteps, Chknextstep As Integer
                Dim SignerID As String ''已選取的下一關簽核人
                ''簽核者ID若為空,帶入申請者ID
                If senduser_id.Length > 0 Then
                    SignerID = senduser_id
                Else
                    SignerID = employee_id
                End If
                ''取出現在簽核關卡
                db.Open()
                'Dim strComCheck As New SqlCommand("SELECT flowsn,stepsid,steps,nextstep FROM flowctl WHERE eformsn = '" & eformsn & "' and empuid = '" & SignerID & "' and hddate is null", db)
                Dim strComCheck As New SqlCommand("SELECT flowsn,stepsid,steps,nextstep FROM flowctl WHERE eformsn = '" & eformsn & "' and hddate is null", db)
                Dim RdrCheck = strComCheck.ExecuteReader()
                If RdrCheck.Read() Then
                    Chkflowsn = CType(RdrCheck.Item("flowsn").ToString(), Integer)
                    Chkstepsid = CType(RdrCheck.Item("stepsid").ToString(), Integer)
                    Chksteps = CType(RdrCheck.Item("steps").ToString(), Integer)
                    Chknextstep = CType(RdrCheck.Item("nextstep").ToString(), Integer)
                End If
                db.Close()



                '審核關卡是否有兩個群組(簽核人)以上簽核
                Dim StepNow As String = ""
                db.Open()
                Dim strStepNow As New SqlCommand("SELECT count(flowsn) as recordcount FROM flowctl WHERE eformsn = '" & eformsn & "' and hddate is null", db)
                Dim RdrStepNow = strStepNow.ExecuteReader()
                If RdrStepNow.Read() Then
                    StepNow = RdrStepNow.Item("recordcount").ToString()
                End If
                db.Close()

                '審核關卡為單關
                If StepNow = "1" Then

                    '下一關是否為會簽
                    Dim CheckTwo As String = ""
                    Dim da_step As DataTable
                    da_step = F_Step(eformid, eformrole, Chkstepsid, conn)
                    CheckTwo = da_step.Rows.Count.ToString()
                    'If da_step.Rows.Count = "1" Then
                    '    CheckTwo = "1"
                    'Else
                    '    CheckTwo = da_step.Rows.Count.ToString()
                    'End If

                    '下一關為最後一關流程結束
                    If Chknextstep = "-1" Then

                        '流程結束
                        db.Open()
                        Dim strComUpd As New SqlCommand("UPDATE flowctl SET hddate = GETDATE(),gonogo = 'E' where flowsn ='" & Chkflowsn & "'", db)
                        strComUpd.ExecuteNonQuery()
                        db.Close()

                        ''20130822 新增log紀錄
                        'If (eformid = "S9QR2W8X6U" Or eformid = "U28r13D6EA") Then
                        '    Call WriteAgentRecord(eformid, employee_id, db)
                        '    Call WriteFlowRecord(eformid, "UPDATE flowctl SET hddate = GETDATE(),gonogo = 'E' where flowsn ='" & Chkflowsn & "'")
                        'End If

                        '找出原本表單Table
                        Dim EformTable As String
                        EformTable = CType(F_EformTable(eformid, conn), String)

                        '填入各個表單流程結束欄位
                        db.Open()
                        Dim strEformUpd As New SqlCommand("UPDATE " & EformTable & " SET PENDFLAG = 'E' where EFORMSN ='" & eformsn & "'", db)
                        strEformUpd.ExecuteNonQuery()
                        db.Close()

                        ''會客洽公申請單及門禁會議管制申請單完成後需將資料寫入外部資料庫 2013/6/28 paul
                        ''核准後需更新外部資料庫 
                        Select Case EformTable
                            Case "P_05"
                                Try
                                    ''insert外部資料庫
                                    Using connA As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString2").ConnectionString)
                                        connA.Open()
                                        Dim trans As SqlTransaction
                                        trans = connA.BeginTransaction
                                        ''會客洽公申請單
                                        Dim P_05AData As New P_05A(eformsn)
                                        P_05AData.Insert(trans, connA)
                                        ''會客人員明細資料表
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
                                        Dim cmd As New SqlCommand("UPDATE " & EformTable & " SET nCheckDT=GetDate() WHERE EFORMSN='" & eformsn & "'", connA, tran)
                                        Dim iSuccess As Integer = cmd.ExecuteNonQuery
                                        tran.Commit()
                                        tran.Dispose()
                                    End Using
                                Catch sqlex As SqlException
                                    If (Not "-1".Equals(sqlex.Message.IndexOf("無法開啟至 SQL Server 的連接", StringComparison.Ordinal).ToString) Or Not "-1".Equals(sqlex.Message.IndexOf("登入失敗").ToString)) Then
                                        MessageBox.Show("批核完成\n門禁會議管制申請寫入門禁系統連線錯誤，請連繫系統管理員")
                                    Else
                                        MessageBox.Show("批核完成\n門禁會議管制申請單寫入門禁系統錯誤\n" & sqlex.Message)
                                    End If
                                Catch ex As Exception
                                    MessageBox.Show(ex.Message)
                                End Try
                            Case "P_09"
                                Try
                                    ''insert外部資料庫
                                    Using connA As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString2").ConnectionString)
                                        connA.Open()
                                        Dim trans As SqlTransaction
                                        trans = connA.BeginTransaction
                                        ''門禁會議管制申請單
                                        Dim P_09AData As New P_09A(eformsn)
                                        P_09AData.Insert(trans, connA)
                                        ''進出營門
                                        'Dim P_0901AData As New P_0901A(eformsn)
                                        'P_0901AData.Insert(trans, connA)
                                        ' ''上傳明細
                                        'Dim P_09UploadData As New UploadA(eformsn)
                                        'P_09UploadData.Insert(trans, connA)
                                        trans.Commit()
                                        trans.Dispose()
                                    End Using
                                    ''更新門禁欄位已匯出時間
                                    Using connA As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
                                        connA.Open()
                                        Dim tran As SqlTransaction
                                        tran = connA.BeginTransaction
                                        Dim cmd As New SqlCommand("UPDATE " & EformTable & " SET nCheckDT=GetDate() WHERE EFORMSN='" & eformsn & "'", connA, tran)
                                        Dim iSuccess As Integer = cmd.ExecuteNonQuery
                                        tran.Commit()
                                        tran.Dispose()
                                    End Using
                                Catch sqlex As SqlException
                                    If (Not "-1".Equals(sqlex.Message.IndexOf("無法開啟至 SQL Server 的連接", StringComparison.Ordinal).ToString) Or Not "-1".Equals(sqlex.Message.IndexOf("登入失敗").ToString)) Then
                                        MessageBox.Show("批核完成\n門禁會議管制申請寫入門禁系統連線錯誤，請連繫系統管理員")
                                    Else
                                        MessageBox.Show("批核完成\n門禁會議管制申請單寫入門禁系統錯誤\n" & sqlex.Message)
                                    End If
                                Catch ex As Exception
                                    MessageBox.Show(ex.Message)
                                End Try

                        End Select

                        '找出原申請者帳號以及名稱
                        Dim SendUser As String = ""
                        ''代理人資料
                        Dim admin_agent As New BasicEmployee(tool.GetAgentID(SendUser, DateTime.Now.ToShortDateString())) '代理人資料

                        db.Open()
                        Dim strComUser As New SqlCommand("SELECT empuid FROM flowctl WHERE eformsn = '" & eformsn & "' and gonogo='-'", db)
                        Dim RdrUser = strComUser.ExecuteReader()
                        If RdrUser.Read() Then
                            SendUser = RdrUser.Item("empuid").ToString()
                        End If
                        db.Close()

                        '找出原申請者Mail
                        Dim OldChinese_name As String = ""
                        Dim OldEmpemail As String = ""

                        db.Open()
                        Dim OldPerCom As New SqlCommand("select emp_chinese_name,empemail from employee where employee_id = '" & SendUser & "'", db)
                        Dim RdvOld = OldPerCom.ExecuteReader()
                        If RdvOld.Read() Then
                            OldChinese_name = RdvOld.Item("emp_chinese_name").ToString()
                            OldEmpemail = RdvOld.Item("empemail").ToString()
                        End If
                        db.Close()

                        '發送Mail給填表者
                        Dim MailBody As String = ""
                        MailBody += "申請人:" & OldChinese_name & "<br>"
                        MailBody += " <a href='" & MOAServer & strPathRead & "'>" & streformName & "</a>"

                        '判斷是否寄送Mail
                        If MailYN = "Y" Then
                            F_MailGO(SystemMail, "系統通知", SmtpHost, OldEmpemail, streformName & "批核完成", MailBody)
                        End If

                        '回傳值給如影隨行
                        'eformid,eformsn,gonogo,stepsid,goto_steps,sShowMsg,eformrole
                        F_Send = eformid & "," & eformsn & "," & "1," & Chkstepsid & ",-1," & "成功," & eformrole

                    Else

                        '找尋關卡
                        Dim steps, nextstep, group_id, nextnextstep, bypass As String

                        '下一關是否為會簽
                        If CheckTwo = 1 Then

                            '繼續審核
                            stepid = getstepsid(eformsn, employee_id, conn)
                            Dim PresentFlowctl As New flowctl(eformid, eformsn, "?")

                            db.Open()
                            If stepid = "21027" Or stepid = "14968" Then '修繕
                                strSql = "UPDATE flowctl SET hddate = GETDATE(),gonogo = 'F' where flowsn ='" & Chkflowsn & "'"
                            ElseIf stepid = "21080" Then '完工
                                strSql = "UPDATE flowctl SET hddate = GETDATE(),gonogo = 'C' where flowsn ='" & Chkflowsn & "'"
                            ElseIf PresentFlowctl.PreStep().group_name = "資訊維修單位" And PresentFlowctl.PreStep().PreStep().group_name = "資訊報修管制單位" Then ''資訊設備維修同意修繕
                                strSql = "UPDATE flowctl SET hddate = GETDATE(),gonogo = 'F' where flowsn ='" & Chkflowsn & "'"
                            Else
                                strSql = "UPDATE flowctl SET hddate = GETDATE(),gonogo = '1' where flowsn ='" & Chkflowsn & "'"
                            End If
                            Dim strComUpd As New SqlCommand(strSql, db)
                            strComUpd.ExecuteNonQuery()
                            db.Close()

                            ''20130822 新增log紀錄
                            'If (eformid = "S9QR2W8X6U" Or eformid = "U28r13D6EA") Then
                            '    Call WriteAgentRecord(eformid, employee_id, db)
                            '    Call WriteFlowRecord(eformid, strSql)
                            'End If

                            '找尋關卡
                            steps = da_step.Rows(0).Item(0).ToString()          '下一關stepsid
                            nextstep = da_step.Rows(0).Item(1).ToString()       '下一關steps
                            group_id = da_step.Rows(0).Item(2).ToString()       '下一關group_id
                            nextnextstep = da_step.Rows(0).Item(3).ToString()   '下下一關nextstep

                            '找出群組資料
                            Dim da_group As DataTable
                            If senduser_id <> "" AndAlso group_id <> "Z860" AndAlso eformid = "BL7U2QP3IG" Then ''因不知是否會造成其他表單流程影響，暫時只針對資訊設備維修使用這樣方式
                                da_group = CType(F_GroupEmp(group_id, conn, senduser_id), DataTable)
                            Else
                                da_group = CType(F_GroupEmp(group_id, conn), DataTable)
                            End If


                            Dim FullName As String = GetEmployeeName(senduser_id)     '審核者的中文姓名及帳號

                            '簽核人數可能超過兩人以上
                            Dim check_per As Integer = 0
                            For check_per = 0 To da_group.Rows.Count - 1

                                Dim group_emp, object_name, object_type As String
                                object_name = da_group.Rows(check_per).Item(1).ToString()      '群組名稱
                                object_type = da_group.Rows(check_per).Item(2).ToString()      '群組類型
                                group_emp = ""                                                 '群組人員

                                Dim SendFlag, Org_One1, Org_One2 As String
                                SendFlag = ""                                       '送件Flag
                                Org_One1 = ""                                       '審核者單位
                                Org_One2 = ""                                       '關卡人員單位

                                'object_type=1(全關卡),object_type=2(上一級主管),object_type=3(同單位),object_type=4(指定群組)

                                If object_type = "1" Then

                                    If senduser_id = "" Then
                                        group_emp = da_group.Rows(check_per).Item(0).ToString()
                                    Else
                                        group_emp = senduser_id
                                    End If

                                    SendFlag = "1"

                                ElseIf object_type = "2" Then

                                    If senduser_id = "" Then

                                        '上一級主管則往上一個層級長官送件
                                        '找出填表者上一級單位
                                        db.Open()
                                        Dim strTopOrg As New SqlCommand("SELECT employee_id FROM EMPLOYEE WHERE ORG_UID IN (SELECT PARENT_ORG_UID FROM ADMINGROUP WHERE ORG_UID = '" & org_uid & "')", db)
                                        Dim RdTopOrg = strTopOrg.ExecuteReader()
                                        If RdTopOrg.Read() Then
                                            group_emp = RdTopOrg.Item("employee_id").ToString()
                                        End If
                                        db.Close()

                                    Else
                                        '指定上一級主管
                                        group_emp = senduser_id
                                    End If

                                    SendFlag = "1"

                                ElseIf object_type = "3" Then

                                    If senduser_id = "" Then

                                        '審核者單位
                                        db.Open()
                                        Dim strEmpOrg1 As New SqlCommand("SELECT ORG_UID FROM EMPLOYEE WHERE employee_id = '" & employee_id & "'", db)
                                        Dim RdTopOrg1 = strEmpOrg1.ExecuteReader()
                                        If RdTopOrg1.Read() Then
                                            Org_One1 = RdTopOrg1.Item("ORG_UID").ToString()
                                        End If
                                        db.Close()

                                        '關卡人員單位
                                        db.Open()
                                        Dim strEmpOrg2 As New SqlCommand("SELECT ORG_UID FROM EMPLOYEE WHERE employee_id = '" & da_group.Rows(check_per).Item(0).ToString() & "'", db)
                                        Dim RdTopOrg2 = strEmpOrg2.ExecuteReader()
                                        If RdTopOrg2.Read() Then
                                            Org_One2 = RdTopOrg2.Item("ORG_UID").ToString()
                                        End If
                                        db.Close()

                                        If Org_One1 = Org_One2 Then
                                            group_emp = da_group.Rows(check_per).Item(0).ToString()
                                            SendFlag = "1"
                                        End If

                                    Else
                                        '指定送給單位行政官
                                        If check_per = 0 Then
                                            group_emp = senduser_id
                                            SendFlag = "1"
                                        End If
                                    End If

                                ElseIf object_type = "4" Then

                                    If eformid = "4rM2YFP73N" Then

                                        '會議室指定群組單位人員
                                        db.Open()
                                        Dim strTopOrg As New SqlCommand("SELECT Owner FROM P_0201,P_0204 WHERE P_0201.MeetSn = P_0204.MeetSn AND EFORMSN = '" & eformsn & "'", db)
                                        Dim RdTopOrg = strTopOrg.ExecuteReader()
                                        If RdTopOrg.Read() Then
                                            group_emp = RdTopOrg.Item("Owner").ToString()
                                        End If
                                        db.Close()

                                        '此人員為該會議室管理員
                                        If group_emp = da_group.Rows(check_per).Item(0) Then
                                            SendFlag = "1"
                                        End If

                                    End If

                                ElseIf object_type = "5" Then

                                    If senduser_id = "" Then
                                        '關卡為申請者
                                        db.Open()
                                        Dim strTopOrg As New SqlCommand("SELECT empuid FROM flowctl WHERE eformsn = '" & eformsn & "' AND gonogo='-'", db)
                                        Dim RdTopOrg = strTopOrg.ExecuteReader()
                                        If RdTopOrg.Read() Then
                                            group_emp = RdTopOrg.Item("empuid").ToString()
                                        End If
                                        db.Close()
                                    Else
                                        group_emp = senduser_id
                                    End If

                                    SendFlag = "1"

                                End If

                                '寫入流程
                                If SendFlag = "1" Then

                                    '找出群組人員的詳細資料
                                    Dim da_admin As DataTable
                                    da_admin = CType(F_AdminEmp(group_emp, conn), DataTable)

                                    Dim admin_chinese_name, admin_empemail, admin_member_uid, admin_org_uid, admin_GROUP_NAME As String
                                    admin_chinese_name = da_admin.Rows(0).Item(0).ToString()     '中文姓名
                                    admin_empemail = da_admin.Rows(0).Item(1).ToString()         'Email
                                    admin_member_uid = da_admin.Rows(0).Item(2).ToString()       '身份證字號
                                    admin_org_uid = da_admin.Rows(0).Item(3).ToString()          '單位代號
                                    admin_GROUP_NAME = da_admin.Rows(0).Item(4).ToString()       '單位名稱
                                    Dim admin_agent As New BasicEmployee(tool.GetAgentID(admin_member_uid, DateTime.Now.ToShortDateString())) '代理人資料

                                    FullName = admin_chinese_name & "(" & group_emp & ")"

                                    '釋放DataTable
                                    da_admin.Dispose()

                                    '新增FlowCTL關聯基本資料
                                    db.Open()
                                    Dim insComNext As New SqlCommand("INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','" & nextstep & "','" & group_emp & "',N'" & admin_chinese_name & "','" & object_name & "','?', '" & nextnextstep & "','" & steps & "','1',GETDATE(),GETDATE(),'" & admin_org_uid & "',GETDATE())", db)
                                    insComNext.ExecuteNonQuery()
                                    db.Close()

                                    ''20130822 新增log紀錄
                                    'If (eformid = "S9QR2W8X6U" Or eformid = "U28r13D6EA") Then
                                    '    Call WriteAgentRecord(eformid, employee_id, db)
                                    '    Call WriteFlowRecord(eformid, "INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','" & nextstep & "','" & group_emp & "',N'" & admin_chinese_name & "','" & object_name & "','?', '" & nextnextstep & "','" & steps & "','1',GETDATE(),GETDATE(),'" & admin_org_uid & "',GETDATE())")
                                    'End If

                                    '發送Mail給批核長官
                                    Dim MailBody As String = ""
                                    MailBody += "申請人:" & chinese_name & "<br>"
                                    MailBody += " <a href='" & MOAServer & strPath & "'>" & streformName & "</a>"

                                    '判斷是否寄送Mail
                                    If MailYN = "Y" Then
                                        If admin_agent.employee_id.Length > 0 Then
                                            F_MailGO(SystemMail, "系統通知", SmtpHost, admin_empemail + "," + admin_agent.empemail, streformName, MailBody)
                                        Else
                                            F_MailGO(SystemMail, "系統通知", SmtpHost, admin_empemail, streformName, MailBody)
                                        End If
                                    End If
                                End If
                                If group_id = "E38K865C8O" Then Exit For
                            Next

                            '回傳值給如影隨行
                            'eformid,eformsn,gonogo,stepsid,goto_steps,sShowMsg,eformrole
                            F_Send = eformid & "," & eformsn & "," & "1," & Chkstepsid & "," & Chknextstep & "," & "表單成功送給-" & FullName & "," & eformrole

                            '釋放DataTable
                            da_group.Dispose()

                        ElseIf CheckTwo > 1 Then

                            Dim check_group As Integer = 0
                            '多關審核
                            For check_group = 0 To CheckTwo - 1

                                '繼續審核
                                db.Open()
                                Dim strComUpd As New SqlCommand("UPDATE flowctl SET hddate = GETDATE(),gonogo = '1' where flowsn ='" & Chkflowsn & "'", db)
                                strComUpd.ExecuteNonQuery()
                                db.Close()

                                ''20130822 新增log紀錄
                                'If (eformid = "S9QR2W8X6U" Or eformid = "U28r13D6EA") Then
                                '    Call WriteAgentRecord(eformid, employee_id, db)
                                '    Call WriteFlowRecord(eformid, "UPDATE flowctl SET hddate = GETDATE(),gonogo = '1' where flowsn ='" & Chkflowsn & "'")
                                'End If

                                steps = da_step.Rows(check_group).Item(0).ToString()          '下一關stepsid
                                nextstep = da_step.Rows(check_group).Item(1).ToString()       '下一關steps
                                group_id = da_step.Rows(check_group).Item(2).ToString()       '下一關group_id
                                nextnextstep = da_step.Rows(check_group).Item(3).ToString()   '下下一關nextstep
                                bypass = da_step.Rows(check_group).Item(4).ToString()         '0送件給該關卡

                                Dim RoomFlag As String = ""

                                '傳送選擇的會客室入口
                                If eformid = "U28r13D6EA" Then

                                    Dim strRoomName As String = ""

                                    '會客室指定群組
                                    db.Open()
                                    Dim strRoom As New SqlCommand("SELECT nRECROOM FROM P_05 WHERE EFORMSN = '" & eformsn & "'", db)
                                    Dim RdRoom = strRoom.ExecuteReader()
                                    If RdRoom.Read() Then
                                        strRoomName = RdRoom.Item("nRECROOM").ToString()
                                    End If
                                    db.Close()

                                    RoomFlag = CheckRoom(strRoomName, group_id, conn)
                                    'If strRoomName = "第一會客室" And group_id = "E539" Then
                                    '    RoomFlag = ""
                                    'ElseIf strRoomName = "第二會客室" And group_id = "Y965" Then
                                    '    RoomFlag = ""
                                    'Else
                                    '    RoomFlag = "1"
                                    'End If

                                End If

                                If bypass = 0 And RoomFlag = "" Then

                                    '找出群組資料
                                    Dim da_group As DataTable
                                    da_group = CType(F_GroupEmp(group_id, conn), DataTable)

                                    Dim FullName As String = GetEmployeeName(senduser_id)     '審核者的中文姓名及帳號

                                    '簽核人數可能超過兩人以上
                                    Dim check_per As Integer = 0
                                    For check_per = 0 To da_group.Rows.Count - 1

                                        Dim group_emp, object_name, object_type As String
                                        object_name = da_group.Rows(check_per).Item(1).ToString()      '群組名稱
                                        object_type = da_group.Rows(check_per).Item(2).ToString()      '群組類型
                                        group_emp = ""                                      '群組人員

                                        Dim SendFlag, Org_One1, Org_One2 As String
                                        SendFlag = ""                                       '送件Flag
                                        Org_One1 = ""                                       '審核者單位
                                        Org_One2 = ""                                       '關卡人員單位

                                        'object_type=1(全關卡),object_type=2(上一級主管),object_type=3(同單位),object_type=4(指定群組)

                                        If object_type = "1" Then

                                            If senduser_id = "" Then
                                                group_emp = da_group.Rows(check_per).Item(0).ToString()
                                            Else
                                                group_emp = senduser_id
                                            End If

                                            SendFlag = "1"

                                        ElseIf object_type = "2" Then

                                            If senduser_id = "" Then

                                                '上一級主管則往上一個層級長官送件
                                                '找出填表者上一級單位
                                                db.Open()
                                                Dim strTopOrg As New SqlCommand("SELECT employee_id FROM EMPLOYEE WHERE ORG_UID IN (SELECT PARENT_ORG_UID FROM ADMINGROUP WHERE ORG_UID = '" & org_uid & "')", db)
                                                Dim RdTopOrg = strTopOrg.ExecuteReader()
                                                If RdTopOrg.Read() Then
                                                    group_emp = RdTopOrg.Item("employee_id").ToString()
                                                End If
                                                db.Close()

                                            Else
                                                '指定上一級主管
                                                group_emp = senduser_id
                                            End If

                                            SendFlag = "1"

                                        ElseIf object_type = "3" Then

                                            If senduser_id = "" Then

                                                '審核者單位
                                                db.Open()
                                                Dim strEmpOrg1 As New SqlCommand("SELECT ORG_UID FROM EMPLOYEE WHERE employee_id = '" & employee_id & "'", db)
                                                Dim RdTopOrg1 = strEmpOrg1.ExecuteReader()
                                                If RdTopOrg1.Read() Then
                                                    Org_One1 = RdTopOrg1.Item("ORG_UID").ToString()
                                                End If
                                                db.Close()

                                                '關卡人員單位
                                                db.Open()
                                                Dim strEmpOrg2 As New SqlCommand("SELECT ORG_UID FROM EMPLOYEE WHERE employee_id = '" & da_group.Rows(check_per).Item(0).ToString() & "'", db)
                                                Dim RdTopOrg2 = strEmpOrg2.ExecuteReader()
                                                If RdTopOrg2.Read() Then
                                                    Org_One2 = RdTopOrg2.Item("ORG_UID").ToString()
                                                End If
                                                db.Close()

                                                If Org_One1 = Org_One2 Then
                                                    group_emp = da_group.Rows(check_per).Item(0).ToString()
                                                    SendFlag = "1"
                                                End If

                                            Else
                                                '指定送給單位行政官
                                                If check_per = 0 Then
                                                    group_emp = senduser_id
                                                    SendFlag = "1"
                                                End If
                                            End If

                                        ElseIf object_type = "4" Then

                                            If eformid = "4rM2YFP73N" Then

                                                '會議室指定群組單位人員
                                                db.Open()
                                                Dim strTopOrg As New SqlCommand("SELECT Owner FROM P_0201,P_0204 WHERE P_0201.MeetSn = P_0204.MeetSn AND EFORMSN = '" & eformsn & "'", db)
                                                Dim RdTopOrg = strTopOrg.ExecuteReader()
                                                If RdTopOrg.Read() Then
                                                    group_emp = RdTopOrg.Item("Owner").ToString()
                                                End If
                                                db.Close()

                                                '此人員為該會議室管理員
                                                If group_emp = da_group.Rows(check_per).Item(0) Then
                                                    SendFlag = "1"
                                                End If

                                            End If

                                        ElseIf object_type = "5" Then

                                            If senduser_id = "" Then
                                                '關卡為申請者
                                                db.Open()
                                                Dim strTopOrg As New SqlCommand("SELECT empuid FROM flowctl WHERE eformsn = '" & eformsn & "' AND gonogo='-'", db)
                                                Dim RdTopOrg = strTopOrg.ExecuteReader()
                                                If RdTopOrg.Read() Then
                                                    group_emp = RdTopOrg.Item("empuid").ToString()
                                                End If
                                                db.Close()
                                            Else
                                                group_emp = senduser_id
                                            End If

                                            SendFlag = "1"

                                        End If

                                        '寫入流程
                                        If SendFlag = "1" Then

                                            '找出群組人員的詳細資料
                                            Dim da_admin As DataTable
                                            da_admin = CType(F_AdminEmp(group_emp, conn), DataTable)

                                            Dim admin_chinese_name, admin_empemail, admin_member_uid, admin_org_uid, admin_GROUP_NAME As String
                                            admin_chinese_name = da_admin.Rows(0).Item(0).ToString()     '中文姓名
                                            admin_empemail = da_admin.Rows(0).Item(1).ToString()         'Email
                                            admin_member_uid = da_admin.Rows(0).Item(2).ToString()       '身份證字號
                                            admin_org_uid = da_admin.Rows(0).Item(3).ToString()          '單位代號
                                            admin_GROUP_NAME = da_admin.Rows(0).Item(4).ToString()       '單位名稱
                                            Dim admin_agent As New BasicEmployee(tool.GetAgentID(admin_member_uid, DateTime.Now.ToShortDateString())) '代理人資料

                                            FullName = admin_chinese_name & "(" & group_emp & ")"

                                            '釋放DataTable
                                            da_admin.Dispose()

                                            '新增FlowCTL關聯基本資料
                                            db.Open()
                                            Dim insComNext As New SqlCommand("INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','" & nextstep & "','" & group_emp & "',N'" & admin_chinese_name & "','" & object_name & "','?', '" & nextnextstep & "','" & steps & "','1',GETDATE(),GETDATE(),'" & admin_org_uid & "',GETDATE())", db)
                                            insComNext.ExecuteNonQuery()
                                            db.Close()

                                            ''20130822 新增log紀錄
                                            'If (eformid = "S9QR2W8X6U" Or eformid = "U28r13D6EA") Then
                                            '    Call WriteAgentRecord(eformid, employee_id, db)
                                            '    Call WriteFlowRecord(eformid, "INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','" & nextstep & "','" & group_emp & "',N'" & admin_chinese_name & "','" & object_name & "','?', '" & nextnextstep & "','" & steps & "','1',GETDATE(),GETDATE(),'" & admin_org_uid & "',GETDATE())")
                                            'End If

                                            '發送Mail給批核長官
                                            Dim MailBody As String = ""
                                            MailBody += "申請人:" & chinese_name & "<br>"
                                            MailBody += " <a href='" & MOAServer & strPath & "'>" & streformName & "</a>"

                                            '判斷是否寄送Mail
                                            If MailYN = "Y" Then
                                                If admin_agent.employee_id.Length > 0 Then
                                                    F_MailGO(SystemMail, "系統通知", SmtpHost, admin_empemail + "," + admin_agent.empemail, streformName, MailBody)
                                                Else
                                                    F_MailGO(SystemMail, "系統通知", SmtpHost, admin_empemail, streformName, MailBody)
                                                End If
                                            End If

                                        End If

                                    Next

                                    '回傳值給如影隨行
                                    'eformid,eformsn,gonogo,stepsid,goto_steps,sShowMsg,eformrole
                                    F_Send = eformid & "," & eformsn & "," & "1," & Chkstepsid & "," & Chknextstep & "," & "表單成功送給-" & FullName & "," & eformrole

                                    '釋放DataTable
                                    da_group.Dispose()

                                End If

                            Next



                        Else
                            F_Send = "關卡無人簽核"
                        End If

                    End If

                Else
                    '審核者為多關
                    '下一關是否為會簽
                    Dim CheckTwo As String = ""
                    Dim da_step As DataTable
                    da_step = F_Step(eformid, eformrole, Chkstepsid, conn)
                    If da_step.Rows.Count = "1" Then
                        CheckTwo = "1"
                    Else
                        CheckTwo = da_step.Rows.Count.ToString()
                    End If

                    '下一關為最後一關流程結束
                    If Chknextstep = "-1" Then

                        '流程結束
                        db.Open()
                        Dim strComUpd As New SqlCommand("UPDATE flowctl SET hddate = GETDATE(),gonogo = 'E' where eformsn ='" & eformsn & "' AND steps = '" & Chksteps & "'", db)
                        strComUpd.ExecuteNonQuery()
                        db.Close()

                        ''20130822 新增log紀錄
                        'If (eformid = "S9QR2W8X6U" Or eformid = "U28r13D6EA") Then
                        '    Call WriteAgentRecord(eformid, employee_id, db)
                        '    Call WriteFlowRecord(eformid, "UPDATE flowctl SET hddate = GETDATE(),gonogo = 'E' where eformsn ='" & eformsn & "' AND steps = '" & Chksteps & "'")
                        'End If

                        '找出原本表單Table
                        Dim EformTable As String
                        EformTable = CType(F_EformTable(eformid, conn), String)

                        '填入各個表單流程結束欄位
                        db.Open()
                        Dim strEformUpd As New SqlCommand("UPDATE " & EformTable & " SET PENDFLAG = 'E' where EFORMSN ='" & eformsn & "'", db)
                        strEformUpd.ExecuteNonQuery()
                        db.Close()

                        '找出原申請者帳號以及名稱
                        Dim SendUser As String = ""

                        db.Open()
                        Dim strComUser As New SqlCommand("SELECT empuid FROM flowctl WHERE eformsn = '" & eformsn & "' and gonogo='-'", db)
                        Dim RdrUser = strComUser.ExecuteReader()
                        If RdrUser.Read() Then
                            SendUser = RdrUser.Item("empuid").ToString()
                        End If
                        db.Close()

                        '找出原申請者Mail
                        Dim OldChinese_name As String = ""
                        Dim OldEmpemail As String = ""

                        db.Open()
                        Dim OldPerCom As New SqlCommand("select emp_chinese_name,empemail from employee where employee_id = '" & SendUser & "'", db)
                        Dim RdvOld = OldPerCom.ExecuteReader()
                        If RdvOld.Read() Then
                            OldChinese_name = RdvOld.Item("emp_chinese_name").ToString()
                            OldEmpemail = RdvOld.Item("empemail").ToString()
                        End If
                        db.Close()

                        '發送Mail給填表者
                        Dim MailBody As String = ""
                        MailBody += "申請人:" & OldChinese_name & "<br>"
                        MailBody += " <a href='" & MOAServer & strPathRead & "'>" & streformName & "</a>"

                        '判斷是否寄送Mail
                        If MailYN = "Y" Then
                            F_MailGO(SystemMail, "系統通知", SmtpHost, OldEmpemail, streformName & "批核完成", MailBody)
                        End If

                        '回傳值給如影隨行
                        'eformid,eformsn,gonogo,stepsid,goto_steps,sShowMsg,eformrole
                        F_Send = eformid & "," & eformsn & "," & "1," & Chkstepsid & ",-1," & "成功," & eformrole


                    Else

                        '找尋關卡
                        Dim steps, nextstep, group_id, nextnextstep, bypass As String

                        '下一關是否為會簽
                        If CheckTwo = 1 Then

                            '繼續審核
                            db.Open()
                            Dim strComUpd As New SqlCommand("UPDATE flowctl SET hddate = GETDATE(),gonogo = '1' where eformsn ='" & eformsn & "' AND steps = '" & Chksteps & "'", db)
                            strComUpd.ExecuteNonQuery()
                            db.Close()

                            ''20130822 新增log紀錄
                            'If (eformid = "S9QR2W8X6U" Or eformid = "U28r13D6EA") Then
                            '    Call WriteAgentRecord(eformid, employee_id, db)
                            '    Call WriteFlowRecord(eformid, "UPDATE flowctl SET hddate = GETDATE(),gonogo = '1' where eformsn ='" & eformsn & "' AND steps = '" & Chksteps & "'")
                            'End If

                            '找尋關卡
                            steps = da_step.Rows(0).Item(0).ToString()          '下一關stepsid
                            nextstep = da_step.Rows(0).Item(1).ToString()       '下一關steps
                            group_id = da_step.Rows(0).Item(2).ToString()       '下一關group_id
                            nextnextstep = da_step.Rows(0).Item(3).ToString()   '下下一關nextstep

                            '找出群組資料
                            Dim da_group As DataTable
                            da_group = CType(F_GroupEmp(group_id, conn), DataTable)

                            Dim FullName As String = GetEmployeeName(senduser_id)     '審核者的中文姓名及帳號

                            '簽核人數可能超過兩人以上
                            Dim check_per As Integer = 0
                            For check_per = 0 To da_group.Rows.Count - 1

                                Dim group_emp, object_name, object_type As String
                                object_name = da_group.Rows(check_per).Item(1).ToString()      '群組名稱
                                object_type = da_group.Rows(check_per).Item(2).ToString()      '群組類型
                                group_emp = ""                                      '群組人員

                                Dim SendFlag, Org_One1, Org_One2 As String
                                SendFlag = ""                                       '送件Flag
                                Org_One1 = ""                                       '審核者單位
                                Org_One2 = ""                                       '關卡人員單位

                                'object_type=1(全關卡),object_type=2(上一級主管),object_type=3(同單位),object_type=4(指定群組)

                                If object_type = "1" Then

                                    If senduser_id = "" Then
                                        group_emp = da_group.Rows(check_per).Item(0).ToString()
                                    Else
                                        group_emp = senduser_id
                                    End If

                                    SendFlag = "1"

                                ElseIf object_type = "2" Then

                                    If senduser_id = "" Then

                                        '上一級主管則往上一個層級長官送件
                                        '找出填表者上一級單位
                                        db.Open()
                                        Dim strTopOrg As New SqlCommand("SELECT employee_id FROM EMPLOYEE WHERE ORG_UID IN (SELECT PARENT_ORG_UID FROM ADMINGROUP WHERE ORG_UID = '" & org_uid & "')", db)
                                        Dim RdTopOrg = strTopOrg.ExecuteReader()
                                        If RdTopOrg.Read() Then
                                            group_emp = RdTopOrg.Item("employee_id").ToString()
                                        End If
                                        db.Close()

                                    Else
                                        '指定上一級主管
                                        group_emp = senduser_id
                                    End If

                                    SendFlag = "1"

                                ElseIf object_type = "3" Then

                                    If senduser_id = "" Then

                                        '審核者單位
                                        db.Open()
                                        Dim strEmpOrg1 As New SqlCommand("SELECT ORG_UID FROM EMPLOYEE WHERE employee_id = '" & employee_id & "'", db)
                                        Dim RdTopOrg1 = strEmpOrg1.ExecuteReader()
                                        If RdTopOrg1.Read() Then
                                            Org_One1 = RdTopOrg1.Item("ORG_UID").ToString()
                                        End If
                                        db.Close()

                                        '關卡人員單位
                                        db.Open()
                                        Dim strEmpOrg2 As New SqlCommand("SELECT ORG_UID FROM EMPLOYEE WHERE employee_id = '" & da_group.Rows(check_per).Item(0).ToString() & "'", db)
                                        Dim RdTopOrg2 = strEmpOrg2.ExecuteReader()
                                        If RdTopOrg2.Read() Then
                                            Org_One2 = RdTopOrg2.Item("ORG_UID").ToString()
                                        End If
                                        db.Close()

                                        If Org_One1 = Org_One2 Then
                                            group_emp = da_group.Rows(check_per).Item(0).ToString()
                                            SendFlag = "1"
                                        End If

                                    Else
                                        '指定送給單位行政官
                                        If check_per = 0 Then
                                            group_emp = senduser_id
                                            SendFlag = "1"
                                        End If
                                    End If

                                ElseIf object_type = "4" Then

                                    If eformid = "4rM2YFP73N" Then

                                        '會議室指定群組單位人員
                                        db.Open()
                                        Dim strTopOrg As New SqlCommand("SELECT Owner FROM P_0201,P_0204 WHERE P_0201.MeetSn = P_0204.MeetSn AND EFORMSN = '" & eformsn & "'", db)
                                        Dim RdTopOrg = strTopOrg.ExecuteReader()
                                        If RdTopOrg.Read() Then
                                            group_emp = RdTopOrg.Item("Owner").ToString()
                                        End If
                                        db.Close()

                                        '此人員為該會議室管理員
                                        If group_emp = da_group.Rows(check_per).Item(0) Then
                                            SendFlag = "1"
                                        End If

                                    End If

                                ElseIf object_type = "5" Then

                                    If senduser_id = "" Then
                                        '關卡為申請者
                                        db.Open()
                                        Dim strTopOrg As New SqlCommand("SELECT empuid FROM flowctl WHERE eformsn = '" & eformsn & "' AND gonogo='-'", db)
                                        Dim RdTopOrg = strTopOrg.ExecuteReader()
                                        If RdTopOrg.Read() Then
                                            group_emp = RdTopOrg.Item("empuid").ToString()
                                        End If
                                        db.Close()
                                    Else
                                        group_emp = senduser_id
                                    End If

                                    SendFlag = "1"

                                End If

                                '寫入流程
                                If SendFlag = "1" Then

                                    '找出群組人員的詳細資料
                                    Dim da_admin As DataTable
                                    da_admin = CType(F_AdminEmp(group_emp, conn), DataTable)

                                    Dim admin_chinese_name, admin_empemail, admin_member_uid, admin_org_uid, admin_GROUP_NAME As String
                                    admin_chinese_name = da_admin.Rows(0).Item(0).ToString()     '中文姓名
                                    admin_empemail = da_admin.Rows(0).Item(1).ToString()         'Email
                                    admin_member_uid = da_admin.Rows(0).Item(2).ToString()       '身份證字號
                                    admin_org_uid = da_admin.Rows(0).Item(3).ToString()          '單位代號
                                    admin_GROUP_NAME = da_admin.Rows(0).Item(4).ToString()       '單位名稱
                                    Dim admin_agent As New BasicEmployee(tool.GetAgentID(admin_member_uid, DateTime.Now.ToShortDateString())) '代理人資料

                                    FullName = admin_chinese_name & "(" & group_emp & ")"

                                    '釋放DataTable
                                    da_admin.Dispose()

                                    '新增FlowCTL關聯基本資料
                                    db.Open()
                                    Dim insComNext As New SqlCommand("INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','" & nextstep & "','" & group_emp & "',N'" & admin_chinese_name & "','" & object_name & "','?', '" & nextnextstep & "','" & steps & "','1',GETDATE(),GETDATE(),'" & admin_org_uid & "',GETDATE())", db)
                                    insComNext.ExecuteNonQuery()
                                    db.Close()

                                    ''20130822 新增log紀錄
                                    'If (eformid = "S9QR2W8X6U" Or eformid = "U28r13D6EA") Then
                                    '    Call WriteAgentRecord(eformid, employee_id, db)
                                    '    Call WriteFlowRecord(eformid, "INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','" & nextstep & "','" & group_emp & "',N'" & admin_chinese_name & "','" & object_name & "','?', '" & nextnextstep & "','" & steps & "','1',GETDATE(),GETDATE(),'" & admin_org_uid & "',GETDATE())")
                                    'End If

                                    '發送Mail給批核長官
                                    Dim MailBody As String = ""
                                    MailBody += "申請人:" & chinese_name & "<br>"
                                    MailBody += " <a href='" & MOAServer & strPath & "'>" & streformName & "</a>"

                                    '判斷是否寄送Mail
                                    If MailYN = "Y" Then
                                        If admin_agent.employee_id.Length > 0 Then
                                            F_MailGO(SystemMail, "系統通知", SmtpHost, admin_empemail + "," + admin_agent.empemail, streformName, MailBody)
                                        Else
                                            F_MailGO(SystemMail, "系統通知", SmtpHost, admin_empemail, streformName, MailBody)
                                        End If
                                    End If

                                End If

                            Next

                            '回傳值給如影隨行
                            'eformid,eformsn,gonogo,stepsid,goto_steps,sShowMsg,eformrole
                            F_Send = eformid & "," & eformsn & "," & "1," & Chkstepsid & "," & Chknextstep & "," & "表單成功送給-" & FullName & "," & eformrole

                            '釋放DataTable
                            da_group.Dispose()

                        ElseIf CheckTwo > 1 Then

                            Dim check_group As Integer = 0
                            '多關審核
                            For check_group = 0 To CheckTwo - 1

                                '繼續審核
                                db.Open()
                                Dim strComUpd As New SqlCommand("UPDATE flowctl SET hddate = GETDATE(),gonogo = '1' where eformsn ='" & eformsn & "' AND steps = '" & Chksteps & "'", db)
                                strComUpd.ExecuteNonQuery()
                                db.Close()

                                ''20130822 新增log紀錄
                                'If (eformid = "S9QR2W8X6U" Or eformid = "U28r13D6EA") Then
                                '    Call WriteAgentRecord(eformid, employee_id, db)
                                '    Call WriteFlowRecord(eformid, "UPDATE flowctl SET hddate = GETDATE(),gonogo = '1' where eformsn ='" & eformsn & "' AND steps = '" & Chksteps & "'")
                                'End If

                                steps = da_step.Rows(check_group).Item(0).ToString()          '下一關stepsid
                                nextstep = da_step.Rows(check_group).Item(1).ToString()       '下一關steps
                                group_id = da_step.Rows(check_group).Item(2).ToString()       '下一關group_id
                                nextnextstep = da_step.Rows(check_group).Item(3).ToString()   '下下一關nextstep
                                bypass = da_step.Rows(check_group).Item(4).ToString()         '0送件給該關卡

                                If bypass = 0 Then

                                    '找出群組資料
                                    Dim da_group As DataTable
                                    da_group = CType(F_GroupEmp(group_id, conn), DataTable)

                                    Dim FullName As String = GetEmployeeName(senduser_id)     '審核者的中文姓名及帳號

                                    '簽核人數可能超過兩人以上
                                    Dim check_per As Integer = 0
                                    For check_per = 0 To da_group.Rows.Count - 1

                                        Dim group_emp, object_name, object_type As String
                                        object_name = da_group.Rows(check_per).Item(1).ToString()      '群組名稱
                                        object_type = da_group.Rows(check_per).Item(2).ToString()      '群組類型
                                        group_emp = ""                                                 '群組人員

                                        Dim SendFlag, Org_One1, Org_One2 As String
                                        SendFlag = ""                                       '送件Flag
                                        Org_One1 = ""                                       '審核者單位
                                        Org_One2 = ""                                       '關卡人員單位

                                        'object_type=1(全關卡),object_type=2(上一級主管),object_type=3(同單位),object_type=4(指定群組)

                                        If object_type = "1" Then

                                            If senduser_id = "" Then
                                                group_emp = da_group.Rows(check_per).Item(0).ToString()
                                            Else
                                                group_emp = senduser_id
                                            End If

                                            SendFlag = "1"

                                        ElseIf object_type = "2" Then

                                            If senduser_id = "" Then

                                                '上一級主管則往上一個層級長官送件
                                                '找出填表者上一級單位
                                                db.Open()
                                                Dim strTopOrg As New SqlCommand("SELECT employee_id FROM EMPLOYEE WHERE ORG_UID IN (SELECT PARENT_ORG_UID FROM ADMINGROUP WHERE ORG_UID = '" & org_uid & "')", db)
                                                Dim RdTopOrg = strTopOrg.ExecuteReader()
                                                If RdTopOrg.Read() Then
                                                    group_emp = RdTopOrg.Item("employee_id").ToString()
                                                End If
                                                db.Close()

                                            Else
                                                '指定上一級主管
                                                group_emp = senduser_id
                                            End If

                                            SendFlag = "1"

                                        ElseIf object_type = "3" Then

                                            If senduser_id = "" Then

                                                '審核者單位
                                                db.Open()
                                                Dim strEmpOrg1 As New SqlCommand("SELECT ORG_UID FROM EMPLOYEE WHERE employee_id = '" & employee_id & "'", db)
                                                Dim RdTopOrg1 = strEmpOrg1.ExecuteReader()
                                                If RdTopOrg1.Read() Then
                                                    Org_One1 = RdTopOrg1.Item("ORG_UID").ToString()
                                                End If
                                                db.Close()

                                                '關卡人員單位
                                                db.Open()
                                                Dim strEmpOrg2 As New SqlCommand("SELECT ORG_UID FROM EMPLOYEE WHERE employee_id = '" & da_group.Rows(check_per).Item(0).ToString() & "'", db)
                                                Dim RdTopOrg2 = strEmpOrg2.ExecuteReader()
                                                If RdTopOrg2.Read() Then
                                                    Org_One2 = RdTopOrg2.Item("ORG_UID").ToString()
                                                End If
                                                db.Close()

                                                If Org_One1 = Org_One2 Then
                                                    group_emp = da_group.Rows(check_per).Item(0).ToString()
                                                    SendFlag = "1"
                                                End If

                                            Else
                                                '指定送給單位行政官
                                                If check_per = 0 Then
                                                    group_emp = senduser_id
                                                    SendFlag = "1"
                                                End If
                                            End If

                                        ElseIf object_type = "4" Then

                                            If eformid = "4rM2YFP73N" Then

                                                '會議室指定群組單位人員
                                                db.Open()
                                                Dim strTopOrg As New SqlCommand("SELECT Owner FROM P_0201,P_0204 WHERE P_0201.MeetSn = P_0204.MeetSn AND EFORMSN = '" & eformsn & "'", db)
                                                Dim RdTopOrg = strTopOrg.ExecuteReader()
                                                If RdTopOrg.Read() Then
                                                    group_emp = RdTopOrg.Item("Owner").ToString()
                                                End If
                                                db.Close()

                                                '此人員為該會議室管理員
                                                If group_emp = da_group.Rows(check_per).Item(0) Then
                                                    SendFlag = "1"
                                                End If

                                            End If

                                        ElseIf object_type = "5" Then

                                            If senduser_id = "" Then
                                                '關卡為申請者
                                                db.Open()
                                                Dim strTopOrg As New SqlCommand("SELECT empuid FROM flowctl WHERE eformsn = '" & eformsn & "' AND gonogo='-'", db)
                                                Dim RdTopOrg = strTopOrg.ExecuteReader()
                                                If RdTopOrg.Read() Then
                                                    group_emp = RdTopOrg.Item("empuid").ToString()
                                                End If
                                                db.Close()
                                            Else
                                                group_emp = senduser_id
                                            End If

                                            SendFlag = "1"

                                        End If

                                        '寫入流程
                                        If SendFlag = "1" Then

                                            '找出群組人員的詳細資料
                                            Dim da_admin As DataTable
                                            da_admin = CType(F_AdminEmp(group_emp, conn), DataTable)

                                            Dim admin_chinese_name, admin_empemail, admin_member_uid, admin_org_uid, admin_GROUP_NAME As String
                                            admin_chinese_name = da_admin.Rows(0).Item(0).ToString()     '中文姓名
                                            admin_empemail = da_admin.Rows(0).Item(1).ToString()         'Email
                                            admin_member_uid = da_admin.Rows(0).Item(2).ToString()       '身份證字號
                                            admin_org_uid = da_admin.Rows(0).Item(3).ToString()          '單位代號
                                            admin_GROUP_NAME = da_admin.Rows(0).Item(4).ToString()       '單位名稱
                                            Dim admin_agent As New BasicEmployee(tool.GetAgentID(admin_member_uid, DateTime.Now.ToShortDateString())) '代理人資料

                                            FullName = admin_chinese_name & "(" & group_emp & ")"

                                            '釋放DataTable
                                            da_admin.Dispose()

                                            '新增FlowCTL關聯基本資料
                                            db.Open()
                                            Dim insComNext As New SqlCommand("INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','" & nextstep & "','" & group_emp & "',N'" & admin_chinese_name & "','" & object_name & "','?', '" & nextnextstep & "','" & steps & "','1',GETDATE(),GETDATE(),'" & admin_org_uid & "',GETDATE())", db)
                                            insComNext.ExecuteNonQuery()
                                            db.Close()

                                            ''20130822 新增log紀錄
                                            'If (eformid = "S9QR2W8X6U" Or eformid = "U28r13D6EA") Then
                                            '    Call WriteAgentRecord(eformid, employee_id, db)
                                            '    Call WriteFlowRecord(eformid, "INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','" & nextstep & "','" & group_emp & "',N'" & admin_chinese_name & "','" & object_name & "','?', '" & nextnextstep & "','" & steps & "','1',GETDATE(),GETDATE(),'" & admin_org_uid & "',GETDATE())")
                                            'End If

                                            '發送Mail給批核長官
                                            Dim MailBody As String = ""
                                            MailBody += "申請人:" & chinese_name & "<br>"
                                            MailBody += " <a href='" & MOAServer & strPath & "'>" & streformName & "</a>"

                                            '判斷是否寄送Mail
                                            If MailYN = "Y" Then
                                                If admin_agent.employee_id.Length > 0 Then
                                                    F_MailGO(SystemMail, "系統通知", SmtpHost, admin_empemail + "," + admin_agent.empemail, streformName, MailBody)
                                                Else
                                                    F_MailGO(SystemMail, "系統通知", SmtpHost, admin_empemail, streformName, MailBody)
                                                End If
                                            End If

                                        End If

                                    Next

                                    '回傳值給如影隨行
                                    'eformid,eformsn,gonogo,stepsid,goto_steps,sShowMsg,eformrole
                                    F_Send = eformid & "," & eformsn & "," & "1," & Chkstepsid & "," & Chknextstep & "," & "表單成功送給-" & FullName & "," & eformrole

                                    '釋放DataTable
                                    da_group.Dispose()


                                End If

                            Next

                        Else
                            F_Send = "關卡無人簽核"
                        End If

                    End If


                End If

            End If
        Catch ex As Exception
            F_Send = ex.Message
        End Try

    End Function
    ''' <summary>
    ''' 判斷營門資料是否有在資料庫內
    ''' </summary>
    ''' <param name="sRoomName"></param>
    ''' <param name="sGroupID"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CheckRoom(ByVal sRoomName As String, ByVal sGroupID As String, ByVal conn As String) As String
        Dim sReturn As String = "1"
        Dim strSql As String = ""
        Dim db As New SqlConnection(conn)

        strSql = "SELECT * FROM SYSKIND WHERE KIND_NUM='3' AND STATE_NAME='" + sRoomName + "'"
        db.Open()
        Dim DR As SqlDataReader = New SqlCommand(strSql, db).ExecuteReader
        If DR.HasRows Then
            While DR.Read()
                ''判斷會客入口與管制單位是否一致
                If DR("STATE_NAME") = sRoomName And DR("STATE_DESC") = sGroupID Then
                    sReturn = ""
                    Return sReturn
                End If
            End While
        End If
        db.Close()
        CheckRoom = sReturn
    End Function
    ''' <summary>
    ''' 找尋關卡
    ''' </summary>
    ''' <param name="eformid"></param>
    ''' <param name="eformrole"></param>
    ''' <param name="stepsid"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_Step(ByVal eformid As String, ByVal eformrole As String, ByVal stepsid As String, ByVal conn As String) As Object
        Try
            Dim iEformrole As Integer
            Dim iStepsid As Integer

            If (Integer.TryParse(eformrole, iEformrole)) AndAlso Integer.TryParse(stepsid, iStepsid) Then
                Return F_Step(eformid, iEformrole, iStepsid, conn)
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return Nothing
    End Function
    ''' <summary>
    ''' 找尋關卡
    ''' </summary>
    ''' <param name="eformid"></param>
    ''' <param name="eformrole"></param>
    ''' <param name="stepsid"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_Step(ByVal eformid As String, ByVal eformrole As Integer, ByVal stepsid As String, ByVal conn As String) As DataTable
        Try
            Dim iStepsid As Integer

            If Integer.TryParse(stepsid, iStepsid) Then
                Return F_Step(eformid, eformrole, iStepsid, conn)
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return Nothing
    End Function
    ''' <summary>
    ''' 找尋關卡
    ''' </summary>
    ''' <param name="eformid"></param>
    ''' <param name="eformrole"></param>
    ''' <param name="stepsid"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_Step(ByVal eformid As String, ByVal eformrole As String, ByVal stepsid As Integer, ByVal conn As String) As DataTable
        Try
            Dim iEformrole As Integer

            If Integer.TryParse(eformrole, iEformrole) Then
                Return F_Step(eformid, iEformrole, stepsid, conn)
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return Nothing
    End Function
    ''' <summary>
    ''' 找尋關卡
    ''' </summary>
    ''' <param name="eformid"></param>
    ''' <param name="eformrole"></param>
    ''' <param name="stepsid"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_Step(ByVal eformid As String, ByVal eformrole As Integer, ByVal stepsid As Integer, ByVal conn As String) As DataTable
        Try
            Dim nextstep As String = ""

            '開啟連線
            Dim db As New SqlConnection(conn)

            '找出下一關
            db.Open()
            Dim stepcom As New SqlCommand("select nextstep from flow where eformid = '" & eformid & "' and eformrole = '" & eformrole & "' and stepsid = '" & stepsid & "'", db)
            Dim steprd = stepcom.ExecuteReader()
            If steprd.Read() Then
                nextstep = steprd.Item("nextstep").ToString()          '下一關
            End If
            db.Close()

            '輸出下一關資料
            db.Open()
            Dim StepSQL As String = "select stepsid,steps,group_id,nextstep,bypass from flow where eformid = '" & eformid & "' and eformrole = '" & eformrole & "' and steps = '" & nextstep & "'"
            Dim ds As New DataSet
            Dim da As SqlDataAdapter = New SqlDataAdapter(StepSQL, db)
            da.Fill(ds)
            F_Step = ds.Tables(0)
            db.Close()

        Catch ex As Exception
            'F_Step = ex.Message
            Throw New Exception(ex.Message)
        End Try

    End Function

    ''' <summary>
    ''' 補登表單
    ''' </summary>
    ''' <param name="FormArr"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_Supply(ByVal FormArr As String, ByVal conn As String)

        F_Supply = ""

        'eformid,employee_id,eformsn,eformrole
        Dim FormList = Split(FormArr, ",")
        Dim eformid, employee_id, eformsn, eformrole As String
        Dim chinese_name, member_uid, org_uid, title_id As String

        Try

            eformid = FormList(0)
            employee_id = FormList(1)
            eformsn = FormList(2)
            eformrole = FormList(3)

            chinese_name = ""
            member_uid = ""
            org_uid = ""
            title_id = ""

            '開啟連線
            Dim db As New SqlConnection(conn)

            '找出申請者資料
            db.Open()
            Dim perCom As New SqlCommand("select emp_chinese_name,member_uid,org_uid,title_id from employee where employee_id = '" & employee_id & "'", db)
            Dim Rdv = perCom.ExecuteReader()
            If Rdv.Read() Then
                chinese_name = Rdv.Item("emp_chinese_name").ToString()
                member_uid = Rdv.Item("member_uid").ToString()
                org_uid = Rdv.Item("org_uid").ToString()
                title_id = Rdv.Item("title_id").ToString()
            End If
            db.Close()

            '新增FlowCTL基本資料
            db.Open()
            Dim insCom As New SqlCommand("INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,hddate,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','0','" & employee_id & "','" & chinese_name & "','申請人',GETDATE(),'-','-1','1','1',GETDATE(),GETDATE(),'" & org_uid & "',GETDATE())", db)
            insCom.ExecuteNonQuery()
            db.Close()

            '補登FlowCTL資料
            db.Open()
            Dim updCom As New SqlCommand("INSERT INTO flowctl(eformid,eformrole,eformsn,steps,empuid,emp_chinese_name,group_name,hddate,gonogo,nextstep,stepsid,important,recdate,appdate,deptcode,createdate) VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','0','" & employee_id & "','" & chinese_name & "','結束',GETDATE(),'G','-1','1','1',GETDATE(),GETDATE(),'" & org_uid & "',GETDATE())", db)
            updCom.ExecuteNonQuery()
            db.Close()

            '找出原本表單Table
            Dim EformTable As String
            EformTable = F_EformTable(eformid, conn).ToString()

            '填入各個表單流程結束欄位
            db.Open()
            Dim strEformUpd As New SqlCommand("UPDATE " & EformTable & " SET PENDFLAG = 'E' where EFORMSN ='" & eformsn & "'", db)
            strEformUpd.ExecuteNonQuery()
            db.Close()

            F_Supply = "補登完成"

        Catch ex As Exception
            F_Supply = ex.Message
        End Try

    End Function

    ''' <summary>
    ''' 呈轉
    ''' </summary>
    ''' <param name="FormArr"></param>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function F_Transfer(ByVal FormArr As String, ByVal conn As String)
        F_Transfer = ""

        'eformid,employee_id,eformsn,eformrole
        Dim FormList = Split(FormArr, ",")
        Dim eformid, employee_id, eformsn, eformrole As String
        Dim org_uid As String = ""
        Dim Top_empid As String = ""
        Dim Top_empname As String = ""
        Dim Top_emporg As String = ""
        Dim Top_empemail As String = ""
        Dim strTran As String = ""
        Dim D_stepsid As String = ""
        Dim D_steps As String = ""
        Dim D_emp_chinese_name As String = ""
        Dim D_nextstep As String = ""

        Try

            eformid = FormList(0)
            employee_id = FormList(1)
            eformsn = FormList(2)
            eformrole = FormList(3)

            Dim strPath As String = ""
            Dim streformName As String = ""

            If eformid = "YAqBTxRP8P" Then       '請假申請單
                streformName = "請假申請單"
                strPath = "MOA00020.aspx?x=MOA01001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "4rM2YFP73N" Then   '會議室申請單
                streformName = "會議室申請單"
                strPath = "MOA00020.aspx?x=MOA02001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "j2mvKYe3l9" Then   '派車申請單
                streformName = "派車申請單"
                strPath = "MOA00020.aspx?x=MOA03001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "61TY3LELYT" Then   '房舍水電申請單
                streformName = "房舍水電申請單"
                strPath = "MOA00020.aspx?x=MOA04001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "F9MBD7O97G" Then   '房舍水電申請單(博)
                streformName = "房舍水電申請單(博)"
                strPath = "MOA00020.aspx?x=MOA04001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "4ZXNVRV8B6" Then   '完工報告單
                streformName = "完工報告單"
                strPath = "MOA00020.aspx?x=MOA04003&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "U28r13D6EA" Then   '會客洽公申請單
                streformName = "會客洽公申請單"
                strPath = "MOA00020.aspx?x=MOA05001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "D6Y95Y5XSU" Then   '資訊設備媒體攜出入申請單
                streformName = "資訊設備媒體攜出入申請單"
                strPath = "MOA00020.aspx?x=MOA06001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "9JKSDRR5V3" Then   '報修申請單
                streformName = "報修申請單"
                strPath = "MOA00020.aspx?x=MOA07001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "5D82872F5L" Then   '銷假申請單
                streformName = "銷假申請單"
                strPath = "MOA00020.aspx?x=MOA01003&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
            ElseIf eformid = "74BN58683M" Then   '影印使用申請單
                streformName = "影印使用申請單"
                strPath = "MOA00020.aspx?x=MOA08001&y=" & eformid & "&Read_Only=2&EFORMSN=" & eformsn   '表單路徑
            End If

            '開啟連線
            Dim db As New SqlConnection(conn)

            '找出發送MAIL基本資訊
            '網站伺服器位置
            Dim MOAServer As String = ""
            MOAServer = CType(F_MailBase("MOAServer", conn), String)

            'SmtpHost
            Dim SmtpHost As String = ""
            SmtpHost = CType(F_MailBase("SmtpHost", conn), String)

            '系統郵件發送者
            Dim SystemMail As String = ""
            SystemMail = CType(F_MailBase("SystemMail", conn), String)

            '是否寄送EMail
            Dim MailYN As String = ""
            MailYN = CType(F_MailBase("Mail_Flag", conn), String)

            '取得該筆表單資料
            db.Open()
            Dim strTranDetail As New SqlCommand("SELECT * FROM flowctl WHERE (eformsn = '" & eformsn & "') AND (empuid = '" & employee_id & "')  AND (hddate IS NULL)", db)
            Dim RdTranDetail = strTranDetail.ExecuteReader()
            If RdTranDetail.Read() Then
                D_stepsid = RdTranDetail.Item("stepsid").ToString()
                D_steps = RdTranDetail.Item("steps").ToString()
                D_emp_chinese_name = RdTranDetail.Item("emp_chinese_name").ToString()
                D_nextstep = RdTranDetail.Item("nextstep").ToString()
            End If
            db.Close()

            '表單變更成呈轉狀態
            db.Open()
            Dim strComUpd As New SqlCommand("UPDATE flowctl SET hddate = GETDATE(),gonogo = 'T' WHERE (eformsn = '" & eformsn & "') AND (hddate IS NULL) ", db)
            strComUpd.ExecuteNonQuery()
            db.Close()

            ''20140217 新增log紀錄
            'If (eformid = "S9QR2W8X6U" Or eformid = "U28r13D6EA") Then
            '    Call WriteAgentRecord(eformid, employee_id, db)
            '    Call WriteFlowRecord(eformid, "UPDATE flowctl SET hddate = GETDATE(),gonogo = 'T' WHERE (eformsn = '" & eformsn & "') AND (hddate IS NULL) ")
            'End If

            '找出申請者資料
            db.Open()
            Dim perCom As New SqlCommand("select emp_chinese_name,member_uid,org_uid,title_id from employee where employee_id = '" & employee_id & "'", db)
            Dim Rdv = perCom.ExecuteReader()
            If Rdv.Read() Then
                org_uid = CType(Rdv.Item("org_uid"), String)
            End If
            db.Close()

            '取得登入者的上一級主管
            db.Open()
            Dim strTopOrg As New SqlCommand("SELECT employee_id,emp_chinese_name,ORG_UID,empemail FROM EMPLOYEE WHERE ORG_UID IN (SELECT PARENT_ORG_UID FROM ADMINGROUP WHERE ORG_UID = '" & org_uid & "')", db)
            Dim RdTopOrg = strTopOrg.ExecuteReader()
            If RdTopOrg.Read() Then
                Top_empid = CType(RdTopOrg.Item("employee_id"), String)
                Top_empname = CType(RdTopOrg.Item("emp_chinese_name"), String)
                Top_emporg = CType(RdTopOrg.Item("ORG_UID"), String)
                Top_empemail = CType(RdTopOrg.Item("empemail"), String)
            End If
            db.Close()

            '表單呈轉給上一級主管
            db.Open()
            strTran = "INSERT INTO flowctl(eformid,eformrole,eformsn,stepsid,steps,empuid,emp_chinese_name,group_name,gonogo,nextstep,important,recdate,appdate,deptcode,createdate) "
            strTran += " VALUES ('" & eformid & "','" & eformrole & "','" & eformsn & "','" & D_stepsid & "','" & D_steps & "','" & Top_empid & "','" & Top_empname & "','上一級主管','?', '" & D_nextstep & "','1',GETDATE(),GETDATE(),'" & Top_emporg & "',GETDATE()) "
            Dim strTranUpd As New SqlCommand(strTran, db)
            strTranUpd.ExecuteNonQuery()
            db.Close()

            '發送Mail給上一級主管
            Dim MailBody As String = ""
            MailBody += "批核者:" & D_emp_chinese_name & "呈轉" & streformName & "<br>"
            MailBody += " <a href='" & MOAServer & strPath & "'>" & streformName & "</a>"

            '判斷是否寄送Mail
            If MailYN = "Y" Then
                F_MailGO(SystemMail, "系統通知", SmtpHost, Top_empemail, streformName, MailBody)
            End If

            F_Transfer = "呈轉完成"

        Catch ex As Exception
            F_Transfer = ex.Message
        End Try

    End Function

    ''' <summary>
    ''' 紀錄log
    ''' </summary>
    ''' <param name="eformsn"></param>
    ''' <param name="Chkflowsn"></param>
    ''' <param name="Chkstepsid"></param>
    ''' <param name="Chksteps"></param>
    ''' <param name="Chknextstep"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function LogRecord(ByVal eformsn As String, ByVal Chkflowsn As String, ByVal Chkstepsid As String, ByVal Chksteps As String, ByVal Chknextstep As String, ByVal db As SqlConnection) As Boolean
        Dim strLog_Msg As String = ""
        Dim strRecord As String = ""
        Dim strLog_FileName As String = DateTime.Now.ToString("yyyyMMdd")

        Try

            Dim my_Dir As String = HttpContext.Current.Server.MapPath("~/Log") & "/" & DateTime.Now.ToString("yyyyMMdd")
            If Not Directory.Exists(my_Dir) Then
                Directory.CreateDirectory(my_Dir)  '--如果這目錄不存在，就建立它。
            End If


            '--每一個紀錄檔的副檔名都是 .log
            Dim LogFile As String = my_Dir & "\" & strLog_FileName + ".log"

            strLog_Msg = GetEFormNameById("")
            '--每一段訊息，都要記錄完整的日期與時間（年/月/日/時/分/秒）。
            strRecord = System.String.Format("[{0:yyyy/MM/dd hh:mm:ss}]Message : {1}", DateTime.Now, strLog_Msg.Trim())
            '--第一個參數，檔名。
            '--第二個參數，是否採用APPEND的方式？表示新的資料，會附加在檔案末端。
            '--第三個參數，正體中文編碼 System.Text.Encoding.GetEncoding("BIG5")，預設值UTF-8。
            Dim sw As StreamWriter
            sw = New StreamWriter(LogFile, True, Encoding.GetEncoding("BIG5"))
            sw.WriteLine(strRecord)
            sw.Flush()
            sw.Close()
            sw.Dispose()

            ''紀錄代理情形
            'Dim strAgentMsg = GetFlowAgentRecord(eformsn)

            '--每一段訊息，都要記錄完整的日期與時間（年/月/日/時/分/秒）。
            strRecord = System.String.Format("[{0:yyyy/MM/dd hh:mm:ss}]Message : {1}", DateTime.Now, strLog_Msg.Trim())
            '--第一個參數，檔名。
            '--第二個參數，是否採用APPEND的方式？表示新的資料，會附加在檔案末端。
            '--第三個參數，正體中文編碼 System.Text.Encoding.GetEncoding("BIG5")，預設值UTF-8。
            sw = New StreamWriter(LogFile, True, Encoding.GetEncoding("BIG5"))
            sw.WriteLine(strRecord)
            sw.Flush()
            sw.Close()
            sw.Dispose()



            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    ''' <summary>
    ''' 取得目前關卡
    ''' 的stepid
    ''' </summary>
    ''' <param name="_streformsn"></param>
    ''' <param name="_empuid"></param>
    ''' <param name="connstr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getstepsid(ByVal _streformsn As String, ByVal _empuid As String, ByVal connstr As String) As String
        Dim _strstepsid As String = ""
        Dim conn As New C_SQLFUN

        '開啟連線
        Dim db As New SqlConnection(connstr)

        db.Open()
        Dim strstepChk As New SqlCommand("select stepsid from dbo.flowctl where empuid='" & _empuid & "' and eformsn='" & _streformsn & "' and gonogo ='?'", db)
        Dim RdrstepChk = strstepChk.ExecuteReader()

        If RdrstepChk.Read() Then
            _strstepsid = CType(RdrstepChk.Item("stepsid"), String)
        End If
        db.Close()
        Return _strstepsid
    End Function

    ''' <summary>
    ''' 取得此人的
    ''' TU_Name
    ''' 級職名稱by TU_ID from [TITLES_U]
    ''' </summary>
    ''' <param name="TU_ID"></param>
    ''' <param name="connstr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getTU_Name(ByVal TU_ID As String, ByVal connstr As String) As String
        Dim _sTU_Name As String = ""
        Dim conn As New C_SQLFUN
        Try
            '開啟連線
            Dim db As New SqlConnection(connstr)
            db.Open()
            Dim strsTU_Name As New SqlCommand("select TU_Name from TITLES_U with(nolock) where TU_ID =" & TU_ID, db)
            Dim RdrTU_Name As IDataReader = strsTU_Name.ExecuteReader()

            If RdrTU_Name.Read() Then
                _sTU_Name = CType(RdrTU_Name.Item("TU_Name"), String)
            End If
            db.Close()

        Catch ex As Exception
            _sTU_Name = "error"
        End Try
        Return _sTU_Name
    End Function

    ''' <summary>
    ''' 取得此人
    ''' 的ORG_NAME
    ''' 單位名稱by ORG_UID from [AdminGroup]
    ''' </summary>
    ''' <param name="ORG_UID"></param>
    ''' <param name="connstr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getORG_Name(ByVal ORG_UID As String, ByVal connstr As String) As String
        Dim _sORG_NAME As String = ""
        Dim conn As New C_SQLFUN
        Try
            '開啟連線
            Dim db As New SqlConnection(connstr)
            db.Open()
            Dim strsORG_NAME As New SqlCommand("SELECT ORG_NAME FROM AdminGroup WITH(NOLOCK) WHERE ORG_UID ='" & ORG_UID & "'", db)
            Dim RdrORG_NAME As IDataReader = strsORG_NAME.ExecuteReader()

            If RdrORG_NAME.Read() Then
                _sORG_NAME = CType(RdrORG_NAME.Item("ORG_NAME"), String)
            End If
            db.Close()

        Catch ex As Exception
            _sORG_NAME = "error"
        End Try
        Return _sORG_NAME
    End Function

    ''' <summary>
    ''' 取得此人的
    ''' Emp_chinese_name中文名稱
    ''' by employee_id from [Employee]
    ''' </summary>
    ''' <param name="employee_id"></param>
    ''' <param name="connstr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getEmp_chinese_name(ByVal employee_id As String, ByVal connstr As String) As String
        Dim _schinese_name As String = ""
        Dim conn As New C_SQLFUN
        Try
            '開啟連線
            Dim db As New SqlConnection(connstr)
            db.Open()
            Dim strschinese_name As New SqlCommand("select emp_chinese_name from Employee with(nolock) where employee_id = '" & employee_id & "'", db)
            Dim Rdrchinese_name As IDataReader = strschinese_name.ExecuteReader()

            If Rdrchinese_name.Read() Then
                _schinese_name = CType(Rdrchinese_name.Item("emp_chinese_name"), String)
            End If
            db.Close()

        Catch ex As Exception
            _schinese_name = "error"
        End Try
        Return _schinese_name
    End Function

    ''' <summary>
    ''' 取得表單種類
    ''' 的ID
    ''' </summary>
    ''' <param name="eformName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetEFormId(ByVal eformName As String) As String
        Dim conn As New C_SQLFUN
        Dim connstr As String = conn.G_conn_string
        Dim db As New SqlConnection(connstr)
        db.Open()
        Dim dr As SqlDataReader = New SqlCommand("SELECT eformid FROM EFORMS WHERE frm_chinese_name = '" + eformName + "'", db).ExecuteReader
        If dr.Read Then
            GetEFormId = CType(dr("eformid"), String)
        Else
            GetEFormId = ""
        End If
        'GetEFormId = sqlcomm.ExecuteScalar().ToString()
        db.Close()
    End Function

    ''' <summary>
    ''' 取得員工
    ''' 的中文姓名
    ''' </summary>
    ''' <param name="employeeID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetEmployeeName(ByVal employeeID As String) As String
        Dim conn As New C_SQLFUN
        Dim connstr As String = conn.G_conn_string
        Dim db As New SqlConnection(connstr)
        db.Open()
        Dim sqlcomm As New SqlCommand("SELECT emp_chinese_name FROM EMPLOYEE WHERE employee_id = @employeeID", db)
        sqlcomm.Parameters.Add("@employeeID", SqlDbType.VarChar, 10).Value = employeeID
        Dim dr As SqlDataReader = sqlcomm.ExecuteReader
        If dr.Read Then
            GetEmployeeName = CType(dr("emp_chinese_name"), String)
        Else
            GetEmployeeName = ""
        End If
        'GetEmployeeName = sqlcomm.ExecuteScalar().ToString()
        db.Close()
    End Function

    ''' <summary>
    ''' 取得表單種類
    ''' 的Name By ID
    ''' </summary>
    ''' <param name="eformid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetEFormNameById(ByVal eformid As String) As String
        Dim conn As New C_SQLFUN
        Dim connstr As String = conn.G_conn_string
        Dim db As New SqlConnection(connstr)
        db.Open()
        Dim dr As SqlDataReader = New SqlCommand("SELECT frm_chinese_name FROM EFORMS WHERE eformid = '" + eformid + "'", db).ExecuteReader
        If dr.Read Then
            GetEFormNameById = CType(dr("frm_chinese_name"), String)
        Else
            GetEFormNameById = ""
        End If
        'GetEFormNameById = sqlcomm.ExecuteScalar().ToString()
        db.Close()
    End Function

    Public Shared Sub WriteAgentRecord(ByVal seformid As String, ByVal empid As String, ByVal db As SqlConnection)
        Dim strSql As String
        db.Open()
        ''有代理人才做紀錄
        strSql = "SELECT * FROM AGENT WHERE AGENT1='" & empid & "' OR AGENT2='" & empid & "' AND Agent_SDate>='" & DateTime.Now.ToString("yyyy/M/d H:m:s") & "' AND Agent_EDate<='" & DateTime.Now.ToString("yyyy/M/d H:m:s") & "'"
        Dim RdTopOrg = New SqlCommand((strSql), db).ExecuteReader
        Dim dtWithAgent As New DataTable
        db.Close()

        dtWithAgent.Load(RdTopOrg)

        If dtWithAgent.Rows.Count > 0 Then

            '修改成db模式
            Dim strLog_Msg As String = ""
            Dim strRecord As String = ""

            strLog_Msg = GetEFormNameById(seformid)

            '--每一段訊息，都要記錄完整的日期與時間（年/月/日/時/分/秒）。
            strRecord = System.String.Format("[{0:yyyy/MM/dd hh:mm:ss}]Message : {1}", DateTime.Now, strLog_Msg.Trim())
            '--第一個參數，檔名。
            '--第二個參數，是否採用APPEND的方式？表示新的資料，會附加在檔案末端。
            '--第三個參數，正體中文編碼 System.Text.Encoding.GetEncoding("BIG5")，預設值UTF-8。
            db.Open()
            strSql = "INSERT INTO ERRORLOG (MSG) VALUES ('" & strRecord & "')"
            Dim cmd = New SqlCommand(strSql, db).ExecuteNonQuery
            db.Close()

            ''紀錄代理情形
            'Dim strAgentMsg = GetFlowAgentRecord(eformsn)
            Dim strAgentMsg As String = "代理人：" & RdTopOrg("Agent1").ToString() & "  被代理人：" & RdTopOrg("Agent2").ToString() & "  代理區間：" & RdTopOrg("Agent_SDate").ToString() & "~" & RdTopOrg("Agent_EDate").ToString()
            '--每一段訊息，都要記錄完整的日期與時間（年/月/日/時/分/秒）。
            strRecord = System.String.Format("[{0:yyyy/MM/dd hh:mm:ss}]Message : {1}", DateTime.Now, strAgentMsg.Trim())
            '--第一個參數，檔名。
            '--第二個參數，是否採用APPEND的方式？表示新的資料，會附加在檔案末端。
            '--第三個參數，正體中文編碼 System.Text.Encoding.GetEncoding("BIG5")，預設值UTF-8。
            db.Open()
            strSql = "INSERT INTO ERRORLOG (MSG) VALUES ('" & strRecord & "')"
            'Dim cmd = New SqlCommand(strSql, db).ExecuteNonQuery
            db.Close()

            'Dim strLog_Msg As String = ""
            'Dim strRecord As String = ""
            'Dim strLog_FileName As String = DateTime.Now.ToString("yyyyMMdd") & "Agent"

            'Dim my_Dir As String = HttpContext.Current.Server.MapPath("~/Log") & "/" & DateTime.Now.ToString("yyyyMMdd")
            'If Not Directory.Exists(my_Dir) Then
            '    Directory.CreateDirectory(my_Dir)  '--如果這目錄不存在，就建立它。
            'End If


            ''--每一個紀錄檔的副檔名都是 .log
            'Dim LogFile As String = my_Dir & "\" & strLog_FileName + ".log"

            ''strLog_Msg = GetEFormNameById("74M36O86UE")
            'strLog_Msg = GetEFormNameById(seformid)

            ''--每一段訊息，都要記錄完整的日期與時間（年/月/日/時/分/秒）。
            'strRecord = System.String.Format("[{0:yyyy/MM/dd hh:mm:ss}]Message : {1}", DateTime.Now, strLog_Msg.Trim())
            ''--第一個參數，檔名。
            ''--第二個參數，是否採用APPEND的方式？表示新的資料，會附加在檔案末端。
            ''--第三個參數，正體中文編碼 System.Text.Encoding.GetEncoding("BIG5")，預設值UTF-8。
            'Dim sw As StreamWriter
            'sw = New StreamWriter(LogFile, True, Encoding.GetEncoding("BIG5"))
            'sw.WriteLine(strRecord)
            'sw.Flush()
            'sw.Close()
            'sw.Dispose()

            ' ''紀錄代理情形
            ''Dim strAgentMsg = GetFlowAgentRecord(eformsn)
            'Dim strAgentMsg As String = "代理人：" & RdTopOrg("Agent1").ToString() & "  被代理人：" & RdTopOrg("Agent2").ToString() & "  代理區間：" & RdTopOrg("Agent_SDate").ToString() & "~" & RdTopOrg("Agent_EDate").ToString()
            ''--每一段訊息，都要記錄完整的日期與時間（年/月/日/時/分/秒）。
            'strRecord = System.String.Format("[{0:yyyy/MM/dd hh:mm:ss}]Message : {1}", DateTime.Now, strAgentMsg.Trim())
            ''--第一個參數，檔名。
            ''--第二個參數，是否採用APPEND的方式？表示新的資料，會附加在檔案末端。
            ''--第三個參數，正體中文編碼 System.Text.Encoding.GetEncoding("BIG5")，預設值UTF-8。
            'sw = New StreamWriter(LogFile, True, Encoding.GetEncoding("BIG5"))
            'sw.WriteLine(strRecord)
            'sw.Flush()
            'sw.Close()
            'sw.Dispose()



        End If

    End Sub

    Public Shared Sub WriteFlowRecord(ByVal seformid As String, ByVal sSql As String)
        Dim strLog_Msg As String = ""
        Dim strRecord As String = ""
        Dim strLog_FileName As String = DateTime.Now.ToString("yyyyMMdd") & "Flow"
        Dim strSql As String = ""
        Dim conn As New C_SQLFUN
        Dim connstr As String = conn.G_conn_string

        Dim db As New SqlConnection(connstr)

        strLog_Msg = GetEFormNameById(seformid)
        Dim iResult

        '--每一段訊息，都要記錄完整的日期與時間（年/月/日/時/分/秒）。
        strRecord = System.String.Format("[{0:yyyy/MM/dd hh:mm:ss}]Message : {1}", DateTime.Now, strLog_Msg.Trim())
        '--第一個參數，檔名。
        '--第二個參數，是否採用APPEND的方式？表示新的資料，會附加在檔案末端。
        '--第三個參數，正體中文編碼 System.Text.Encoding.GetEncoding("BIG5")，預設值UTF-8。
        db.Open()
        strSql = "INSERT INTO ERRORLOG (MSG) VALUES ('" & strRecord & "')"
        iResult = New SqlCommand(strSql, db).ExecuteNonQuery()
        db.Close()

        ''紀錄代理情形
        'Dim strAgentMsg = GetFlowAgentRecord(eformsn)
        Dim strAgentMsg As String = "異動T-Sql：" & sSql.Replace("'", "''")
        '--每一段訊息，都要記錄完整的日期與時間（年/月/日/時/分/秒）。
        strRecord = System.String.Format("[{0:yyyy/MM/dd hh:mm:ss}]Message : {1}", DateTime.Now, strAgentMsg.Trim())
        '--第一個參數，檔名。
        '--第二個參數，是否採用APPEND的方式？表示新的資料，會附加在檔案末端。
        '--第三個參數，正體中文編碼 System.Text.Encoding.GetEncoding("BIG5")，預設值UTF-8。
        db.Open()
        strSql = "INSERT INTO ERRORLOG (MSG) VALUES ('" & strRecord & "')"
        iResult = New SqlCommand(strSql, db).ExecuteNonQuery
        db.Close()

        'Dim strLog_Msg As String = ""
        'Dim strRecord As String = ""
        'Dim strLog_FileName As String = DateTime.Now.ToString("yyyyMMdd") & "Flow"

        'Dim my_Dir As String = HttpContext.Current.Server.MapPath("~/Log") & "\" & DateTime.Now.ToString("yyyyMMdd")
        'If Not Directory.Exists(my_Dir) Then
        '    Directory.CreateDirectory(my_Dir)  '--如果這目錄不存在，就建立它。
        'End If


        ''--每一個紀錄檔的副檔名都是 .log
        'Dim LogFile As String = my_Dir & "\" & strLog_FileName + ".log"

        ''strLog_Msg = GetEFormNameById("74M36O86UE")
        'strLog_Msg = GetEFormNameById(seformid)

        ''--每一段訊息，都要記錄完整的日期與時間（年/月/日/時/分/秒）。
        'strRecord = System.String.Format("[{0:yyyy/MM/dd hh:mm:ss}]Message : {1}", DateTime.Now, strLog_Msg.Trim())
        ''--第一個參數，檔名。
        ''--第二個參數，是否採用APPEND的方式？表示新的資料，會附加在檔案末端。
        ''--第三個參數，正體中文編碼 System.Text.Encoding.GetEncoding("BIG5")，預設值UTF-8。
        'Dim sw As StreamWriter
        'sw = New StreamWriter(LogFile, True, Encoding.GetEncoding("BIG5"))
        'sw.WriteLine(strRecord)
        'sw.Flush()
        'sw.Close()
        'sw.Dispose()

        ' ''紀錄代理情形
        ''Dim strAgentMsg = GetFlowAgentRecord(eformsn)
        'Dim strAgentMsg As String = "異動T-Sql：" & sSql
        ''--每一段訊息，都要記錄完整的日期與時間（年/月/日/時/分/秒）。
        'strRecord = System.String.Format("[{0:yyyy/MM/dd hh:mm:ss}]Message : {1}", DateTime.Now, strAgentMsg.Trim())
        ''--第一個參數，檔名。
        ''--第二個參數，是否採用APPEND的方式？表示新的資料，會附加在檔案末端。
        ''--第三個參數，正體中文編碼 System.Text.Encoding.GetEncoding("BIG5")，預設值UTF-8。
        'sw = New StreamWriter(LogFile, True, Encoding.GetEncoding("BIG5"))
        'sw.WriteLine(strRecord)
        'sw.Flush()
        'sw.Close()
        'sw.Dispose()


    End Sub

    ''' <summary>
    ''' 取得登入者的上一級主管
    ''' </summary>
    ''' <param name="user_id"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetSupervisor(ByVal user_id As String) As ArrayList
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim oReturn As New ArrayList

        Dim strSql As String = "SELECT employee_id,emp_chinese_name,ORG_UID,empemail FROM EMPLOYEE WHERE ORG_UID IN (SELECT PARENT_ORG_UID FROM ADMINGROUP WHERE ORG_UID = '" & user_id & "')"
        DR = DC.CreateReader(strSql)
        If DR.Read() Then
            Dim emp As New BasicEmployee
            emp.employee_id = CType(DR.Item("employee_id"), String)
            emp.emp_chinese_name = CType(DR.Item("emp_chinese_name"), String)
            emp.ORG_UID = CType(DR.Item("ORG_UID"), String)
            emp.empemail = CType(DR.Item("empemail"), String)
            oReturn.Add(emp)
        End If
        DC.Dispose()
        GetSupervisor = oReturn
    End Function
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

'''******************** 以下暫將資料表物件放置通用模組下 ****************************
''' <summary>
''' 會客洽公申請資料表(外部資料庫)
''' </summary>
''' <remarks></remarks>
Public Class P_05A
#Region "Private Declare"
    Private _P_NUM As Object = ""
    Private _EFORMSN As Object = ""
    Private _PWUNIT As Object = ""
    Private _PWTITLE As Object = ""
    Private _PWNAME As Object = ""
    Private _PWIDNO As Object = ""
    Private _PAUNIT As Object = ""
    Private _PANAME As Object = ""
    Private _PATITLE As Object = ""
    Private _PAIDNO As Object = ""
    Private _nAPPTIME As Object = ""
    Private _nRECROOM As Object = ""
    Private _nRECEXIT As Object = ""
    Private _nPLACE As Object = ""
    Private _nPHONE As Object = ""
    Private _nRECDATE As Object = ""
    Private _nSTARTTIME As Object = ""
    Private _nENDTIME As Object = ""
    Private _nSMin As Object = ""
    Private _nEMin As Object = ""
    Private _nREASON As Object = ""
    ''db用參數
    'Private __P_NUM, __EFORMSN, __PWUNIT, __PWTITLE, __PWNAME, __PWIDNO, __PAUNIT, __PANAME, __PATITLE, __PAIDNO, __nAPPTIME, __nRECROOM, __nRECEXIT, __nPLACE, __nPHONE, __nRECDATE, __nSTARTTIME, __nENDTIME, __nSMin, __nEMin, __nREASON As New SqlParameter

#End Region

#Region "Public field value method"
    ''' <summary>
    ''' 資料流水號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property P_NUM As String
        Get
            Return CType(_P_NUM, String)
        End Get

        Set(value As String)
            _P_NUM = value
        End Set
    End Property
    ''' <summary>
    ''' 表單資料ID
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EFORMSN As String
        Get
            Return CType(_EFORMSN, String)
        End Get
        Set(value As String)
            _EFORMSN = value
        End Set
    End Property
    ''' <summary>
    ''' 填表人單位中文名稱
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PWUNIT As String
        Get
            Return CType(_PWUNIT, String)
        End Get
        Set(value As String)
            _PWUNIT = value
        End Set
    End Property
    ''' <summary>
    ''' 填表人級職中文名稱
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PWTITLE As String
        Get
            Return CType(_PWTITLE, String)
        End Get
        Set(value As String)
            _PWTITLE = value
        End Set
    End Property
    ''' <summary>
    ''' 填表人姓名中文名稱
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PWNAME As String
        Get
            Return CType(_PWNAME, String)
        End Get
        Set(value As String)
            _PWNAME = value
        End Set
    End Property
    ''' <summary>
    ''' 填表人人員編號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PWIDNO As String
        Get
            Return CType(_PWIDNO, String)
        End Get
        Set(value As String)
            _PWIDNO = value
        End Set
    End Property
    ''' <summary>
    ''' 申請人單位中文名稱
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PAUNIT As String
        Get
            Return CType(_PAUNIT, String)
        End Get
        Set(value As String)
            _PAUNIT = value
        End Set
    End Property
    ''' <summary>
    ''' 申請人級職中文名稱
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PANAME As String
        Get
            Return CType(_PANAME, String)
        End Get
        Set(value As String)
            _PANAME = value
        End Set
    End Property
    ''' <summary>
    ''' 申請人姓名中文名稱
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PATITLE As String
        Get
            Return CType(_PATITLE, String)
        End Get
        Set(value As String)
            _PATITLE = value
        End Set
    End Property
    ''' <summary>
    ''' 申請人人員編號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PAIDNO As String
        Get
            Return CType(_PAIDNO, String)
        End Get
        Set(value As String)
            _PAIDNO = value
        End Set
    End Property
    ''' <summary>
    ''' 填表時間
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property nAPPTIME As String
        Get
            Return CType(_nAPPTIME, String)
        End Get
        Set(value As String)
            _nAPPTIME = value
        End Set
    End Property
    ''' <summary>
    ''' 會客室
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property nRECROOM As String
        Get
            Return CType(_nRECROOM, String)
        End Get
        Set(value As String)
            _nRECROOM = value
        End Set
    End Property
    ''' <summary>
    ''' 會客入口
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property nRECEXIT As String
        Get
            Return CType(_nRECEXIT, String)
        End Get
        Set(value As String)
            _nRECEXIT = value
        End Set
    End Property
    ''' <summary>
    ''' 地點
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property nPLACE As String
        Get
            Return CType(_nPLACE, String)
        End Get
        Set(value As String)
            _nPLACE = value
        End Set
    End Property
    ''' <summary>
    ''' 連絡電話
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property nPHONE As String
        Get
            Return CType(_nPHONE, String)
        End Get
        Set(value As String)
            _nPHONE = value
        End Set
    End Property
    ''' <summary>
    ''' 會客日期
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property nRECDATE As String
        Get
            Return CType(_nRECDATE, String)
        End Get
        Set(value As String)
            _nRECDATE = value
        End Set
    End Property
    ''' <summary>
    ''' 會客開始時間(時)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property nSTARTTIME As String
        Get
            Return CType(_nSTARTTIME, String)
        End Get
        Set(value As String)
            _nSTARTTIME = value
        End Set
    End Property
    ''' <summary>
    ''' 會客結束時間(時)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property nENDTIME As String
        Get
            Return CType(_nENDTIME, String)
        End Get
        Set(value As String)
            _nENDTIME = value
        End Set
    End Property
    ''' <summary>
    ''' 會客開始時間(分)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property nSMin As String
        Get
            Return CType(_nSMin, String)
        End Get
        Set(value As String)
            _nSMin = value
        End Set
    End Property
    ''' <summary>
    ''' 會客結束時間(分)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property nEMin As String
        Get
            Return CType(_nEMin, String)
        End Get
        Set(value As String)
            _nEMin = value
        End Set
    End Property
    ''' <summary>
    ''' 事由
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property nREASON As String
        Get
            Return CType(_nREASON, String)
        End Get
        Set(value As String)
            _nREASON = value
        End Set
    End Property
#End Region

#Region "Public method"
    Public Sub New()

    End Sub
    ''' <summary>
    ''' 從P_05讀資料至P_05A
    ''' </summary>
    ''' <param name="sEFormSN"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal sEFormSN As String)
        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
            conn.Open()
            Dim strSql As String = "SELECT * FROM P_05 WHERE EFORMSN='" & sEFormSN & "'"
            Dim DR As SqlDataReader
            Dim dt As DateTime
            Dim cmd As New SqlCommand(strSql, conn)
            DR = cmd.ExecuteReader
            If DR.HasRows Then
                If DR.Read Then
                    _P_NUM = IIf(IsDBNull(DR("P_NUM")), Nothing, DR("P_NUM").ToString)
                    _EFORMSN = IIf(IsDBNull(DR("EFORMSN")), Nothing, DR("EFORMSN").ToString)
                    _PWUNIT = IIf(IsDBNull(DR("PWUNIT")), Nothing, DR("PWUNIT").ToString)
                    _PWTITLE = IIf(IsDBNull(DR("PWTITLE")), Nothing, DR("PWTITLE").ToString)
                    _PWNAME = IIf(IsDBNull(DR("PWNAME")), Nothing, DR("PWNAME").ToString)
                    _PWIDNO = IIf(IsDBNull(DR("PWIDNO")), Nothing, DR("PWIDNO").ToString)
                    _PAUNIT = IIf(IsDBNull(DR("PAUNIT")), Nothing, DR("PAUNIT").ToString)
                    _PANAME = IIf(IsDBNull(DR("PANAME")), Nothing, DR("PANAME").ToString)
                    _PATITLE = IIf(IsDBNull(DR("PATITLE")), Nothing, DR("PATITLE").ToString)
                    _PAIDNO = IIf(IsDBNull(DR("PAIDNO")), Nothing, DR("PAIDNO").ToString)

                    DateTime.TryParse(IIf(IsDBNull(DR("nAPPTIME")), New DateTime(1900, 1, 1).ToString(), DR("nAPPTIME")).ToString(), dt)
                    '_nAPPTIME = dt.ToString("yyyy/MM/dd HH:mm:ss")
                    _nAPPTIME = IIf(IsDBNull(DR("nAPPTIME")), Nothing, dt.ToString("yyyy/MM/dd HH:mm:ss"))

                    _nRECROOM = IIf(IsDBNull(DR("nRECROOM")), Nothing, DR("nRECROOM").ToString)
                    _nRECEXIT = IIf(IsDBNull(DR("nRECEXIT")), Nothing, DR("nRECEXIT").ToString)
                    _nPLACE = IIf(IsDBNull(DR("nPLACE")), Nothing, DR("nPLACE").ToString)
                    _nPHONE = IIf(IsDBNull(DR("nPHONE")), Nothing, DR("nPHONE").ToString)

                    DateTime.TryParse(IIf(IsDBNull(DR("nRECDATE")), New DateTime(1900, 1, 1), DR("nRECDATE")).ToString(), dt)
                    '_nRECDATE = dt.ToString("yyyy/MM/dd HH:mm:ss")
                    _nRECDATE = IIf(IsDBNull(DR("nRECDATE")), Nothing, dt.ToString("yyyy/MM/dd HH:mm:ss"))

                    _nSTARTTIME = IIf(IsDBNull(DR("nSTARTTIME")), Nothing, DR("nSTARTTIME").ToString)
                    _nENDTIME = IIf(IsDBNull(DR("nENDTIME")), Nothing, DR("nENDTIME").ToString)
                    _nSMin = IIf(IsDBNull(DR("nSMin")), Nothing, DR("nSMin").ToString)
                    _nEMin = IIf(IsDBNull(DR("nEMin")), Nothing, DR("nEMin").ToString)
                    _nREASON = IIf(IsDBNull(DR("nREASON")), Nothing, DR("nREASON").ToString)
                End If
            End If
        End Using
    End Sub

    Public Sub Insert()
        Try
            Dim iSuccess As Integer = 0
            Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
                conn.Open()
                Dim strSql As String = ""
                strSql = "INSERT INTO P_05A(EFORMSN,PWUNIT,PWTITLE,PWNAME,PWIDNO,PAUNIT,PANAME,PATITLE,PAIDNO,nAPPTIME,nRECROOM,nRECEXIT,nPLACE,nPHONE,nRECDATE,nSTARTTIME,nENDTIME,nSMin,nEMin,nREASON) VALUES ("
                strSql += IIf(_EFORMSN Is Nothing, "null", "'" & _EFORMSN & "'").ToString()
                strSql += IIf(_PWUNIT Is Nothing, ",null", ",'" & _PWUNIT & "'").ToString()
                strSql += IIf(_PWTITLE Is Nothing, ",null", ",'" & _PWTITLE & "'").ToString()
                strSql += IIf(_PWNAME Is Nothing, ",null", ",'" & _PWNAME & "'").ToString()
                strSql += IIf(_PWIDNO Is Nothing, ",null", ",'" & _PWIDNO & "'").ToString()
                strSql += IIf(_PAUNIT Is Nothing, ",null", ",'" & _PAUNIT & "'").ToString()
                strSql += IIf(_PANAME Is Nothing, ",null", ",'" & _PANAME & "'").ToString()
                strSql += IIf(_PATITLE Is Nothing, ",null", ",'" & _PATITLE & "'").ToString()
                strSql += IIf(_PAIDNO Is Nothing, ",null", ",'" & _PAIDNO & "'").ToString()
                strSql += IIf(_nAPPTIME Is Nothing, ",null", ",'" & _nAPPTIME & "'").ToString()
                strSql += IIf(_nRECROOM Is Nothing, ",null", ",'" & _nRECROOM & "'").ToString()
                strSql += IIf(_nRECEXIT Is Nothing, ",null", ",'" & _nRECEXIT & "'").ToString()
                strSql += IIf(_nPLACE Is Nothing, ",null", ",'" & _nPLACE & "'").ToString()
                strSql += IIf(_nPHONE Is Nothing, ",null", ",'" & _nPHONE & "'").ToString()
                strSql += IIf(_nRECDATE Is Nothing, ",null", ",'" & _nRECDATE & "'").ToString()
                strSql += IIf(_nSTARTTIME Is Nothing, ",null", ",'" & _nSTARTTIME & "'").ToString()
                strSql += IIf(_nENDTIME Is Nothing, ",null", ",'" & _nENDTIME & "'").ToString()
                strSql += IIf(_nSMin Is Nothing, ",null", ",'" & _nSMin & "'").ToString()
                strSql += IIf(_nEMin Is Nothing, ",null", ",'" & _nEMin & "'").ToString()
                strSql += IIf(_nREASON Is Nothing, ",null", ",'" & _nREASON & "'").ToString()
                strSql += ")"
                Dim cmd As New SqlCommand(strSql, conn)
                'strSql = "INSERT INTO P_05A(EFORMSN,PWUNIT,PWTITLE,PWNAME,PWIDNO,PAUNIT,PANAME,PATITLE,PAIDNO,nAPPTIME,nRECROOM,nRECEXIT,nPLACE,nPHONE,nRECDATE,nSTARTTIME,nENDTIME,nSMin,nEMin,nREASON) VALUES (@EFORMSN,@PWUNIT,@PWTITLE,@PWNAME,@PWIDNO,@PAUNIT,@PANAME,@PATITLE,@PAIDNO,@nAPPTIME,@nRECROOM,@nRECEXIT,@nPLACE,@nPHONE,@nRECDATE,@nSTARTTIME,@nENDTIME,@nSMin,@nEMin,@nREASON)"
                'Dim cmd As New SqlCommand(strSql, conn)
                'cmd.Parameters.AddWithValue("@EFORMSN", __EFORMSN)
                'cmd.Parameters.AddWithValue("@PWUNIT", __PWUNIT)
                'cmd.Parameters.AddWithValue("@PWTITLE", __PWTITLE)
                'cmd.Parameters.AddWithValue("@PWNAME", __PWNAME)
                'cmd.Parameters.AddWithValue("@PWIDNO", __PWIDNO)
                'cmd.Parameters.AddWithValue("@PAUNIT", _PAUNIT)
                'cmd.Parameters.AddWithValue("@PANAME", __PANAME)
                'cmd.Parameters.AddWithValue("@PATITLE", __PATITLE)
                'cmd.Parameters.AddWithValue("@PAIDNO", __PAIDNO)
                'cmd.Parameters.AddWithValue("@nAPPTIME", __nAPPTIME)
                'cmd.Parameters.AddWithValue("@nRECROOM", __nRECROOM)
                'cmd.Parameters.AddWithValue("@nRECEXIT", __nRECEXIT)
                'cmd.Parameters.AddWithValue("@nPLACE", __nPLACE)
                'cmd.Parameters.AddWithValue("@nPHONE", __nPHONE)
                'cmd.Parameters.AddWithValue("@nRECDATE", __nRECDATE)
                'cmd.Parameters.AddWithValue("@nSTARTTIME", __nSTARTTIME)
                'cmd.Parameters.AddWithValue("@nENDTIME", __nENDTIME)
                'cmd.Parameters.AddWithValue("@nSMin", __nSMin)
                'cmd.Parameters.AddWithValue("@nEMin", __nEMin)
                'cmd.Parameters.AddWithValue("@nREASON", __nREASON)
                iSuccess = cmd.ExecuteNonQuery
                ' ReSharper disable once RedundantAssignment
                cmd = Nothing
                If iSuccess <> 0 Then
                    MessageBox.Show("新增外部資料庫失敗，請重新操作")
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub Insert(ByVal tran As SqlTransaction, ByVal conn As SqlConnection)
        Try
            Dim strSql As String = ""
            If tran IsNot Nothing Then
                Dim iSuccess As Integer = 0
                ''確認對方有無資料
                strSql = "SELECT EFORMSN FROM P_05A WHERE EFORMSN='" & _EFORMSN.ToString() & "'"
                Dim cmdB As New SqlCommand(strSql, conn, tran)
                Dim DRB As SqlDataReader
                DRB = cmdB.ExecuteReader
                If (DRB.HasRows) Then
                    ''有資料則更新P_05的nCheckDT欄位
                    Using connA As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString2").ConnectionString)
                        connA.Open()
                        strSql = "UPDATE P_05 SET nCheckDT = GetDate() WHERE EFORMSN='" & _EFORMSN.ToString() & "'"
                        Dim cmdA As New SqlCommand(strSql, connA)
                        iSuccess = cmdA.ExecuteNonQuery
                        ' ReSharper disable once RedundantAssignment
                        cmdA = Nothing
                        If Not iSuccess > 0 Then
                            MessageBox.Show("更新資料庫(P_05)失敗，請重新操作")
                        End If
                    End Using
                Else
                    Using connA As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString2").ConnectionString)
                        connA.Open()
                        Dim tranA As SqlTransaction = connA.BeginTransaction

                        ''未有資料則新增
                        strSql = "INSERT INTO P_05A(EFORMSN,PWUNIT,PWTITLE,PWNAME,PWIDNO,PAUNIT,PANAME,PATITLE,PAIDNO,nAPPTIME,nRECROOM,nRECEXIT,nPLACE,nPHONE,nRECDATE,nSTARTTIME,nENDTIME,nSMin,nEMin,nREASON) VALUES ("
                        strSql += IIf(_EFORMSN Is Nothing, "null", "'" & _EFORMSN & "'").ToString()
                        strSql += IIf(_PWUNIT Is Nothing, ",null", ",'" & _PWUNIT & "'").ToString()
                        strSql += IIf(_PWTITLE Is Nothing, ",null", ",'" & _PWTITLE & "'").ToString()
                        strSql += IIf(_PWNAME Is Nothing, ",null", ",'" & _PWNAME & "'").ToString()
                        strSql += IIf(_PWIDNO Is Nothing, ",null", ",'" & _PWIDNO & "'").ToString()
                        strSql += IIf(_PAUNIT Is Nothing, ",null", ",'" & _PAUNIT & "'").ToString()
                        strSql += IIf(_PANAME Is Nothing, ",null", ",'" & _PANAME & "'").ToString()
                        strSql += IIf(_PATITLE Is Nothing, ",null", ",'" & _PATITLE & "'").ToString()
                        strSql += IIf(_PAIDNO Is Nothing, ",null", ",'" & _PAIDNO & "'").ToString()
                        strSql += IIf(_nAPPTIME Is Nothing, ",null", ",'" & _nAPPTIME & "'").ToString()
                        strSql += IIf(_nRECROOM Is Nothing, ",null", ",'" & _nRECROOM & "'").ToString()
                        strSql += IIf(_nRECEXIT Is Nothing, ",null", ",'" & _nRECEXIT & "'").ToString()
                        strSql += IIf(_nPLACE Is Nothing, ",null", ",'" & _nPLACE & "'").ToString()
                        strSql += IIf(_nPHONE Is Nothing, ",null", ",'" & _nPHONE & "'").ToString()
                        strSql += IIf(_nRECDATE Is Nothing, ",null", ",'" & _nRECDATE & "'").ToString()
                        strSql += IIf(_nSTARTTIME Is Nothing, ",null", ",'" & _nSTARTTIME & "'").ToString()
                        strSql += IIf(_nENDTIME Is Nothing, ",null", ",'" & _nENDTIME & "'").ToString()
                        strSql += IIf(_nSMin Is Nothing, ",null", ",'" & _nSMin & "'").ToString()
                        strSql += IIf(_nEMin Is Nothing, ",null", ",'" & _nEMin & "'").ToString()
                        strSql += IIf(_nREASON Is Nothing, ",null", ",'" & _nREASON & "'").ToString()
                        strSql += ")"
                        Dim cmdA As New SqlCommand(strSql, connA, tranA)
                        'strSql = "INSERT INTO P_05A(EFORMSN,PWUNIT,PWTITLE,PWNAME,PWIDNO,PAUNIT,PANAME,PATITLE,PAIDNO,nAPPTIME,nRECROOM,nRECEXIT,nPLACE,nPHONE,nRECDATE,nSTARTTIME,nENDTIME,nSMin,nEMin,nREASON) VALUES (@EFORMSN,@PWUNIT,@PWTITLE,@PWNAME,@PWIDNO,@PAUNIT,@PANAME,@PATITLE,@PAIDNO,@nAPPTIME,@nRECROOM,@nRECEXIT,@nPLACE,@nPHONE,@nRECDATE,@nSTARTTIME,@nENDTIME,@nSMin,@nEMin,@nREASON)"
                        'Dim cmd As New SqlCommand(strSql, conn)
                        'cmd.Parameters.AddWithValue("@EFORMSN", __EFORMSN)
                        'cmd.Parameters.AddWithValue("@PWUNIT", __PWUNIT)
                        'cmd.Parameters.AddWithValue("@PWTITLE", __PWTITLE)
                        'cmd.Parameters.AddWithValue("@PWNAME", __PWNAME)
                        'cmd.Parameters.AddWithValue("@PWIDNO", __PWIDNO)
                        'cmd.Parameters.AddWithValue("@PAUNIT", _PAUNIT)
                        'cmd.Parameters.AddWithValue("@PANAME", __PANAME)
                        'cmd.Parameters.AddWithValue("@PATITLE", __PATITLE)
                        'cmd.Parameters.AddWithValue("@PAIDNO", __PAIDNO)
                        'cmd.Parameters.AddWithValue("@nAPPTIME", __nAPPTIME)
                        'cmd.Parameters.AddWithValue("@nRECROOM", __nRECROOM)
                        'cmd.Parameters.AddWithValue("@nRECEXIT", __nRECEXIT)
                        'cmd.Parameters.AddWithValue("@nPLACE", __nPLACE)
                        'cmd.Parameters.AddWithValue("@nPHONE", __nPHONE)
                        'cmd.Parameters.AddWithValue("@nRECDATE", __nRECDATE)
                        'cmd.Parameters.AddWithValue("@nSTARTTIME", __nSTARTTIME)
                        'cmd.Parameters.AddWithValue("@nENDTIME", __nENDTIME)
                        'cmd.Parameters.AddWithValue("@nSMin", __nSMin)
                        'cmd.Parameters.AddWithValue("@nEMin", __nEMin)
                        'cmd.Parameters.AddWithValue("@nREASON", __nREASON)
                        iSuccess = cmdA.ExecuteNonQuery
                        'cmd.ExecuteNonQuery()
                        cmdA = Nothing
                        If Not iSuccess > 0 Then
                            MessageBox.Show("新增外部資料庫失敗，請重新操作")
                        End If
                        ''會客人員明細
                        Dim P_0501AData As New P_0501A(_EFORMSN.ToString())
                        P_0501AData.Insert(tranA, connA)
                        tranA.Commit()
                        tranA.Dispose()
                    End Using
                End If
                DRB.Close()
                DRB = Nothing
            Else
                MessageBox.Show("未有Transaction物件，請重新操作")
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try

    End Sub
#End Region
End Class

''' <summary>
''' 會客洽公人員明細資料表(外部資料庫)
''' </summary>
''' <remarks></remarks>
Public Class P_0501A
#Region "Private Declare"
    Private _Receive_Num As String = ""
    Private _EFORMSN As String = ""
    Private _nNAME As String = ""
    Private _nSEX As String = ""
    Private _nSERVICE As String = ""
    Private _nID As String = ""
    Private _nKIND As String = ""
    Private _nCreateDate As String = ""
    Private _nCarNo As String = ""
    Private DRS As ArrayList
    ''db用參數
    'Private __Receive_Num , __EFORMSN , __nNAME , __nSEX , __nSERVICE , __nID , __nKIND , __nCreateDate As SqlParameter
#End Region

#Region "Public field value method"
    ''' <summary>
    ''' 資料流水號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Receive_Num As String
        Get
            Return _Receive_Num
        End Get
        Set(value As String)
            _Receive_Num = value
        End Set
    End Property
    ''' <summary>
    ''' 表單資料ID
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EFORMSN As String
        Get
            Return _EFORMSN
        End Get
        Set(value As String)
            _EFORMSN = value
        End Set
    End Property
    ''' <summary>
    ''' 姓名   
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property nNAME As String
        Get
            Return _nNAME
        End Get
        Set(value As String)
            _nNAME = value
        End Set
    End Property
    ''' <summary>
    ''' 性別
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property nSEX As String
        Get
            Return _nSEX
        End Get
        Set(value As String)
            _nSEX = value
        End Set
    End Property
    ''' <summary>
    ''' 服務單位
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property nSERVICE As String
        Get
            Return _nSERVICE
        End Get
        Set(value As String)
            _nSERVICE = value
        End Set
    End Property
    ''' <summary>
    ''' 身分證字號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property nID As String
        Get
            Return _nID
        End Get
        Set(value As String)
            _nID = value
        End Set
    End Property
    ''' <summary>
    ''' 類別(民, 軍)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property nKIND As String
        Get
            Return _nKIND
        End Get
        Set(value As String)
            _nKIND = value
        End Set
    End Property
    ''' <summary>
    ''' 資料產生日期
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property nCreateDate As String
        Get
            Return _nCreateDate
        End Get
        Set(value As String)
            _nCreateDate = value
        End Set
    End Property

    Public Property nCarNo() as string
        Get
            Return _nCarNo
        End Get
        Set(value As String)
            _nCarNo = value
        End Set
    End Property

    Public Function MultiData() As ArrayList
        Return DRS
    End Function
#End Region

#Region "Public method"
    Public Sub New()

    End Sub
    Public Sub New(ByVal sEFormSN As String, ByVal sNID As String)
        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
            conn.Open()
            Dim strSql As String = ""
            strSql = "SELECT * FROM P_0501 WHERE EFORMSN='" & sEFormSN & "' AND NID='" & sNID & "'"
            Dim DR As SqlDataReader
            Dim dt As DateTime
            Dim cmd As New SqlCommand(strSql, conn)
            DR = cmd.ExecuteReader
            If DR.HasRows Then
                If DR.Read Then
                    _Receive_Num = CType(IIf(IsDBNull(DR("Receive_Num")), Nothing, DR("Receive_Num").ToString), String)
                    _EFORMSN = CType(IIf(IsDBNull(DR("EFORMSN")), Nothing, DR("EFORMSN").ToString), String)
                    _nNAME = CType(IIf(IsDBNull(DR("nNAME")), Nothing, DR("nNAME").ToString), String)
                    _nSEX = CType(IIf(IsDBNull(DR("nSEX")), Nothing, DR("nSEX").ToString), String)
                    _nSERVICE = CType(IIf(IsDBNull(DR("nSERVICE")), Nothing, DR("nSERVICE").ToString), String)
                    _nID = CType(IIf(IsDBNull(DR("nID")), Nothing, DR("nID").ToString), String)
                    _nKIND = CType(IIf(IsDBNull(DR("nKIND")), Nothing, DR("nKIND").ToString), String)
                    _nCarNo = CType(IIf(IsDBNull(DR("nCarNo")), Nothing, DR("nCarNo").ToString), String)
                    DateTime.TryParse(CType(IIf(IsDBNull(DR("nCreateDate")), New DateTime(1900, 1, 1), DR("nCreateDate")), String), dt)
                    '_nCreateDate = dt.ToString("yyyy/MM/dd HH:m:s")
                    _nCreateDate = CType(IIf(IsDBNull(DR("nCreateDate")), Nothing, dt.ToString("yyyy/MM/dd HH:m:s")), String)
                End If
            End If
        End Using
    End Sub
    Public Sub New(ByVal sEFormSN As String)
        DRS = New ArrayList
        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
            conn.Open()
            Dim strSql As String = ""
            strSql = "SELECT * FROM P_0501 WHERE EFORMSN='" & sEFormSN & "'"
            Dim DR As SqlDataReader
            Dim dt As DateTime
            Dim cmd As New SqlCommand(strSql, conn)
            DR = cmd.ExecuteReader
            If DR.HasRows Then
                While DR.Read
                    Dim x As New P_0501A()

                    x.Receive_Num = CType(IIf(IsDBNull(DR("Receive_Num")), Nothing, DR("Receive_Num").ToString), String)
                    x.EFORMSN = CType(IIf(IsDBNull(DR("EFORMSN")), Nothing, DR("EFORMSN").ToString), String)
                    x.nNAME = CType(IIf(IsDBNull(DR("nNAME")), Nothing, DR("nNAME").ToString), String)
                    x.nSEX = CType(IIf(IsDBNull(DR("nSEX")), Nothing, DR("nSEX").ToString), String)
                    x.nSERVICE = CType(IIf(IsDBNull(DR("nSERVICE")), Nothing, DR("nSERVICE").ToString), String)
                    x.nID = CType(IIf(IsDBNull(DR("nID")), Nothing, DR("nID").ToString), String)
                    x.nKIND = CType(IIf(IsDBNull(DR("nKIND")), Nothing, DR("nKIND").ToString), String)

                    DateTime.TryParse(CType(IIf(IsDBNull(DR("nCreateDate")), New DateTime(1900, 1, 1), DR("nCreateDate")), String), dt)
                    'x.nCreateDate = dt.ToString("yyyy/MM/dd HH:mm:ss")
                    x.nCreateDate = CType(IIf(IsDBNull(DR("nCreateDate")), Nothing, dt.ToString("yyyy/MM/dd HH:mm:ss")), String)
                    'x.nCreateDate_DB = IIf(IsDBNull(DR("nCreateDate")), DBNull.Value, dt.ToString("yyyy/MM/dd HH:mm:ss"))
                    x.nCarNo = CType(IIf(IsDBNull(DR("nCarNo")), Nothing, DR("nCarNo").ToString), String)

                    DRS.Add(x)
                End While
            End If
        End Using
    End Sub
    Public Sub Insert()

    End Sub
    Public Sub Insert(ByVal tran As SqlTransaction, ByVal conn As SqlConnection)
        Try
            If tran IsNot Nothing Then
                'Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString2").ConnectionString)
                '    conn.Open()                
                Dim strSql As String = ""
                Dim iSuccess As Integer = 0
                For Each data As P_0501A In DRS
                    ''strSql += "INSERT INTO P_0501A(Receive_Num,EFORMSN,nNAME,nSEX,nSERVICE,nID,nKIND,nCreateDate) VALUES ("
                    strSql = "INSERT INTO P_0501A(EFORMSN,nNAME,nSEX,nSERVICE,nID,nKIND,nCreateDate,nCarNo) VALUES ("
                    ''strSql += "'" & data.Receive_Num & "'"
                    strSql += CType(IIf(data.EFORMSN Is Nothing, "null", "'" & data.EFORMSN & "'"), String)
                    strSql += CType(IIf(data.nNAME Is Nothing, ",null", ",'" & data.nNAME & "'"), String)
                    strSql += CType(IIf(data.nSEX Is Nothing, ",null", ",'" & data.nSEX & "'"), String)
                    strSql += CType(IIf(data.nSERVICE Is Nothing, ",null", ",'" & data.nSERVICE & "'"), String)
                    strSql += CType(IIf(data.nID Is Nothing, ",null", ",'" & data.nID & "'"), String)
                    strSql += CType(IIf(data.nKIND Is Nothing, ",null", ",'" & data.nKIND & "'"), String)
                    strSql += CType(IIf(data.nCreateDate Is Nothing, ",null", ",'" & data.nCreateDate & "'"), String)
                    strSql += CType(IIf(data.nCarNo Is Nothing, ",null", ",'" & data.nCarNo & "'"), String)
                    strSql += ")"
                    Dim cmd As New SqlCommand(strSql, conn, tran)
                    'strSql = "INSERT INTO P_0501A(Receive_Num,EFORMSN,nNAME,nSEX,nSERVICE,nID,nKIND,nCreateDate) VALUES (@Receive_Num,@EFORMSN,@nNAME,@nSEX,@nSERVICE,@nID,@nKIND,@nCreateDate)"
                    'Dim cmd As New SqlCommand(strSql, conn)
                    'cmd.Parameters.AddWithValue("@Receive_Num", data.Receive_Num_DB)
                    'cmd.Parameters.AddWithValue("@EFORMSN", data.EFORMSN_DB)
                    'cmd.Parameters.AddWithValue("@nNAME", data.nNAME_DB)
                    'cmd.Parameters.AddWithValue("@nSEX", data.nSEX_DB)
                    'cmd.Parameters.AddWithValue("@nSERVICE", data.nSERVICE_DB)
                    'cmd.Parameters.AddWithValue("@nID", data.nID_DB)
                    'cmd.Parameters.AddWithValue("@nKIND", data.nKIND_DB)
                    'cmd.Parameters.AddWithValue("@nCreateDate", data.nCreateDate_DB)

                    iSuccess = cmd.ExecuteNonQuery
                    cmd = Nothing
                    If Not iSuccess > 0 Then
                        MessageBox.Show("新增外部資料庫失敗，請重新操作")
                    End If
                Next
                'End Using
            Else
                MessageBox.Show("未有Transaction物件，請重新操作")
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try

    End Sub
#End Region

End Class

''' <summary>
''' 門禁會議管制申請資料表(外部資料庫)
''' </summary>
''' <remarks></remarks>
Public Class P_09A
#Region "Private Declare"
    Private _ID As String = ""
    Private _EFORMSN As String = ""
    Private _FillOutBy As String = ""
    Private _CreateBy As String = ""
    Private _Subject As String = ""
    Private _MeetingDate As String = ""
    Private _MeetingHour As String = ""
    Private _MeetingMinute As String = ""
    Private _Location As String = ""
    Private _Sponsor As String = ""
    Private _Moderator As String = ""
    Private _PhoneNumber As String = ""
    Private _DocumentNo As String = ""
    Private _EnteringPeopleNumber As String = ""
    Private _Status As String = ""
    Private _CreateDate As String = ""
    Private _ModifyBy As String = ""
    Private _ModifyDate As String = ""
    Private _PENDFLAG As String = ""
    Private _nCheckDT As String = ""

#End Region

#Region "Public field value method"
    Public Property ID As String
        Get
            Return _ID
        End Get
        Set(value As String)
            _ID = value
        End Set
    End Property
    Public Property EFORMSN As String
        Get
            Return _EFORMSN
        End Get
        Set(value As String)
            _EFORMSN = value
        End Set
    End Property
    Public Property FillOutBy As String
        Get
            Return _FillOutBy
        End Get
        Set(value As String)
            _FillOutBy = value
        End Set
    End Property
    Public Property CreateBy As String
        Get
            Return _CreateBy
        End Get
        Set(value As String)
            _CreateBy = value
        End Set
    End Property
    Public Property Subject As String
        Get
            Return _Subject
        End Get
        Set(value As String)
            _Subject = value
        End Set
    End Property
    Public Property MeetingDate As String
        Get
            Return _MeetingDate
        End Get
        Set(value As String)
            _MeetingDate = value
        End Set
    End Property
    Public Property MeetingHour As String
        Get
            Return _MeetingHour
        End Get
        Set(value As String)
            _MeetingHour = value
        End Set
    End Property
    Public Property MeetingMinute As String
        Get
            Return _MeetingMinute
        End Get
        Set(value As String)
            _MeetingMinute = value
        End Set
    End Property
    Public Property Location As String
        Get
            Return _Location
        End Get
        Set(value As String)
            _Location = value
        End Set
    End Property
    Public Property Sponsor As String
        Get
            Return _Sponsor
        End Get
        Set(value As String)
            _Sponsor = value
        End Set
    End Property
    Public Property Moderator As String
        Get
            Return _Moderator
        End Get
        Set(value As String)
            _Moderator = value
        End Set
    End Property
    Public Property PhoneNumber As String
        Get
            Return _PhoneNumber
        End Get
        Set(value As String)
            _PhoneNumber = value
        End Set
    End Property
    Public Property DocumentNo As String
        Get
            Return _DocumentNo
        End Get
        Set(value As String)
            _DocumentNo = value
        End Set
    End Property
    Public Property EnteringPeopleNumber As String
        Get
            Return _EnteringPeopleNumber
        End Get
        Set(value As String)
            _EnteringPeopleNumber = value
        End Set
    End Property
    Public Property Status As String
        Get
            Return _Status
        End Get
        Set(value As String)
            _Status = value
        End Set
    End Property
    Public Property CreateDate As String
        Get
            Return _CreateDate
        End Get
        Set(value As String)
            _CreateDate = value
        End Set
    End Property
    Public Property ModifyBy As String
        Get
            Return _ModifyBy
        End Get
        Set(value As String)
            _ModifyBy = value
        End Set
    End Property
    Public Property ModifyDate As String
        Get
            Return _ModifyDate
        End Get
        Set(value As String)
            _ModifyDate = value
        End Set
    End Property
    Public Property PENDFLAG As String
        Get
            Return _PENDFLAG
        End Get
        Set(value As String)
            _PENDFLAG = value
        End Set
    End Property
    Public Property nCheckDT As String
        Get
            Return _nCheckDT
        End Get
        Set(value As String)
            _nCheckDT = value
        End Set
    End Property
#End Region

#Region "Public method"
    Public Sub New()

    End Sub
    Public Sub New(ByVal sEFormSN As String)
        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
            conn.Open()
            Dim strSql As String = ""
            strSql = "SELECT * FROM P_09 WHERE EFORMSN='" & sEFormSN & "'"
            Dim DR As SqlDataReader
            Dim dt As DateTime
            Dim cmd As New SqlCommand(strSql, conn)
            DR = cmd.ExecuteReader
            If DR.HasRows Then
                If DR.Read Then
                    _ID = CType(IIf(IsDBNull(DR("ID")), Nothing, DR("ID").ToString), String)
                    _EFORMSN = CType(IIf(IsDBNull(DR("EFORMSN")), Nothing, DR("EFORMSN").ToString), String)
                    _FillOutBy = CType(IIf(IsDBNull(DR("FillOutBy")), Nothing, DR("FillOutBy").ToString), String)
                    _CreateBy = CType(IIf(IsDBNull(DR("CreateBy")), Nothing, DR("CreateBy").ToString), String)
                    _Subject = CType(IIf(IsDBNull(DR("Subject")), Nothing, DR("Subject").ToString), String)

                    DateTime.TryParse(CType(IIf(IsDBNull(DR("MeetingDate")), New DateTime(1900, 1, 1), DR("MeetingDate")), String), dt)
                    '_MeetingDate = dt.ToString("yyyy/MM/dd HH:mm:ss")
                    _MeetingDate = CType(IIf(IsDBNull(DR("MeetingDate")), Nothing, dt.ToString("yyyy/MM/dd HH:mm:ss")), String)

                    _MeetingHour = CType(IIf(IsDBNull(DR("MeetingHour")), Nothing, DR("MeetingHour").ToString), String)
                    _MeetingMinute = CType(IIf(IsDBNull(DR("MeetingMinute")), Nothing, DR("MeetingMinute").ToString), String)
                    _Location = CType(IIf(IsDBNull(DR("Location")), Nothing, DR("Location").ToString), String)
                    _Sponsor = CType(IIf(IsDBNull(DR("Sponsor")), Nothing, DR("Sponsor").ToString), String)
                    _Moderator = CType(IIf(IsDBNull(DR("Moderator")), Nothing, DR("Moderator").ToString), String)
                    _PhoneNumber = CType(IIf(IsDBNull(DR("PhoneNumber")), Nothing, DR("PhoneNumber").ToString), String)
                    _DocumentNo = CType(IIf(IsDBNull(DR("DocumentNo")), Nothing, DR("DocumentNo").ToString), String)
                    _EnteringPeopleNumber = CType(IIf(IsDBNull(DR("EnteringPeopleNumber")), Nothing, DR("EnteringPeopleNumber").ToString), String)
                    _Status = CType(IIf(IsDBNull(DR("Status")), Nothing, DR("Status").ToString), String)

                    DateTime.TryParse(CType(IIf(IsDBNull(DR("CreateDate")), New DateTime(1900, 1, 1), DR("CreateDate")), String), dt)
                    '_CreateDate = dt.ToString("yyyy/MM/dd HH:mm:ss")
                    _CreateDate = CType(IIf(IsDBNull(DR("CreateDate")), Nothing, dt.ToString("yyyy/MM/dd HH:mm:ss")), String)

                    _ModifyBy = CType(IIf(IsDBNull(DR("ModifyBy")), Nothing, DR("ModifyBy").ToString), String)

                    DateTime.TryParse(CType(IIf(IsDBNull(DR("ModifyDate")), New DateTime(1900, 1, 1), DR("ModifyDate")), String), dt)
                    '_ModifyDate = dt.ToString("yyyy/MM/dd HH:mm:ss")
                    _ModifyDate = CType(IIf(IsDBNull(DR("ModifyDate")), Nothing, dt.ToString("yyyy/MM/dd HH:mm:ss")), String)

                    _PENDFLAG = CType(IIf(IsDBNull(DR("PENDFLAG")), Nothing, DR("PENDFLAG").ToString), String)

                    DateTime.TryParse(CType(IIf(IsDBNull(DR("nCheckDT")), New DateTime(1900, 1, 1), DR("nCheckDT")), String), dt)
                    '_nCheckDT = dt.ToString("yyyy/MM/dd HH:mm:ss")
                    _nCheckDT = CType(IIf(IsDBNull(DR("nCheckDT")), Nothing, dt.ToString("yyyy/MM/dd HH:mm:ss")), String)

                    'DateTime.TryParse(DR("nCreateDate"), dt)
                    'x.nCreateDate = dt.ToString("yyyy/MM/dd H:m:s")

                End If
            End If
        End Using
    End Sub

    Public Sub Insert()

    End Sub

    Public Sub Insert(ByVal tran As SqlTransaction, ByVal conn As SqlConnection)
        Try
            Dim strSql As String = ""
            If tran IsNot Nothing Then
                Dim iSuccess As Integer = 0
                ''確認對方有無資料
                strSql = "SELECT EFORMSN FROM P_09A WHERE EFORMSN='" & _EFORMSN & "'"
                Dim cmdB As New SqlCommand(strSql, conn, tran)
                Dim DRB As SqlDataReader
                DRB = cmdB.ExecuteReader
                If (DRB.HasRows) Then
                    ''有資料則更新P_09的nCheckDT欄位
                    Using connA As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
                        connA.Open()
                        strSql = "UPDATE P_09 SET nCheckDT = GetDate() WHERE EFORMSN='" & _EFORMSN & "'"
                        Dim cmdA As New SqlCommand(strSql, connA)
                        iSuccess = cmdA.ExecuteNonQuery
                        cmdA = Nothing
                        If Not iSuccess > 0 Then
                            MessageBox.Show("更新資料庫(P_09)失敗，請重新操作")
                        End If
                    End Using
                Else
                    '未有則新增
                    Using connA As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString2").ConnectionString)
                        connA.Open()
                        Dim tranA As SqlTransaction = connA.BeginTransaction
                        strSql = "INSERT INTO P_09A(ID,EFORMSN,FillOutBy,CreateBy,Subject,MeetingDate,MeetingHour,MeetingMinute,Location,Sponsor,Moderator,PhoneNumber,DocumentNo,EnteringPeopleNumber,CreateDate,ModifyBy,ModifyDate,PENDFLAG,Status) VALUES ("

                        strSql += CType(IIf(_ID Is Nothing, "null", "'" & _ID & "'"), String)
                        strSql += CType(IIf(_EFORMSN Is Nothing, ",null", ",'" & _EFORMSN & "'"), String)
                        strSql += CType(IIf(_FillOutBy Is Nothing, ",null", ",'" & _FillOutBy & "'"), String)
                        strSql += CType(IIf(_CreateBy Is Nothing, ",null", ",'" & _CreateBy & "'"), String)
                        strSql += CType(IIf(_Subject Is Nothing, ",null", ",'" & _Subject & "'"), String)
                        strSql += CType(IIf(_MeetingDate Is Nothing, ",null", ",'" & _MeetingDate & "'"), String)
                        strSql += CType(IIf(_MeetingHour Is Nothing, ",null", ",'" & _MeetingHour & "'"), String)
                        strSql += CType(IIf(_MeetingMinute Is Nothing, ",null", ",'" & _MeetingMinute & "'"), String)
                        strSql += CType(IIf(_Location Is Nothing, ",null", ",'" & _Location & "'"), String)
                        strSql += CType(IIf(_Sponsor Is Nothing, ",null", ",'" & _Sponsor & "'"), String)
                        strSql += CType(IIf(_Moderator Is Nothing, ",null", ",'" & _Moderator & "'"), String)
                        strSql += CType(IIf(_PhoneNumber Is Nothing, ",null", ",'" & _PhoneNumber & "'"), String)
                        strSql += CType(IIf(_DocumentNo Is Nothing, ",null", ",'" & _DocumentNo & "'"), String)
                        strSql += CType(IIf(_EnteringPeopleNumber Is Nothing, ",null", ",'" & _EnteringPeopleNumber & "'"), String)
                        strSql += CType(IIf(_CreateDate Is Nothing, ",null", ",'" & _CreateDate & "'"), String)
                        strSql += CType(IIf(_ModifyBy Is Nothing, ",null", ",'" & _ModifyBy & "'"), String)
                        strSql += CType(IIf(_ModifyDate Is Nothing, ",null", ",'" & _ModifyDate & "'"), String)
                        strSql += CType(IIf(_PENDFLAG Is Nothing, ",null", ",'" & _PENDFLAG & "'"), String)
                        strSql += CType(IIf(_Status Is Nothing, ",null", ",'" & _Status & "'"), String)
                        strSql += ")"
                        Dim cmdA As New SqlCommand(strSql, connA, tranA)
                        iSuccess = cmdA.ExecuteNonQuery
                        cmdA = Nothing
                        If Not iSuccess > 0 Then
                            MessageBox.Show("新增外部資料庫失敗，請重新操作")
                        End If
                        ''進出營門
                        Dim P_0901AData As New P_0901A(_EFORMSN)
                        P_0901AData.Insert(tranA, connA)
                        ''上傳明細
                        Dim P_09UploadData As New UploadA(_EFORMSN)
                        P_09UploadData.Insert(tranA, connA)
                        tranA.Commit()
                        tranA.Dispose()
                    End Using

                End If
                DRB.Close()
                DRB = Nothing
            Else
                MessageBox.Show("未有Transaction物件，請重新操作")
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try

    End Sub
#End Region

End Class

''' <summary>
''' 門禁會議管制進出營門資料表(外部資料庫)
''' </summary>
''' <remarks></remarks>
Public Class P_0901A
#Region "Private Declare"
    Private _ID As String = ""
    Private _EFORMSN As String = ""
    Private _GateNumber As String = ""
    Private DRS As ArrayList
#End Region

#Region "Public field value method"
    Public Property ID As String
        Get
            Return _ID
        End Get
        Set(value As String)
            _ID = value
        End Set
    End Property
    Public Property EFORMSN As String
        Get
            Return _EFORMSN
        End Get
        Set(value As String)
            _EFORMSN = value
        End Set
    End Property
    Public Property GateNumber As String
        Get
            Return _GateNumber
        End Get
        Set(value As String)
            _GateNumber = value
        End Set
    End Property
#End Region

#Region "Public method"
    Public Sub New()

    End Sub

    Public Sub New(ByVal sEFormSN As String)
        Try
            DRS = New ArrayList
            Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
                conn.Open()
                Dim strSql As String = ""
                strSql = "SELECT * FROM P_0901 WHERE EFORMSN='" & sEFormSN & "'"
                Dim DR As SqlDataReader
                Dim cmd As New SqlCommand(strSql, conn)
                DR = cmd.ExecuteReader
                If DR.HasRows Then
                    While DR.Read
                        Dim x As New P_0901A
                        x.ID = CType(IIf(IsDBNull(DR("ID")), Nothing, DR("ID").ToString), String)
                        x.EFORMSN = CType(IIf(IsDBNull(DR("EFORMSN")), Nothing, DR("EFORMSN").ToString), String)
                        x.GateNumber = CType(IIf(IsDBNull(DR("GateNumber")), Nothing, DR("GateNumber").ToString), String)
                        DRS.Add(x)
                    End While
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Public Sub Insert()

    End Sub

    Public Sub Insert(ByVal tran As SqlTransaction, ByVal conn As SqlConnection)
        Try
            If tran IsNot Nothing Then
                'Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString2").ConnectionString)
                '    conn.Open()
                Dim strSql As String = ""
                Dim iSuccess As Integer = 0
                For Each data As P_0901A In DRS
                    strSql = "INSERT INTO P_0901A(ID,EFORMSN,GateNumber) VALUES ("
                    strSql += CType(IIf(data._ID Is Nothing, "null", "'" & data._ID & "'"), String)
                    strSql += CType(IIf(data._EFORMSN Is Nothing, ",null", ",'" & data._EFORMSN & "'"), String)
                    strSql += CType(IIf(data._GateNumber Is Nothing, ",null", ",'" & data._GateNumber & "'"), String)
                    strSql += ")"
                    Dim cmd As New SqlCommand(strSql, conn, tran)
                    iSuccess = cmd.ExecuteNonQuery
                    cmd = Nothing
                    If Not iSuccess > 0 Then
                        MessageBox.Show("新增外部資料庫失敗，請重新操作")
                    End If
                Next
                'End Using
            Else
                MessageBox.Show("未有Transaction物件，請重新操作")
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try
    End Sub
#End Region
End Class

''' <summary>
''' 上傳檔案明細資料表(外部資料庫)
''' </summary>
''' <remarks></remarks>
Public Class UploadA
#Region "Private Declare"
    Private _Upload_id As String = ""
    Private _eformsn As String = ""
    Private _FileName As String = ""
    Private _FilePath As String = ""
    Private _Upload_Time As String = ""
    Private DRS As ArrayList
#End Region

#Region "Public field value method"
    Public Property Upload_id As String
        Get
            Return _Upload_id
        End Get
        Set(value As String)
            _Upload_id = value
        End Set
    End Property
    Public Property EFORMSN As String
        Get
            Return _eformsn
        End Get
        Set(value As String)
            _eformsn = value
        End Set
    End Property
    Public Property FileName As String
        Get
            Return _FileName
        End Get
        Set(value As String)
            _FileName = value
        End Set
    End Property
    Public Property FilePath As String
        Get
            Return _FilePath
        End Get
        Set(value As String)
            _FilePath = value
        End Set
    End Property
    Public Property Upload_Time As String
        Get
            Return _Upload_Time
        End Get
        Set(value As String)
            _Upload_Time = value
        End Set
    End Property
#End Region

#Region "Public method"
    Public Sub New()

    End Sub
    Public Sub New(ByVal sEFormSN As String)
        Try

            DRS = New ArrayList
            Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
                conn.Open()
                Dim strSql As String = ""
                strSql = "SELECT * FROM UPLOAD WHERE EFORMSN='" & sEFormSN & "'"
                Dim DR As SqlDataReader
                Dim dt As DateTime
                Dim cmd As New SqlCommand(strSql, conn)
                DR = cmd.ExecuteReader
                If DR.HasRows Then
                    If DR.Read Then
                        Dim x As New UploadA()

                        x.Upload_id = CType(IIf(IsDBNull(DR("Upload_id")), Nothing, DR("Upload_id").ToString), String)
                        x.EFORMSN = CType(IIf(IsDBNull(DR("EFORMSN")), Nothing, DR("EFORMSN").ToString), String)
                        x.FileName = CType(IIf(IsDBNull(DR("FileName")), Nothing, DR("FileName").ToString), String)
                        x.FilePath = CType(IIf(IsDBNull(DR("FilePath")), Nothing, DR("FilePath").ToString), String)

                        DateTime.TryParse(CType(IIf(IsDBNull(DR("Upload_Time")), New DateTime(1900, 1, 1), DR("Upload_Time")), String), dt)
                        'x.Upload_Time = dt.ToString("yyyy/MM/dd HH:mm:ss")
                        x.Upload_Time = CType(IIf(IsDBNull(DR("Upload_Time")), Nothing, dt.ToString("yyyy/MM/dd HH:mm:ss")), String)
                        DRS.Add(x)
                    End If
                End If
            End Using
        Catch ex As Exception

        End Try

    End Sub
    Public Sub Insert()

    End Sub

    Public Sub Insert(ByVal tran As SqlTransaction, ByVal conn As SqlConnection)
        Try
            Dim tool As New C_Public()
            If tran IsNot Nothing Then
                Dim iSuccess As Integer = 0
                'Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString2").ConnectionString)
                '    conn.Open()
                Dim strSql As String = ""
                For Each data As UploadA In DRS
                    'strSql = "INSERT INTO Upload(eformsn,FileName,FilePath,Upload_Time,Data) VALUES ("
                    'strSql += CType(IIf(data._eformsn Is Nothing, "null", "'" & data._eformsn & "'"), String)
                    'strSql += CType(IIf(data._FileName Is Nothing, ",null", ",'" & data._FileName & "'"), String)
                    'strSql += CType(IIf(data._FilePath Is Nothing, ",null", ",'" & data._FilePath & "'"), String)
                    'strSql += CType(IIf(data._Upload_Time Is Nothing, ",null", ",'" & data._Upload_Time & "'"), String)
                    'strSql += ")"

                    strSql = "INSERT INTO Upload(Upload_id,eformsn,FileName,FilePath,Upload_Time,Data) VALUES (@Upload_id,@eformsn,@FileName,@FilePath,@Upload_Time,@Data)"
                    Dim cmd As New SqlCommand(strSql, conn, tran)
                    cmd.Parameters.Add("@Upload_id", SqlDbType.VarChar).Value = data.Upload_id
                    cmd.Parameters.Add("@eformsn", SqlDbType.VarChar).Value = data.EFORMSN
                    cmd.Parameters.Add("@FileName", SqlDbType.VarChar).Value = data.FileName
                    cmd.Parameters.Add("@FilePath", SqlDbType.VarChar).Value = data.FilePath
                    cmd.Parameters.Add("@Upload_Time", SqlDbType.DateTime).Value = data.Upload_Time
                    cmd.Parameters.Add("@Data", SqlDbType.Binary).Value = tool.ConvertFileToByteArray(data.FilePath, data.FileName)

                    'Dim cmd As New SqlCommand(strSql, conn, tran)
                    iSuccess = cmd.ExecuteNonQuery
                    cmd = Nothing
                    If Not iSuccess > 0 Then
                        MessageBox.Show("新增外部資料庫失敗，請重新操作")
                    End If
                Next
                'End Using
            Else
                MessageBox.Show("未有Transaction物件，請重新操作")
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try

    End Sub
#End Region
End Class
