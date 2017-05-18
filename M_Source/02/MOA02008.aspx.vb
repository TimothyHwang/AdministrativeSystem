Imports System.Data.SqlClient
Partial Class Source_02_MOA02008
    Inherits System.Web.UI.Page

    Dim connstr, MeetName As String
    Dim user_id, org_uid As String
    Dim Meetsn As Integer
    Dim PAIDNO As String
    Dim SDate, EDate As String
    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged

        '分頁
        Dim strOrd As String

        strOrd = " ORDER BY P_0204.MeetTime DESC"

        SqlDataSource1.SelectCommand = SQLALL(Meetsn, PAIDNO, SDate, EDate) & strOrd

    End Sub
    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated

        If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then
            '隱藏MTT_Num
            e.Row.Cells(6).Visible = False
        End If

    End Sub

    Protected Sub BackBtn_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BackBtn.Click

        '回上頁
        Server.Transfer("MOA02003.aspx")

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            user_id = Session("user_id")
            org_uid = Session("ORG_UID")

            'session被清空回首頁
            If user_id = "" Or org_uid = "" Then

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")

            Else

                Meetsn = Request.QueryString("Meetsn")
                PAIDNO = Request.QueryString("PAIDNO")
                SDate = Request.QueryString("SDate")
                EDate = Request.QueryString("EDate")

                If IsPostBack = False Then

                    '會議室代碼
                    MeetName = ""

                    Dim conn As New C_SQLFUN
                    connstr = conn.G_conn_string

                    '開啟連線
                    Dim db As New SqlConnection(connstr)

                    '會議室名稱
                    db.Open()
                    Dim strPer As New SqlCommand("SELECT MeetName FROM P_0201 WHERE MeetSn = '" & Meetsn & "'", db)
                    Dim RdPer = strPer.ExecuteReader()
                    If RdPer.read() Then
                        MeetName = RdPer("MeetName")
                    End If
                    db.Close()

                    LabTitle.Text += "(" + MeetName + ")"

                    Dim strOrd As String

                    strOrd = " ORDER BY P_0204.MeetTime DESC"

                    SqlDataSource1.SelectCommand = SQLALL(Meetsn, PAIDNO, SDate, EDate) & strOrd

                End If

            End If

        Catch ex As Exception

        End Try


    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting

        Dim strComment As TextBox = GridView1.Rows(e.RowIndex).FindControl("txtComment")

        '沒輸入註銷理由不可註銷會議室
        If strComment.Text = "" Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('沒輸入註銷理由，不可註銷會議室');")
            Response.Write(" </script>")
            e.Cancel = True
        Else

            Dim MTT_Num As String = ""
            Dim eformsn As String = ""
            Dim AuthPer As String = ""

            '取得刪除的會議室流水號
            MTT_Num = GridView1.DataKeys(e.RowIndex).Value

            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '將會議室申請時段放到移除資料表
            db.Open()
            Dim insCom As New SqlCommand("INSERT INTO P_0205 (MTT_Num, EFORMSN, MeetSn, MeetTime, MeetHour, DelUser, DelComment) SELECT MTT_Num, EFORMSN, MeetSn, MeetTime, MeetHour, '" & Session("USER_ID") & "', '" & strComment.Text & "' FROM P_0204 WHERE MTT_Num='" & MTT_Num & "'", db)
            insCom.ExecuteNonQuery()
            db.Close()

            '取得表單eformsn
            db.Open()
            Dim strPer As New SqlCommand("SELECT eformsn FROM P_0204 WHERE MTT_Num = '" & MTT_Num & "'", db)
            Dim RdPer = strPer.ExecuteReader()
            If RdPer.read() Then
                eformsn = RdPer("eformsn")
            End If
            db.Close()

            Dim FC As New C_FlowSend.C_FlowSend

            '找出表單是否尚未批核完成
            db.Open()
            Dim strAuth As New SqlCommand("select empuid from flowctl where eformsn = '" & eformsn & "' and hddate is Null ", db)
            Dim RdvAuth = strAuth.ExecuteReader()
            If RdvAuth.read() Then
                AuthPer = RdvAuth.Item("empuid")
            End If
            db.Close()

            If AuthPer <> "" Then

                '駁回表單
                Dim Val_P As String = ""

                Dim SendVal As String = "4rM2YFP73N" & "," & AuthPer & "," & eformsn & "," & "1"

                Val_P = FC.F_Back(SendVal, ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

            End If

            '寄送Mail
            Dim MOAServer As String = ""
            Dim SmtpHost As String = ""
            Dim SystemMail As String = ""
            Dim MailYN As String = ""
            MOAServer = FC.F_MailBase("MOAServer", connstr)
            SmtpHost = FC.F_MailBase("SmtpHost", connstr)
            SystemMail = FC.F_MailBase("SystemMail", connstr)
            MailYN = FC.F_MailBase("Mail_Flag", connstr)

            Dim PWIDNO As String = ""
            Dim chinese_name As String = ""
            Dim empemail As String = ""

            '找出填表者id
            db.Open()
            Dim peridCom As New SqlCommand("select PWIDNO from P_02 where eformsn = '" & eformsn & "'", db)
            Dim RdvId = peridCom.ExecuteReader()
            If RdvId.read() Then
                PWIDNO = RdvId.Item("PWIDNO")
            End If
            db.Close()

            '找出填表者資料
            db.Open()
            Dim perCom As New SqlCommand("select emp_chinese_name,empemail from employee where employee_id = '" & PWIDNO & "'", db)
            Dim Rdv = perCom.ExecuteReader()
            If Rdv.read() Then
                chinese_name = Rdv.Item("emp_chinese_name")
                empemail = Rdv.Item("empemail")
            End If
            db.Close()

            Dim strPath As String = ""
            strPath = "MOA00020.aspx?x=MOA02001&y=4rM2YFP73N&Read_Only=1&EFORMSN=" & eformsn

            '發送Mail給請假代理人
            Dim MailBody As String = ""
            MailBody += "您所申請的會議室申請單，已經被會議室管理者撤銷" & "<br>"
            MailBody += "<a href='" & MOAServer & strPath & "'>會議室申請單</a><br>"

            '判斷是否寄送Mail
            If MailYN = "Y" And empemail <> "" Then
                FC.F_MailGO(SystemMail, "系統通知", SmtpHost, empemail, "表單撤銷通知", MailBody)
            End If

        End If

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        Dim strOrd As String

        '排序條件
        strOrd = " ORDER BY " & GridView1.SortExpression.ToString()

        SqlDataSource1.SelectCommand = SQLALL(Meetsn, PAIDNO, SDate, EDate) & strOrd

    End Sub

    Protected Function FunStatus(ByVal str As String) As String
        Try
            '轉換表單狀態代號
            Dim tmpStr = Eval(str)

            If tmpStr = "" Then
                tmpStr = "審核中"
            ElseIf tmpStr = "E" Then
                tmpStr = "完成"
            ElseIf tmpStr = "0" Then
                tmpStr = "退件"
            End If

            FunStatus = tmpStr

        Catch ex As Exception
            FunStatus = ""
        End Try
    End Function

    Public Function SQLALL(ByVal Meetsn, ByVal PAIDNO, ByVal SDate, ByVal EDate)

        Dim sqlcom As String = "SELECT P_02.PAUNIT, P_02.PANAME, P_02.nPHONE, P_02.nMEETNAME, P_02.nPLACE, P_0204.MTT_Num, P_0204.EFORMSN, P_0204.MeetSn, CONVERT (char(12), P_0204.MeetTime, 111) AS MeetTime, P_0204.MeetHour, P_02.PENDFLAG FROM P_02,P_0204 WHERE P_02.EFORMSN = P_0204.EFORMSN "

        Dim sqlMeet As String = ""
        Dim sqlUser As String = ""
        Dim sqlDate As String = ""

        If Meetsn > 0 Then
            '會議室搜尋
            sqlMeet = " AND P_0204.MeetSn = " & Meetsn
        End If

        If PAIDNO <> "" Then
            '人員搜尋
            sqlUser = " AND P_02.PAIDNO='" & PAIDNO & "'"
        End If

        '申請日期
        If SDate <> "" And EDate <> "" Then
            '申請日期搜尋
            sqlDate = " AND (P_02.nAPPLYTIME between '" & SDate & " 00:00:00 ' AND '" & EDate & " 23:59:59')"
        End If

        SQLALL = sqlcom & sqlMeet & sqlUser & sqlDate

    End Function

End Class
