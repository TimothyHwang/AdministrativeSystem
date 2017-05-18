Imports System.Data
Imports System.Data.SqlClient
Partial Class Source_01_MOA01003
    Inherits System.Web.UI.Page

    Dim conn As New C_SQLFUN
    Public read_only As String = ""
    Dim eformid, user_id, org_uid, streformsn, connstr As String
    Dim OLDEformsn As String
    Dim AgentEmpuid As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        '取得登入者帳號
        If Page.User.Identity.Name.ToString.IndexOf("\") > 0 Then

            Dim LoginAll As String = Page.User.Identity.Name.ToString

            Dim LoginID() As String = Split(LoginAll, "\")

            user_id = LoginID(1)
        Else
            user_id = Page.User.Identity.Name.ToString
        End If

        'user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        eformid = Request.QueryString("eformid")
        streformsn = Request.QueryString("eformsn")
        read_only = Request.QueryString("read_only")

        'session被清空回首頁
        If user_id = "" And org_uid = "" And streformsn = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.parent.location='../../index.aspx';")
            Response.Write(" </script>")

        Else

            '判斷選擇哪個填表人
            Try

                connstr = conn.G_conn_string

                '開啟連線
                Dim db As New SqlConnection(connstr)

                If read_only = "2" Then
                    send.Text = "核准"
                    '判斷是否由入口網站送件
                    If Session("user_id") = "" Then
                        user_id = Page.User.Identity.Name.Substring(Page.User.Identity.Name.IndexOf("\") + 1)
                    End If
                End If

                '找出表單審核者
                If read_only = "2" Then

                    db.Open()
                    Dim strAgentCheck As New SqlCommand("SELECT empuid FROM flowctl WHERE eformsn = '" & streformsn & "' and hddate is null", db)
                    Dim RdrAgentCheck = strAgentCheck.ExecuteReader()
                    If RdrAgentCheck.read() Then
                        AgentEmpuid = RdrAgentCheck.item("empuid")
                    End If
                    db.Close()

                End If

                '填表人和申請人資料
                db.Open()
                Dim strPer As New SqlCommand("SELECT ORG_NAME,emp_chinese_name,AD_Title FROM V_EmpInfo WHERE employee_id = '" & user_id & "'", db)
                Dim RdPer = strPer.ExecuteReader()
                If RdPer.read() Then
                    Lab_ORG_NAME_1.Text = RdPer("ORG_NAME")
                    Lab_emp_chinese_name.Text = RdPer("emp_chinese_name")
                    Lab_title_name_1.Text = RdPer("AD_Title")
                    Lab_ORG_NAME_2.Text = RdPer("ORG_NAME")
                    Lab_emp_chinese_name2.Text = RdPer("emp_chinese_name")
                    Lab_title_name_2.Text = RdPer("AD_Title")
                End If
                db.Close()

            Catch ex As Exception

            End Try


        End If


    End Sub

    Protected Sub send_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles send.Click

        Try

            '新增資料
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '表單送件
            Dim Val_P As String
            Val_P = ""

            '表單審核
            Dim FC As New C_FlowSend.C_FlowSend

            Dim SendVal As String = ""

            '判斷是否為代理人批核的表單
            If AgentEmpuid = "" Then
                SendVal = eformid & "," & user_id & "," & streformsn & "," & "1" & ","
            Else
                SendVal = eformid & "," & AgentEmpuid & "," & streformsn & "," & "1" & ","
            End If

            Dim NextPer As Integer = 0

            '關卡為上一級主管有多少人
            NextPer = FC.F_NextStep(SendVal, connstr)

            Dim Chknextstep As Integer

            '判斷表單關卡
            db.Open()
            Dim strComCheck As New SqlCommand("SELECT nextstep FROM flowctl WHERE eformsn = '" & streformsn & "' and empuid = '" & user_id & "' and hddate is null", db)
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

            If read_only = "2" Then

                Dim strAgentName As String = ""

                '判斷是否為代理人批核
                If UCase(user_id) = UCase(AgentEmpuid) Then

                    '增加批核意見
                    insComment(txtcomment.Text, streformsn, user_id)

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
                    insComment(strComment, streformsn, AgentEmpuid)

                End If

            End If

            '判斷下一關為上一級主管時人數是否超過一人
            If (NextPer > 1 And read_only = "") Or (NextPer > 1 And strgroup_id = "Z860") Then
                Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & streformsn & "&SelFlag=1")
            Else
                Val_P = FC.F_Send(SendVal, connstr)

                Dim PageUp As String = ""

                If read_only = "" Then
                    PageUp = "New"
                End If

                Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=" & PageUp)
            End If



        Catch ex As Exception

        End Try

    End Sub

    Protected Sub backBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles backBtn.Click

        Try

            Dim Val_P As String
            Val_P = ""

            '表單駁回
            Dim FC As New C_FlowSend.C_FlowSend

            Dim SendVal As String = ""

            '判斷是否為代理人批核的表單
            If AgentEmpuid = "" Then
                SendVal = eformid & "," & user_id & "," & streformsn & "," & "1"
            Else
                SendVal = eformid & "," & AgentEmpuid & "," & streformsn & "," & "1"
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

            Dim Val_P As String
            Val_P = ""

            '表單呈轉
            Dim FC As New C_FlowSend.C_FlowSend

            Dim SendVal As String = ""

            '判斷是否為代理人批核的表單
            If AgentEmpuid = "" Then
                SendVal = eformid & "," & user_id & "," & streformsn & "," & "1"
            Else
                SendVal = eformid & "," & AgentEmpuid & "," & streformsn & "," & "1"
            End If

            Val_P = FC.F_Transfer(SendVal, connstr)

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('表單已呈轉給上一級主管');")
            '重新整理頁面
            Response.Write(" window.dialogArguments.location='../00/MOA00010.aspx';")
            Response.Write(" window.close();")
            Response.Write(" </script>")

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '單純讀取表單不可送件
        If read_only = "1" Then

            send.Visible = False
            backBtn.Visible = False
            tranBtn.Visible = False
            ShowDetail(streformsn)

            '讀取表單可送件
        ElseIf read_only = "2" Then
            ShowDetail(streformsn)

            Dim PerCount As Integer = 0

            '判斷批核者是否為副主管,假如批核者為副主管則不可送件只可呈轉
            Dim ParentFlag As String = ""
            Dim ParentVal = user_id & "," & streformsn

            Dim FC As New C_FlowSend.C_FlowSend
            ParentFlag = FC.F_NextChief(ParentVal, connstr)

            If ParentFlag = "1" Then
                'send.Visible = False

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
                    send.Visible = True
                    tranBtn.Visible = False
                Else
                    send.Visible = True
                    'send.Visible = False
                End If

            End If

            '找出同級單位以下全部單位
            Dim Org_Down As New C_Public

            '沒有批核權限不可執行動作
            If (Org_Down.ApproveAuth(streformsn, user_id)) = "" Then
                send.Visible = False
                backBtn.Visible = False
                tranBtn.Visible = False
            End If

        Else
            backBtn.Visible = False
            tranBtn.Visible = False
        End If

    End Sub

    Public Function ShowDetail(ByVal eformsn As String)

        '判斷選擇哪個申請人
        Try

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '撤銷申請單資料
            db.Open()
            Dim strMeet As New SqlCommand("SELECT * FROM P_0101 WHERE EFORMSN = '" & eformsn & "'", db)
            Dim RdMeet = strMeet.ExecuteReader()

            If RdMeet.read() Then

                Lab_ORG_NAME_1.Text = RdMeet("PWUNIT")
                Lab_ORG_NAME_2.Text = RdMeet("PAUNIT")
                Lab_emp_chinese_name.Text = RdMeet("PWNAME")
                Lab_emp_chinese_name2.Text = RdMeet("PWNAME")
                Lab_title_name_1.Text = RdMeet("PWTITLE")
                Lab_title_name_2.Text = RdMeet("PATITLE")


                AppDate.Text = RdMeet("nAPPTIME")

                '申請單eformsn
                OLDEformsn = RdMeet("nEFORMSN")

                SqlDataSource3.SelectCommand = "SELECT [nAPPTIME], [nTYPE], CONVERT (char(12), nSTARTTIME, 111) AS nSTARTTIME, [nSTHOUR], CONVERT (char(12), nENDTIME, 111) AS nENDTIME, [nETHOUR], [nREASON] FROM [P_01] WHERE Eformsn = '" & OLDEformsn & " '"

            End If
            db.Close()

        Catch ex As Exception

        End Try

        ShowDetail = ""

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
            Div_grid10.Style("Top") = "280px"
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
