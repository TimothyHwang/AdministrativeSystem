Imports System.Data.SqlClient
Partial Class Source_01_MOA01004
    Inherits System.Web.UI.Page

    Dim conn As New C_SQLFUN
    Dim eformid, user_id, org_uid, streformsn, connstr, eformsnOne As String
    Dim Org_Name, User_Name, Title_Name As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        user_id = Session("user_id")
        org_uid = Session("ORG_UID")

        'session被清空回首頁
        If user_id = "" Or org_uid = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")

        Else

            Try

                Dim C_Public As New C_Public
                'eformsnOne = C_Public.randstr(16)            '建立唯一的eformsn
                eformsnOne = C_Public.CreateNewEFormSN()      '加入重覆防呆功能 20130710 paul

                connstr = conn.G_conn_string

                '開啟連線
                Dim db As New SqlConnection(connstr)

                '登入者資料
                db.Open()
                Dim strPer As New SqlCommand("SELECT ORG_NAME,emp_chinese_name,AD_TITLE FROM V_EmpInfo WHERE employee_id = '" & user_id & "'", db)
                Dim RdPer = strPer.ExecuteReader()
                If RdPer.read() Then
                    Org_Name = RdPer("ORG_NAME")
                    User_Name = RdPer("emp_chinese_name")
                    Title_Name = RdPer("AD_TITLE")
                End If
                db.Close()

                If IsPostBack = False Then

                    '設定Default日期
                    Dim Sdate As TextBox = Me.FindControl("Sdate")
                    Dim Edate As TextBox = Me.FindControl("Edate")

                    Dim dt As Date = Now()
                    If (Sdate.Text = "") Then
                        Sdate.Text = dt.AddDays(-14).Date
                    End If

                    If (Edate.Text = "") Then
                        Edate.Text = dt.Date
                    End If

                    SqlDataSource2.SelectCommand = SQLALL(DDLType.SelectedValue, Sdate.Text, Edate.Text) & " ORDER BY nAPPTIME DESC"

                End If

            Catch ex As Exception

            End Try


        End If


    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        If Not IsPostBack Then

            Dim tLItm As New ListItem("請選擇", "")

            '假別加請選擇
            DDLType.Items.Insert(0, tLItm)
            If DDLType.Items.Count > 1 Then
                DDLType.Items(1).Selected = False
            End If

        End If

    End Sub

    Public Function SQLALL(ByVal TypeSel, ByVal SDate, ByVal EDate)

        '整合SQL搜尋字串
        Dim strsql As String = ""
        Dim strsel As String = ""
        Dim strDate As String = ""
        strsql = "SELECT P_NUM,EFORMSN,PAIDNO,nAPPTIME,nTYPE,CONVERT (char(12), nSTARTTIME, 111) AS nSTARTTIME,nSTHOUR,CONVERT (char(12), nENDTIME, 111) AS nENDTIME,nETHOUR,nREASON,PENDFLAG FROM P_01 WHERE PENDFLAG='E' AND (EFORMSN NOT IN (SELECT nEFORMSN FROM P_0101 WHERE PENDFLAG = 'E')) AND PAIDNO='" & user_id & "'"

        '假別搜尋
        If TypeSel <> "" Then
            strsel = " AND nTYPE = '" & TypeSel & "'"
        End If

        '申請日期搜尋

        If SDate <> "" And EDate <> "" Then
            strDate = " AND (nAPPTIME between '" & SDate & " 00:00:00 ' AND '" & EDate & " 23:59:59')"
        End If

        SQLALL = strsql & strsel & strDate

    End Function

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click

        Dim strOrd As String

        strOrd = " ORDER BY nAPPTIME DESC"

        SqlDataSource2.SelectCommand = SQLALL(DDLType.SelectedValue, Sdate.Text, Edate.Text) & strOrd

    End Sub

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged

        '分頁
        Dim strOrd As String

        strOrd = " ORDER BY nAPPTIME DESC"

        SqlDataSource2.SelectCommand = SQLALL(DDLType.SelectedValue, Sdate.Text, Edate.Text) & strOrd

    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated

        If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then

            '隱藏eformsn
            e.Row.Cells(7).Visible = False

        End If

    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

        '撤銷請假單

        'eformsn
        streformsn = GridView1.Rows(GridView1.SelectedIndex).Cells(7).Text

        'eformid
        eformid = "5D82872F5L"

        '表單審核
        Dim FC As New C_FlowSend.C_FlowSend
        Dim SendVal = eformid & "," & user_id & "," & eformsnOne & "," & "1" & ","

        Dim NextPer As Integer = 0

        '關卡為上一級主管有多少人
        NextPer = FC.F_NextStep(SendVal, connstr)

        '表單送件
        Dim Val_P As String
        Val_P = ""

        '判斷下一關為上一級主管時人數是否超過一人
        If NextPer > 1 Then
            Server.Transfer("../00/MOA00013.aspx?eformid=" & eformid & "&eformsn=" & eformsnOne & "&SelFlag=1&CancelFlag=1&eformsnOld=" & streformsn)
        Else

            '新增撤銷單資料
            InsData()

            Val_P = FC.F_Send(SendVal, connstr)

            Response.Redirect("../00/MOA00007.aspx?val=" & Val_P)

        End If

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        '排序條件
        Dim strOrd As String

        strOrd = " ORDER BY " & GridView1.SortExpression.ToString()

        SqlDataSource2.SelectCommand = SQLALL(DDLType.SelectedValue, Sdate.Text, Edate.Text) & strOrd

    End Sub

    Public Sub InsData()

        '連線字串
        connstr = conn.G_conn_string

        '新增連線
        Dim db As New SqlConnection(connstr)

        '開啟連線
        db.Open()

        '填表人和申請人同為登入者資料
        Dim InsP02 As String = ""
        '表單序號,填表人單位,填表人級職,填表人姓名,填表人身份證字號,申請人單位,申請人姓名,申請人級職,申請人身份證字號
        '撤銷表單編號
        InsP02 = "INSERT INTO P_0101(EFORMSN,PWUNIT,PWTITLE,PWNAME,PWIDNO,PAUNIT,PANAME, PATITLE, PAIDNO "
        InsP02 += ",nEFORMSN) "
        InsP02 += " VALUES ('" & eformsnOne & "','" & Org_Name & "','" & Title_Name & "',N'" & User_Name & "','" & user_id & "','" & Org_Name & "',N'" & User_Name & "','" & Title_Name & "','" & user_id & "'"
        InsP02 += ",'" & streformsn & "')"

        Dim insCom As New SqlCommand(InsP02, db)
        insCom.ExecuteNonQuery()
        db.Close()

    End Sub

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "70px"
        Div_grid.Style("left") = "340px"

        Calendar1.SelectedDate = Sdate.Text

    End Sub

    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click

        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "70px"
        Div_grid2.Style("left") = "480px"

        Calendar2.SelectedDate = Edate.Text

    End Sub
    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

        Sdate.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False

    End Sub

    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged

        Edate.Text = Calendar2.SelectedDate.Date
        Div_grid2.Visible = False

    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

        Div_grid.Visible = False

    End Sub

    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click

        Div_grid2.Visible = False

    End Sub

End Class
