Imports System.Data.SqlClient
Partial Class Source_00_MOA00012
    Inherits System.Web.UI.Page
    Dim conn As New C_SQLFUN
    Dim connstr As String = conn.G_conn_string
    Dim db As New SqlConnection(connstr)
    Dim user_id, org_uid As String
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
            '判斷登入者權限
            Dim LoginCheck As New C_Public
            If LoginCheck.LoginCheck(user_id, "MOA00012") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00012.aspx")
                Response.End()
            End If
            If Not IsPostBack Then
                '先設定起始日期
                Dim dt As Date = Now()
                If (SDate.Text = "") Then
                    SDate.Text = dt.AddDays(-7).Date
                End If

                If (EDate.Text = "") Then
                    EDate.Text = dt.Date
                End If
            End If
        End If
    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete
        If Not IsPostBack Then
            Dim tLItm As New ListItem("請選擇", "")
            EformSel.Items.Insert(0, tLItm)
            EformSel.Items(1).Selected = False
            '登入馬上查詢
            Searchbtn_Click(Nothing, Nothing)
        End If
    End Sub

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        '分頁
        Dim strOrd As String = ""
        strOrd = " ORDER BY flowctl.appdate DESC"
        SqlDataSource2.SelectCommand = SQLALL(EformSel.SelectedValue, StateSel.SelectedValue, SDate.Text, EDate.Text) & strOrd
    End Sub
    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then
            '隱藏eformid,employee_id,eformsn,eformrole
            e.Row.Cells(4).Visible = False
            e.Row.Cells(5).Visible = False
            e.Row.Cells(6).Visible = False
            e.Row.Cells(7).Visible = False
        End If
    End Sub

    Protected Sub Searchbtn_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles Searchbtn.Click
        Dim strOrd As String = ""
        strOrd = " ORDER BY flowctl.appdate DESC"
        SqlDataSource2.SelectCommand = SQLALL(EformSel.SelectedValue, StateSel.SelectedValue, SDate.Text, EDate.Text) & strOrd
    End Sub

    Public Function SQLALL(ByVal sEformSel As String, ByVal sStateSel As String, ByVal sSDate As String, ByVal sEDate As String) As String
        '整合SQL搜尋字串
        Dim strsql As String = "SELECT flowctl.flowsn, flowctl.eformid, flowctl.eformrole, flowctl.appdate, flowctl.empuid, V_EformFlow.emp_chinese_name, V_EformFlow.frm_chinese_name, flowctl.eformsn, flowctl.comment,flowctl.gonogo+'-'+V_EformFlow.status as gonogo, flowctl.hddate FROM flowctl INNER JOIN V_EformFlow ON flowctl.eformsn = V_EformFlow.eformsn WHERE (flowctl.gonogo <> '-') AND (flowctl.gonogo <> 'G') and flowctl.empuid = '" & user_id & "' AND flowctl.gonogo +'-'+V_EformFlow.status IN ('E-E','0-0')"

        '查詢表單
        If sEformSel = "" Then
            strsql = strsql
        Else
            strsql += " and flowctl.eformid = '" & sEformSel & "'"
        End If

        If (sStateSel.Length > 0 AndAlso sStateSel <> "99") Then
            strsql += " and V_EformFlow.status='" + sStateSel + "'"
        End If

        '申請日期搜尋
        strsql += " AND (flowctl.appdate between '" & sSDate & " 00:00:00 ' AND '" & sEDate & " 23:59:59')"
        SQLALL = strsql
    End Function

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted
        '排序條件
        SqlDataSource2.SelectCommand = SQLALL(EformSel.SelectedValue, StateSel.SelectedValue, SDate.Text, EDate.Text)
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

        Dim streformid, streformsn As String
        Dim strPath As String = ""

        '顯示選取的表單資料
        streformid = GridView1.Rows(GridView1.SelectedIndex.ToString()).Cells(4).Text
        streformsn = GridView1.Rows(GridView1.SelectedIndex.ToString()).Cells(6).Text

        '表單資料夾
        Dim printRecordsReportID As String = GetEFormId("影印紀錄呈核單") '取得影印紀錄呈核單ID
        Dim doorAndMeetingControlID As String = GetEFormId("門禁會議管制申請單") '取得門禁會議管制申請單ID
        If streformid = "YAqBTxRP8P" Then       '請假申請單(DB內叫差假申請單)
            strPath = "MOA00020.aspx?x=MOA01001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        ElseIf streformid = "4rM2YFP73N" Then   '會議室申請單
            strPath = "MOA00020.aspx?x=MOA02001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        ElseIf streformid = "j2mvKYe3l9" Then   '派車申請單
            strPath = "MOA00020.aspx?x=MOA03001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        ElseIf streformid = "61TY3LELYT" Then   '房舍水電申請單
            strPath = "MOA00020.aspx?x=MOA04001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        ElseIf streformid = "4ZXNVRV8B6" Then   '完工報告單
            strPath = "MOA00020.aspx?x=MOA04003&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        ElseIf streformid = "U28r13D6EA" Then   '會客洽公申請單
            strPath = "MOA00020.aspx?x=MOA05001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        ElseIf streformid = "D6Y95Y5XSU" Then   '資訊設備媒體攜出入申請單
            strPath = "MOA00020.aspx?x=MOA06001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        ElseIf streformid = "9JKSDRR5V3" Then   '報修申請單
            strPath = "MOA00020.aspx?x=MOA07001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        ElseIf streformid = "5D82872F5L" Then   '銷假申請單 (銷假只有查詢，要撒銷請假單需於表單查詢出該請假單再做撒銷動作)
            strPath = "MOA00020.aspx?x=MOA01003&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        ElseIf streformid = "61TY3LELYT" Then '修繕申請單 (舊版的房舍水電修繕申請單)
            strPath = "MOA00020.aspx?x=MOA04001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        ElseIf streformid = "F9MBD7O97G" Then '修繕申請單 (博新案的房舍水電修繕申請單)
            strPath = "MOA00020.aspx?x=MOA04100&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        ElseIf streformid = "74BN58683M" Then '影印使用申請單
            strPath = "MOA00020.aspx?x=MOA08001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        ElseIf streformid = printRecordsReportID Then   '影印紀錄呈核單
            strPath = "MOA00020.aspx?x=MOA08014&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        ElseIf streformid = doorAndMeetingControlID Then   '門禁會議管制申請單
            strPath = "MOA00020.aspx?x=MOA09001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        ElseIf streformid = "BL7U2QP3IG" Then   '資訊設備維修申請單
            strPath = "MOA00020.aspx?x=MOA11001&y=" & streformid & "&Read_Only=1&EFORMSN=" & streformsn
        End If

        If strPath <> "" Then
            Response.Write(" <script language='javascript'>")
            Response.Write(" sPath = '" & strPath & "';")
            Response.Write(" strFeatures = 'dialogWidth=900px;dialogHeight=700px;help=no;status=no;resizable=yes;scroll=no;dialogTop=100;dialogLeft=100';")
            Response.Write(" window.showModalDialog(sPath,self,strFeatures);")
            Response.Write(" </script>")
            Searchbtn_Click(Nothing, Nothing)
        End If
    End Sub

    '取得表單種類ID
    Private Function GetEFormId(ByVal eformName As String) As String
        db.Open()
        Dim sqlcomm As New SqlCommand("SELECT eformid FROM EFORMS WHERE frm_chinese_name = '" + eformName + "'", db)
        GetEFormId = sqlcomm.ExecuteScalar()
        db.Close()
    End Function

    Protected Function FunStatus(ByVal str As String) As String
        Try
            '轉換表單狀態代號
            Dim tmpStr = Eval(str)
            Dim arrStatus As Array

            arrStatus = tmpStr.ToString.Split("-")

            Select Case arrStatus(1)
                Case "?"
                    tmpStr = "審核中"
                Case "E"
                    tmpStr = "完成"
                Case "0"
                    tmpStr = "駁回"
                Case "B"
                    tmpStr = "撤銷"
                Case Else
                    tmpStr = "不明狀態"
            End Select
           
            'If arrStatus(0) = "?" And arrStatus(1) = "?" Then
            '    tmpStr = "未批核"
            'ElseIf (arrStatus(0) = "1" And arrStatus(1) = "E") Or (arrStatus(0) = "E" And arrStatus(1) = "E") Then
            '    tmpStr = "完成"
            'ElseIf arrStatus(0) = "0" And arrStatus(1) = "0" Then
            '    tmpStr = "退件"
            'Else
            '    tmpStr = "審核中"
            'End If

            FunStatus = tmpStr

        Catch ex As Exception
            FunStatus = ""
        End Try
    End Function

    Protected Sub GridView1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim eformsn As String
        Dim tool = New C_Public
        Dim gv As GridView = CType(sender, GridView)
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                eformsn = e.Row.Cells(6).Text
                e.Row.Cells(8).Text = tool.GetFinalCaseComment(eformsn)

                Select Case tool.GetFinalStatus(eformsn)
                    Case "?"
                        e.Row.Cells(9).Text = "審核中"
                    Case "E"
                        e.Row.Cells(9).Text = "完成"
                    Case "0"
                        e.Row.Cells(9).Text = "駁回"
                    Case "B"
                        e.Row.Cells(9).Text = "撤銷"
                    Case Else
                        e.Row.Cells(9).Text = "不明狀態"
                End Select
        End Select

    End Sub
End Class
