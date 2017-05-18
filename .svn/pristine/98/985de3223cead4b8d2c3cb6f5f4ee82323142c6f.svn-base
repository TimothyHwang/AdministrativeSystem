Imports System.Data.SqlClient
Partial Class Source_00_MOA00106
    Inherits System.Web.UI.Page

    Dim connstr As String
    Public do_sql As New C_SQLFUN
    Dim user_id, org_uid As String

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

                '判斷登入者權限
                Dim LoginCheck As New C_Public

                If LoginCheck.LoginCheck(user_id, "MOA00106") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00106.aspx")
                    Response.End()
                End If


                If Not IsPostBack Then

                    Dim p As Integer = 0

                    '計算年度
                    Dim b_year As Integer = CInt(DatePart(DateInterval.Year, Now) - 4)
                    Dim e_year As Integer = CInt(DatePart(DateInterval.Year, Now))

                    For p = b_year To e_year
                        DrDown_year.Items.Add(p)
                    Next

                    '選擇起始年度
                    DrDown_year.SelectedIndex = 2

                End If

                connstr = do_sql.G_conn_string

                '開啟連線
                Dim db As New SqlConnection(connstr)

                '現有流程數量
                db.Open()
                Dim strCNow As New SqlCommand("SELECT count(DISTINCT eformsn) as flowCount FROM flowctl", db)
                Dim RdCNow = strCNow.ExecuteReader()
                If RdCNow.read() Then
                    LabCNow.Text = RdCNow("flowCount")
                End If
                db.Close()

                '歷史流程數量
                db.Open()
                Dim strCOld As New SqlCommand("SELECT count(DISTINCT eformsn) as flowCount FROM flowctl_HIS", db)
                Dim RdCOld = strCOld.ExecuteReader()
                If RdCOld.read() Then
                    LabCOld.Text = RdCOld("flowCount")
                End If
                db.Close()

                '顯示歷史資料
                Dim numrowsold As Integer = 2
                Dim numcellsold As Integer = 6
                Dim k As Integer
                For k = 0 To numrowsold - 1
                    Dim n As New TableRow()
                    Dim m As Integer
                    For m = -1 To numcellsold - 1

                        '年度
                        Dim strYear As Integer = DatePart(DateInterval.Year, Now) - m

                        Dim strFlowNow As String = ""

                        '現有流程數量
                        db.Open()
                        Dim strYearNow As New SqlCommand("SELECT count(DISTINCT eformsn) as flowCount FROM flowctl_HIS where datepart(yyyy,appdate) = " & strYear, db)
                        Dim RdYearNow = strYearNow.ExecuteReader()
                        If RdYearNow.read() Then
                            strFlowNow = RdYearNow("flowCount")
                        End If
                        db.Close()


                        Dim c As New TableCell()
                        If k = 0 Then
                            If m = -1 Then
                                c.Controls.Add(New LiteralControl("年度"))
                            Else
                                c.Controls.Add(New LiteralControl(strYear - 1911 & "年"))
                            End If
                        Else
                            If m = -1 Then
                                c.Controls.Add(New LiteralControl("歷史資料筆數"))
                            Else
                                c.Controls.Add(New LiteralControl(strFlowNow))
                            End If
                        End If

                        n.Cells.Add(c)
                    Next m
                    TabOld.Rows.Add(n)
                Next k

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub BtnSel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnSel.Click

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '歷史資料批核完成流程數量
        db.Open()
        Dim strYearOld As New SqlCommand("SELECT count(DISTINCT eformsn) as flowCount FROM flowctl where datepart(yyyy,appdate) = " & DrDown_year.SelectedValue & "  AND (gonogo <> '-') AND (hddate IS NOT NULL) ", db)
        Dim RdYearOld = strYearOld.ExecuteReader()
        If RdYearOld.read() Then
            LabStand.Text = RdYearOld("flowCount")
        End If
        db.Close()

        '歷史資料未批核完成流程數量
        db.Open()
        Dim strYearNow As New SqlCommand("SELECT count(DISTINCT eformsn) as flowCount FROM flowctl where datepart(yyyy,appdate) = " & DrDown_year.SelectedValue & " AND (gonogo <> '-') AND (hddate IS NULL) ", db)
        Dim RdYearNow = strYearNow.ExecuteReader()
        If RdYearNow.read() Then
            LabStandErr.Text = RdYearNow("flowCount")
        End If
        db.Close()

    End Sub

    Protected Sub but_exe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles but_exe.Click

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '新增歷史資料
        db.Open()
        Dim insCom As New SqlCommand("insert into FlowCTL_his select * from FlowCTL where datepart(yyyy,appdate) <= '" + DrDown_year.SelectedValue + "' and eformsn not in (select f.eformsn from flowctl f where hddate is null)", db)
        insCom.ExecuteNonQuery()
        db.Close()

        '刪除歷史資料
        db.Open()
        Dim delCom As New SqlCommand("delete from FlowCTL where eformsn in (SELECT f.eformsn from FlowCTL f where datepart(yyyy,appdate) = " & DrDown_year.SelectedValue & " AND (gonogo <> '-') AND (hddate IS NOT NULL) )", db)
        delCom.ExecuteNonQuery()
        db.Close()

        Server.Transfer("MOA00106.aspx")

    End Sub
End Class
