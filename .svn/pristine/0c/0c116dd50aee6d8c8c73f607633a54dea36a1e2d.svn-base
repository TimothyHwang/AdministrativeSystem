Imports System.Data
Imports System.IO
Imports System.Data.SqlClient
Partial Class Source_02_MOA02003
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String
    Public do_sql As New C_SQLFUN
    Dim dr As System.Data.DataRow
    Dim n_table As New System.Data.DataTable
    Public print_file As String = ""
    Dim STMT As String = ""

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        '增加請選擇選項
        If Not IsPostBack Then
            Dim tLItm As New ListItem("請選擇", "")
            MeetSn.Items.Insert(0, tLItm)
            If MeetSn.Items.Count > 1 Then
                MeetSn.Items(1).Selected = False
            End If
        End If

    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click

        Dim strOrd As String

        strOrd = " ORDER BY P_0201.MeetName"

        SqlDataSource2.SelectCommand = SQLALL(MeetSn.SelectedValue, MeetOrg.Text, BorrowPer.Text, SDate.Text, EDate.Text) & strOrd

    End Sub

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

            If IsPostBack = False Then

                '先設定起始日期
                Dim dt As Date = Now()
                If (SDate.Text = "") Then
                    SDate.Text = dt.AddDays(-14).Date
                End If

                If (EDate.Text = "") Then
                    EDate.Text = dt.Date
                End If

                '系統管理員可以看全部的會議室
                If Session("Role") = "1" Then
                    SqlDataSource1.SelectCommand = "SELECT MeetSn, MeetName FROM P_0201 ORDER BY MeetName"
                Else
                    SqlDataSource1.SelectCommand = "SELECT MeetSn, MeetName FROM P_0201 WHERE (P_0201.Owner = '" & user_id & "') ORDER BY MeetName"
                End If

                ImgSearch_Click(Nothing, Nothing)

            End If

        End If


    End Sub

    Public Function SQLALL(ByVal MeetSn, ByVal MeetOrg, ByVal BorrowPer, ByVal SDate, ByVal EDate)

        Dim sqlsel As String = ""
        Dim sqlOrg As String = ""
        Dim sqlPer As String = ""
        Dim strdate As String = ""
        Dim sqlcom As String = "SELECT P_0201.MeetSn, P_0201.Org_Uid, P_0201.MeetName, P_0201.Owner, P_0201.Share, COUNT(P_0204.EFORMSN) AS UseCount, P_02.PAUNIT, P_02.PANAME,P_02.PAIDNO FROM P_0204 INNER JOIN P_02 ON P_0204.EFORMSN = P_02.EFORMSN RIGHT OUTER JOIN P_0201 ON P_0204.MeetSn = P_0201.MeetSn  "
        Dim sqGroup As String = " GROUP BY P_0201.MeetSn, P_0201.Org_Uid, P_0201.MeetName, P_0201.Owner, P_0201.Share, P_02.PAUNIT, P_02.PANAME,P_02.PAIDNO "
        Dim strhaving As String = ""

        '系統管理員可以看全部的會議室
        If Session("Role") = "1" Then
            sqlcom = sqlcom

            strhaving = " HAVING (1=1) "
        Else
            '會議室管理員也可以查詢
            strhaving = " HAVING (P_0201.Owner = '" & user_id & "') "
        End If

        '申請日期
        If SDate <> "" And EDate <> "" Then
            '申請日期搜尋
            strdate = " AND (P_02.nAPPLYTIME between '" & SDate & " 00:00:00 ' AND '" & EDate & " 23:59:59')"
        End If

        If MeetSn <> "" Then
            '會議室搜尋
            sqlsel = " AND P_0201.MeetSn = '" & MeetSn & "'"
        End If

        If MeetOrg <> "" Then
            '單位搜尋
            sqlOrg = " AND P_02.PAUNIT like '%" & Trim(MeetOrg) & "%'"
        End If

        If BorrowPer <> "" Then
            '人員搜尋
            sqlPer = " AND P_02.PANAME like '%" & Trim(BorrowPer) & "%'"
        End If

        SQLALL = sqlcom & strdate & sqGroup & strhaving & sqlsel & sqlOrg & sqlPer

    End Function

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged

        '分頁
        Dim strOrd As String

        strOrd = " ORDER BY P_0201.MeetName"

        SqlDataSource2.SelectCommand = SQLALL(MeetSn.SelectedValue, MeetOrg.Text, BorrowPer.Text, SDate.Text, EDate.Text) & strOrd

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        Dim strOrd As String

        '排序條件
        strOrd = " ORDER BY " & GridView1.SortExpression.ToString()

        SqlDataSource2.SelectCommand = SQLALL(MeetSn.SelectedValue, MeetOrg.Text, BorrowPer.Text, SDate.Text, EDate.Text) & strOrd

    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

        Dim strMeetSn As String
        Dim strPAIDNO As String
        Dim strPath As String = ""

        '顯示選取的表單資料
        strMeetSn = GridView1.Rows(GridView1.SelectedIndex.ToString()).Cells(5).Text
        strPAIDNO = GridView1.Rows(GridView1.SelectedIndex.ToString()).Cells(6).Text

        '開啟的頁面路徑
        strPath = "MOA02008.aspx?MeetSn=" & strMeetSn & "&PAIDNO=" & strPAIDNO & "&SDate=" & SDate.Text & "&EDate=" & EDate.Text

        '顯示被選取的會議室詳細資料
        Server.Transfer(strPath)


    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated

        If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then

            '隱藏MeetSn,PAIDNO
            e.Row.Cells(5).Visible = False
            e.Row.Cells(6).Visible = False

        End If
    End Sub

    'Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

    '    Div_grid.Visible = True
    '    Div_grid.Style("Top") = "95px"
    '    Div_grid.Style("left") = "90px"

    '    Calendar1.SelectedDate = Sdate.Text

    'End Sub

    'Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click

    '    Div_grid2.Visible = True
    '    Div_grid2.Style("Top") = "95px"
    '    Div_grid2.Style("left") = "210px"

    '    Calendar2.SelectedDate = Edate.Text

    'End Sub

    'Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

    '    Sdate.Text = Calendar1.SelectedDate.Date
    '    Div_grid.Visible = False

    'End Sub

    'Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged

    '    Edate.Text = Calendar2.SelectedDate.Date
    '    Div_grid2.Visible = False

    'End Sub

    'Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

    '    Div_grid.Visible = False

    'End Sub

    'Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click

    '    Div_grid2.Visible = False

    'End Sub

    Protected Sub ImagePrint_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImagePrint.Click
        Dim F_file As String
        Dim F_file2 As String
        Dim F_file_name As String
        

        Dim strOrd As String

        strOrd = " ORDER BY P_0201.MeetName"

        STMT = SQLALL(MeetSn.SelectedValue, MeetOrg.Text, BorrowPer.Text, SDate.Text, EDate.Text) & strOrd


        If do_sql.db_sql(STMT, do_sql.G_conn_string) = False Then
            do_sql.G_errmsg = "查詢資料失敗"
            Exit Sub
        End If
        If do_sql.G_table.Rows.Count > 0 Then
            n_table = do_sql.G_table
        Else
            do_sql.G_errmsg = "查無可印資料"
            Exit Sub
        End If



        Dim prn_stmt As String

        F_file_name = "rpt020030"
        F_file = MapPath("../../rpt/" + F_file_name + ".txt")
        If File.Exists(F_file) = False Then
            do_sql.G_errmsg = "檔案不存在"
            Exit Sub
        End If

        print_file = "../../drs/" + F_file_name + "-" + Now.ToString("mmssff") & ".drs"
        F_file2 = MapPath(print_file)

        Call do_sql.inc_file(F_file, F_file2, F_file_name)
        Dim tmp_month As String = ""
        Dim nPage As Integer = 0
        Dim n_line As Integer = 35
        For Each dr In n_table.Rows
            If n_line >= 34 Then
                nPage += 1
                Call do_sql.print_sdata(F_file2, "/init " + F_file_name + ".txt 1", "/newpage null")
                prn_stmt = CDate(SDate.Text).ToString("yyyy/MM/dd") & "~" & CDate(EDate.Text).ToString("yyyy/MM/dd")
                Call do_sql.print_block(F_file2, "日期", 0, 0, prn_stmt)
                n_line = 0
            End If
            Call do_sql.print_block(F_file2, "會議室名稱", 0, n_line * 7, dr("MeetName").ToString)
            Call do_sql.print_block(F_file2, "借用部門", 0, n_line * 7, dr("PAUNIT").ToString)
            Call do_sql.print_block(F_file2, "借用人", 0, n_line * 7, dr("PANAME").ToString)
            Call do_sql.print_block(F_file2, "借用次數", 0, n_line * 7, dr("UseCount").ToString)
            n_line = n_line + 1
        Next

    End Sub

    Protected Sub Img_Export_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles Img_Export.Click
        Dim strsqlAll As String = SQLALL(MeetSn.SelectedValue, MeetOrg.Text, BorrowPer.Text, SDate.Text, EDate.Text) & " ORDER BY P_0201.MeetName"

        '設定檔案路徑
        Dim path As String = "", filename As String = ""

        path = Server.MapPath("~/Drs/")
        filename = Date.Today.ToString("yyyyMMdd") & "_02003.csv"
        path = path & filename

        Dim colname As String = ""
        Dim data As String = ""
        Dim prn_stmt As String = ""

        Dim sw As New System.IO.StreamWriter(path, False, Encoding.GetEncoding("big5"))

        prn_stmt = "會議室使用統計表,," & CDate(SDate.Text).ToString("yyyy/MM/dd") & "~" & CDate(EDate.Text).ToString("yyyy/MM/dd")

        sw.WriteLine(prn_stmt)

        colname = "會議室名稱,借用部門,借用人,借用次數,"

        colname = Left(colname, Len(colname) - 1)

        sw.WriteLine(colname)


        Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

        db.Open()
        Dim dt2 As DataTable = New DataTable("")
        Dim da2 As SqlDataAdapter = New SqlDataAdapter(strsqlAll, db)
        da2.Fill(dt2)
        db.Close()
        For y As Integer = 0 To dt2.Rows.Count - 1

            data += dt2.Rows(y).Item("MeetName").ToString & "," & dt2.Rows(y).Item("PAUNIT").ToString & ","
            data += dt2.Rows(y).Item("PANAME").ToString & "," & dt2.Rows(y).Item("UseCount").ToString & ","
            data = Left(data, Len(data) - 1)
            sw.WriteLine(data)
            data = ""

        Next

        sw.WriteLine("")
        sw.Close()

        '匯出檔案
        Response.Clear()
        Response.ContentType = "application/download"
        Response.AddHeader("content-disposition", "attachment;filename=" & filename)
        Response.WriteFile(path)
        Response.End()
    End Sub
End Class
