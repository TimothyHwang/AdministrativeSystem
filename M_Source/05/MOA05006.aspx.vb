Imports System.Data.SqlClient
Imports System.Data
Imports System.IO
Partial Class M_Source_05_MOA05006
    Inherits System.Web.UI.Page

    Public do_sql As New C_SQLFUN
    Dim dr As System.Data.DataRow
    Dim n_table As New System.Data.DataTable
    Public print_file As String = ""
    Dim conn As New C_SQLFUN
    Dim user_id, org_uid As String
    Dim STMT As String = ""
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim tool As New C_Public
        user_id = Session("user_id")
        org_uid = Session("ORG_UID")

        Dim connstr As String = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '設定Default日期
        Dim Sdate As TextBox = Me.FindControl("Sdate")
        Dim Edate As TextBox = Me.FindControl("Edate")

        Dim dt As Date = Now()
        If (Sdate.Text = "") Then
            Sdate.Text = dt.AddMonths(-1).ToString("yyyy/MM/") & "01"
        End If

        If (Edate.Text = "") Then
            Edate.Text = CDate(dt.ToString("yyyy/MM/") & "01").AddDays(-1).ToString("yyyy/MM/dd")
        End If

        ''判斷登入者是否為會客相關群組人員
        Dim VistorFlag As Boolean = False
        VistorFlag = tool.CheckRoleGroupEMPByName("會客管制群組", user_id)
        'Dim sqlVistor As String = ""
        'db.Open()
        'Dim carCom As New SqlCommand("select object_uid from SYSTEMOBJUSE where employee_id = '" & user_id & "' AND ((object_uid = 'E539') OR (object_uid = 'Y965'))", db)
        'Dim RdvCar = carCom.ExecuteReader()
        'If RdvCar.read() Then
        '    VistorFlag = RdvCar("object_uid")
        'End If
        'db.Close()

        'If VistorFlag = "E539" Then
        '    sqlVistor = " and State_Name='第一會客室' "
        'ElseIf VistorFlag = "Y965" Then
        '    sqlVistor = " and State_Name='第二會客室' "
        'End If

        ''判斷登入者權限
        If Session("Role") = "1" Or VistorFlag = True Then
            '會客室管理者以及系統管理者可以查看全部會客內容
            SqlDataSource1.SelectCommand = "select State_Name from SYSKIND where Kind_Num =3 AND STATE_ENABLED='1'"
            'ElseIf VistorFlag <> "" Then
            'SqlDataSource1.SelectCommand = "select State_Name from SYSKIND where Kind_Num =3 " & sqlVistor
        Else
            SqlDataSource1.SelectCommand = "select State_Name from SYSKIND where 1=2"
        End If

    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click

        '查詢
        Dim strOrd As String = ""

        strOrd = " ORDER BY D.nRECDATE "

        SqlDataSource2.SelectCommand = SQLALL(CDate(Sdate.Text).ToString("yyyy/MM/dd"), CDate(Edate.Text).ToString("yyyy/MM/dd")) & strOrd

    End Sub

    Public Function SQLALL(ByVal SDate, ByVal EDate)
        Dim pi As Integer = 0
        Dim t_date As String = ""

        STMT = "SELECT D.nRECDATE,case DWEEK when 1 then '一' when 2 then '二' when 3 then '三' when 4 then '四' when 5 then '五' "
        STMT += " when 6 then '六' when 7 then '日' end As DWEEK,ISNULL(PN01,0) PN01,ISNULL(PN02,0) aS PN02,ISNULL(PN03,0) aS PN03,ISNULL(PN02,0)+ISNULL(PN03,0) As PN04 FROM ( "
        pi = 0
        Do While CDate(SDate).AddDays(pi).ToString("yyyy/MM/dd") <= EDate
            t_date = CDate(SDate).AddDays(pi).ToString("yyyy/MM/dd")
            If pi > 0 Then
                STMT += " union "
            End If
            STMT += "SELECT '" & t_date & "' aS nRECDATE,"
            STMT += "replace(DATEPART(DW,'" & t_date & "') -1,0,7) DWEEK  "
            'STMT += DateAdd()
            pi += 1
        Loop

        STMT += ") D LEFT OUTER JOIN "

        STMT += "(SELECT a.nRECDATE,PN01,ISNULL(PN02,0) aS PN02,ISNULL(PN03,0) aS PN03 FROM ( "
        STMT += "(SELECT CONVERT(nvarchar,CONVERT(datetime,nRECDATE), 111) aS nRECDATE,COUNT(*) aS PN01 "
        STMT += " FROM P_05 "
        STMT += " WHERE CONVERT(nvarchar,CONVERT(datetime,nRECDATE), 111) >= '" & SDate & "'  "
        STMT += " AND CONVERT(nvarchar,CONVERT(datetime,nRECDATE), 111) <= '" & EDate & "' AND PENDFLAG='E' and nRECROOM= '" & DrDown_nRECROOM.Text & "'"
        STMT += " GROUP BY CONVERT(nvarchar,CONVERT(datetime,nRECDATE), 111)) A LEFT OUTER JOIN "
        STMT += " (SELECT CONVERT(nvarchar,CONVERT(datetime,P_05.nRECDATE), 111) aS nRECDATE,COUNT(*) aS PN02 "
        STMT += " FROM P_05, P_0501 "
        STMT += " WHERE P_05.EFORMSN = P_0501.EFORMSN "
        STMT += " AND CONVERT(nvarchar,CONVERT(datetime,P_05.nRECDATE), 111) >= '" & SDate & "' "
        STMT += " AND CONVERT(nvarchar,CONVERT(datetime,P_05.nRECDATE), 111) <= '" & EDate & "' "
        STMT += " AND PENDFLAG='E' AND P_0501.NKIND='民' and nRECROOM= '" & DrDown_nRECROOM.Text & "'"
        STMT += " GROUP BY CONVERT(nvarchar,CONVERT(datetime,P_05.nRECDATE), 111)) B ON A.nRECDATE=B.nRECDATE  LEFT OUTER JOIN "
        STMT += " (SELECT CONVERT(nvarchar,CONVERT(datetime,P_05.nRECDATE), 111) aS nRECDATE, "
        STMT += " COUNT(*) aS PN03 "
        STMT += " FROM P_05, P_0501 "
        STMT += " WHERE P_05.EFORMSN = P_0501.EFORMSN "
        STMT += " AND CONVERT(nvarchar,CONVERT(datetime,P_05.nRECDATE), 111) >= '" & SDate & "' "
        STMT += " AND CONVERT(nvarchar,CONVERT(datetime,P_05.nRECDATE), 111) <= '" & EDate & "'"
        STMT += " AND PENDFLAG='E' AND P_0501.NKIND='軍' and nRECROOM= '" & DrDown_nRECROOM.Text & "'"
        STMT += " GROUP BY CONVERT(nvarchar,CONVERT(datetime,P_05.nRECDATE), 111)) c ON A.nRECDATE=c.nRECDATE )) E ON  D.nRECDATE=E.nRECDATE "

        SQLALL = STMT

    End Function

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged

        '分頁
        Dim strOrd As String = ""

        strOrd = " ORDER BY D.nRECDATE "

        SqlDataSource2.SelectCommand = SQLALL(CDate(Sdate.Text).ToString("yyyy/MM/dd"), CDate(Edate.Text).ToString("yyyy/MM/dd")) & strOrd

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        '排序
        
        Dim strOrd As String = ""

        strOrd = " ORDER BY " & GridView1.SortExpression.ToString()

        SqlDataSource2.SelectCommand = SQLALL(CDate(Sdate.Text).ToString("yyyy/MM/dd"), CDate(Edate.Text).ToString("yyyy/MM/dd")) & strOrd

    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        If Not IsPostBack Then

            '登入馬上查詢
            ImgSearch_Click(Nothing, Nothing)

        End If

    End Sub


    Protected Sub ImagePrint_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImagePrint.Click
        Dim F_file As String
        Dim F_file2 As String
        Dim F_file_name As String
        Dim print As New C_Xprint

        If CDate(Sdate.Text).ToString("yyyy/MM") <> CDate(Edate.Text).ToString("yyyy/MM") Then
            do_sql.G_errmsg = "查詢日期需同月份"
            Exit Sub
        End If

        STMT = SQLALL(CDate(Sdate.Text).ToString("yyyy/MM/dd"), CDate(Edate.Text).ToString("yyyy/MM/dd")) & " ORDER BY D.nRECDATE "

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

        F_file_name = "rpt050060"
        F_file = MapPath("../../rpt/" + F_file_name + ".txt")
        If File.Exists(F_file) = False Then
            do_sql.G_errmsg = "檔案不存在"
            Exit Sub
        End If

        print_file = "../../drs/" + F_file_name + "-" + Now.ToString("mmssff") & ".drs"
        F_file2 = MapPath(print_file)

        print.C_Xprint(F_file.Split("\")(F_file.Split("\").Length - 1), print_file.Split("/")(print_file.Split("/").Length - 1))
        print.NewPage()

        'Call do_sql.inc_file(F_file, F_file2, F_file_name)
        Dim tmp_month As String = ""
        Dim nPage As Integer = 0
        Dim tot_PN01 As Integer = 0
        Dim tot_PN03 As Integer = 0
        Dim tot_PN02 As Integer = 0
        Dim tot_PN04 As Integer = 0
        Dim n_line As Integer = 0
        For Each dr In n_table.Rows
            If tmp_month <> CDate(Sdate.Text).ToString("yyyy/MM") Then
                tmp_month = CDate(Sdate.Text).ToString("yyyy/MM")
                nPage += 1
                'Call do_sql.print_sdata(F_file2, "/init " + F_file_name + ".txt 1", "/newpage null")
                prn_stmt = (CInt(CDate(Sdate.Text).ToString("yyyy")) - 1911).ToString & " 年 " & CDate(Sdate.Text).ToString("MM") & " 月份"
                prn_stmt += DrDown_nRECROOM.SelectedItem.Text + "電腦系統會客 (洽公) 統計表 "
                'SELECT D.nRECDATE,DWEEK,ISNULL(PN01,0) PN01,ISNULL(PN02,0) aS PN02,ISNULL(PN03,0) aS PN03
                'Call do_sql.print_block(F_file2, "第一會客室", 0, 0, prn_stmt)
                print.Add("第一會客室", prn_stmt, 0, 0)
                n_line = 0
            End If

            'Call do_sql.print_block(F_file2, "日期", 0, n_line * 7, CDate(dr("nRECDATE").ToString()).ToString("dd"))
            'Call do_sql.print_block(F_file2, "星期", 0, n_line * 7, dr("DWEEK").ToString)
            'Call do_sql.print_block(F_file2, "件數", 0, n_line * 7, dr("PN01").ToString)
            'Call do_sql.print_block(F_file2, "軍人", 0, n_line * 7, dr("PN03").ToString)
            'Call do_sql.print_block(F_file2, "民人", 0, n_line * 7, dr("PN02").ToString)
            'Call do_sql.print_block(F_file2, "合計", 0, n_line * 7, dr("PN04").ToString)

            print.Add("日期", CDate(dr("nRECDATE").ToString()).ToString("dd"), 0, n_line * 7)
            print.Add("星期", dr("DWEEK").ToString, 0, n_line * 7)
            print.Add("件數", dr("PN01").ToString, 0, n_line * 7)
            print.Add("軍人", dr("PN03").ToString, 0, n_line * 7)
            print.Add("民人", dr("PN02").ToString, 0, n_line * 7)
            print.Add("合計", dr("PN04").ToString, 0, n_line * 7)

            tot_PN01 += CInt(dr("PN01").ToString)
            tot_PN02 += CInt(dr("PN02").ToString)
            tot_PN03 += CInt(dr("PN03").ToString)
            tot_PN04 += CInt(dr("PN04").ToString)
            n_line = n_line + 1
        Next
        'Call do_sql.print_block(F_file2, "件數", 0, 31 * 7, tot_PN01.ToString)
        'Call do_sql.print_block(F_file2, "軍人", 0, 31 * 7, tot_PN03.ToString)
        'Call do_sql.print_block(F_file2, "民人", 0, 31 * 7, tot_PN02.ToString)
        'Call do_sql.print_block(F_file2, "合計", 0, 31 * 7, tot_PN04.ToString)

        print.Add("件數", tot_PN01.ToString, 0, 31 * 7)
        print.Add("軍人", tot_PN03.ToString, 0, 31 * 7)
        print.Add("民人", tot_PN02.ToString, 0, 31 * 7)
        print.Add("合計", tot_PN04.ToString, 0, 31 * 7)

        print.EndFile()

    End Sub

   
    Protected Sub Img_Export_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles Img_Export.Click
        Dim strsqlAll As String = SQLALL(CDate(Sdate.Text).ToString("yyyy/MM/dd"), CDate(Edate.Text).ToString("yyyy/MM/dd")) & " ORDER BY D.nRECDATE "

        '設定檔案路徑
        Dim path As String = "", filename As String = ""

        path = Server.MapPath("~/Drs/")
        filename = Date.Today.ToString("yyyyMMdd") & "_05006.csv"
        path = path & filename

        Dim colname As String = ""
        Dim data As String = ""
        Dim prn_stmt As String = ""

        Dim sw As New System.IO.StreamWriter(path, False, Encoding.GetEncoding("big5"))

        prn_stmt = (CInt(CDate(Sdate.Text).ToString("yyyy")) - 1911).ToString & " 年 " & CDate(Sdate.Text).ToString("MM") & " 月份"
        prn_stmt += DrDown_nRECROOM.SelectedItem.Text + "電腦系統會客 (洽公) 統計表 "

        sw.WriteLine(prn_stmt)

        colname = "日期,星期,件數,軍人,民人,合計,"

        colname = Left(colname, Len(colname) - 1)

        sw.WriteLine(colname)


        Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

        db.Open()
        Dim dt2 As DataTable = New DataTable("")
        Dim da2 As SqlDataAdapter = New SqlDataAdapter(strsqlAll, db)
        da2.Fill(dt2)
        db.Close()
        'SELECT D.nRECDATE,DWEEK,ISNULL(PN01,0) PN01,ISNULL(PN02,0) aS PN02,ISNULL(PN03,0) aS PN03
        For y As Integer = 0 To dt2.Rows.Count - 1

            data += dt2.Rows(y).Item("nRECDATE").ToString & "," & dt2.Rows(y).Item("DWEEK").ToString & ","
            data += dt2.Rows(y).Item("PN01").ToString & "," & dt2.Rows(y).Item("PN03").ToString & ","
            data += dt2.Rows(y).Item("PN02").ToString & "," & dt2.Rows(y).Item("PN04").ToString & ","
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
