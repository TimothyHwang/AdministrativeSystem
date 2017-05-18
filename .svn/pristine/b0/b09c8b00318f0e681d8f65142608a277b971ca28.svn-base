Imports System.Data.SqlClient
Imports System.Data
Imports System.IO
Partial Class M_Source_04_MOA04006
    Inherits System.Web.UI.Page

    Public do_sql As New C_SQLFUN
    Dim dr As System.Data.DataRow
    Dim n_table As New System.Data.DataTable
    Dim conn As New C_SQLFUN
    Public OrgChange As String      '判斷組織是否變更
    Dim user_id, org_uid As String
    Dim UnitFlag As String = ""
    Dim ParentFlag As String = ""
    Dim STMT As String = ""
    Public print_file As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        user_id = Session("user_id")
        org_uid = Session("ORG_UID")

        Dim connstr As String = conn.G_conn_string

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '判斷是否有下一級單位
        db.Open()
        Dim strPer As New SqlCommand("SELECT ORG_UID FROM ADMINGROUP WHERE PARENT_ORG_UID = '" & org_uid & "'", db)
        Dim RdPer = strPer.ExecuteReader()
        If RdPer.read() Then
            ParentFlag = "Y"
        End If
        db.Close()

        '判斷是否為房屋水電修繕人員
        db.Open()
        Dim strUnit As New SqlCommand("SELECT Role_Num FROM ROLEGROUPITEM WHERE Group_Uid = '5K6V1X2I65' AND employee_id = '" & user_id & "'", db)
        Dim RdUnit = strUnit.ExecuteReader()
        If RdUnit.read() Then
            UnitFlag = "Y"
        End If
        db.Close()

        If IsPostBack = False Then

            '找出同級單位以下全部單位
            Dim Org_Down As New C_Public

            '找出登入者的一級單位
            Dim strTopOrg As String = ""
            strTopOrg = Org_Down.getUporg(org_uid, 1)

            '判斷登入者權限
            If Session("Role") = "1" Or UnitFlag = "Y" Then
                SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
            ElseIf Session("Role") = "2" Then
                SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(strTopOrg) & ") ORDER BY ORG_NAME"
            Else
                SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ") ORDER BY ORG_NAME"
            End If

            '設定Default日期
            Dim Sdate As TextBox = Me.FindControl("Sdate")
            Dim Edate As TextBox = Me.FindControl("Edate")

            Dim dt As Date = Now()
            If (Sdate.Text = "") Then
                Sdate.Text = dt.AddDays(-7).Date
            End If

            If (Edate.Text = "") Then
                Edate.Text = dt.Date
            End If

        End If

    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        '查詢
        Dim strOrd As String = ""

        strOrd = " ORDER BY PAUNIT"

        SqlDataSource2.SelectCommand = SQLALL(OrgSel.SelectedValue, Sdate.Text, Edate.Text) & strOrd

        OrgChange = ""
    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        If Not IsPostBack Then

            Dim tLItm As New ListItem("請選擇", "")

            '系統管理員組織加請選擇

            OrgSel.Items.Insert(0, tLItm)
            If OrgSel.Items.Count > 1 Then
                OrgSel.Items(1).Selected = False
            End If

            '登入馬上查詢
            ImgSearch_Click(Nothing, Nothing)

        End If

    End Sub

    Public Function SQLALL(ByVal OrgSel, ByVal SDate, ByVal EDate)

        '整合SQL搜尋字串
        Dim strsql As String = ""
        Dim strsel As String = ""
        Dim strSqlAll As String = ""
        Dim strSqlFirst As String = ""
        Dim strSqlLast As String = ""

        '找出同級單位以下全部單位
        Dim Org_Down As New C_Public

        '組織搜尋
        If OrgSel <> "" Then
            strsel = " AND PAUNIT='" & OrgSel & "'"
        End If

        strsql = "select B.PAUNIT,isnull(B.p01,0) As p01,isnull(C.p02,0) As p02 "
        strsql += " from (select PAUNIT,count(*) As p01 from P_04 "
        strsql += " where pendflag='E' "
        strsql += strsel
        strsql += " and nAPPTIME >= '" & SDate & "' and  nAPPTIME <= '" & EDate & "'"
        strsql += " group by PAUNIT ) B  left outer join  "
        strsql += " (select P_04.PAUNIT,count(*) As p02 from P_04,P_0401   "
        strsql += " where P_04.pendflag='E' "

        If OrgSel <> "" Then
            strsql += " AND P_04.PAUNIT='" & OrgSel & "'"
        End If

        strsql += " and P_04.nAPPTIME >= '" & SDate & "' and  P_04.nAPPTIME <= '" & EDate & "' "
        strsql += " and P_0401.pendflag='E'  "
        strsql += " and P_04.eformsn = P_0401.neformsn group by P_04.PAUNIT) C on B.PAUNIT=C.PAUNIT "

        SQLALL = strsql

    End Function

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged

        '分頁
        Dim strOrd As String = ""

        strOrd = " ORDER BY PAUNIT"

        SqlDataSource2.SelectCommand = SQLALL(OrgSel.SelectedValue, Sdate.Text, Edate.Text) & strOrd

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        '排序
        Dim strOrd As String = ""

        strOrd = " ORDER BY " & GridView1.SortExpression.ToString()

        SqlDataSource2.SelectCommand = SQLALL(OrgSel.SelectedValue, Sdate.Text, Edate.Text) & strOrd

    End Sub

    Protected Sub ImagePrint_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImagePrint.Click
        Dim F_file As String
        Dim F_file2 As String
        Dim F_file_name As String
        Dim strOrd As String = ""
        Dim print As New C_Xprint

        STMT = SQLALL(OrgSel.SelectedValue, Sdate.Text, Edate.Text) & " ORDER BY PAUNIT"

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

        F_file_name = "rpt040060"
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
        Dim nPage As Integer = 0
        Dim n_line As Integer = 35
        Dim tmp_add As String = ""
        Dim tmp_name As String = ""
        Dim tmp_month As String = ""
        For Each dr In n_table.Rows
            If n_line >= 34 Then
                nPage += 1
                'Call do_sql.print_sdata(F_file2, "/init " + F_file_name + ".txt 1", "/newpage null")                
                prn_stmt = CDate(Sdate.Text).ToString("yyyy/MM/dd") & "~" & CDate(Edate.Text).ToString("yyyy/MM/dd")
                'Call do_sql.print_block(F_file2, "日期", 0, 0, prn_stmt)
                print.Add("日期", prn_stmt, 0, 0)
                n_line = 0
            End If
            'SELECT employee_id,emp_chinese_name,CONVERT(nvarchar, CONVERT(datetime,'" & SDate & "'), 111) AS ShowDay,SIGNDATE_d,In_Time_nvc,Out_Time_nvc,REASON_nvc
            'Call do_sql.print_block(F_file2, "部門名稱", 0, n_line * 7, dr("PAUNIT").ToString)
            'Call do_sql.print_block(F_file2, "核派", 0, n_line * 7, dr("p01").ToString())
            print.Add("部門名稱", dr("PAUNIT").ToString, 0, n_line * 7)
            print.Add("核派", dr("p01").ToString(), 0, n_line * 7)

            
            'Call do_sql.print_block(F_file2, "完成", 0, n_line * 7, dr("p02").ToString())
            print.Add("完成", dr("p02").ToString(), 0, n_line * 7)

            n_line = n_line + 1
        Next
        print.EndFile()

    End Sub
    

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        'Sdate.Text
        Server.Transfer("MOA04007.aspx?x=MOA04007&PAUNIT=" & GridView1.SelectedValue & "&Sdate=" & Sdate.Text & "&Edate=" & Edate.Text)
        'Server.Transfer("../04/MOA04007.aspx?x=MOA04007&EFORMSN=" & GridView1.SelectedValue)
        'do_sql.G_errmsg = GridView1.SelectedValue
    End Sub
    Protected Sub Img_Export_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles Img_Export.Click

        Dim strsqlAll As String = SQLALL(OrgSel.SelectedValue, Sdate.Text, Edate.Text) & " ORDER BY PAUNIT"

        '設定檔案路徑
        Dim path As String = "", filename As String = ""

        path = Server.MapPath("~/Drs/")
        filename = Date.Today.ToString("yyyyMMdd") & "_040060.csv"
        path = path & filename

        Dim colname As String = ""
        Dim data As String = ""
        Dim prn_stmt As String = ""

        Dim sw As New System.IO.StreamWriter(path, False, Encoding.GetEncoding("big5"))

        prn_stmt = (CInt(CDate(Sdate.Text).ToString("yyyy")) - 1911).ToString & " 年 " & CDate(Sdate.Text).ToString("MM") & " 月份"
        prn_stmt += "房屋水電修繕統計表 "

        sw.WriteLine(prn_stmt)

        colname = "部門名稱,核派,完成,"

        colname = Left(colname, Len(colname) - 1)

        sw.WriteLine(colname)


        Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

        db.Open()
        Dim dt2 As DataTable = New DataTable("")
        Dim da2 As SqlDataAdapter = New SqlDataAdapter(strsqlAll, db)
        da2.Fill(dt2)
        db.Close()
        Dim tmp_add As String = ""

        For y As Integer = 0 To dt2.Rows.Count - 1

            data += dt2.Rows(y).Item("PAUNIT").ToString & "," & dt2.Rows(y).Item("p01").ToString & ","
            data += dt2.Rows(y).Item("p02").ToString & ","
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
