Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Web.Hosting

Partial Class Source_03_MOA03002
    Inherits System.Web.UI.Page
    Public do_sql As New C_SQLFUN
    Dim dr As System.Data.DataRow
    Dim n_table As New System.Data.DataTable
    Dim n_table2 As New System.Data.DataTable
    Dim p As Integer = 0
    Public display_flag As Boolean = False
    Public x_point As Integer = 489
    Public y_point As Integer = 46
    Public stmt As String
    Dim select_flag As Boolean = False
    Public print_file As String = ""
    Dim connstr As String
    Dim strRowIndex As String = ""
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

            ' Call add_stmt()
            do_sql.G_user_id = Session("user_id")
            'do_sql.G_user_id = "oasadmin"

            strRowIndex = 0

            If IsPostBack Then
                Exit Sub
            End If
            Txt_nARRDATE.Text = Now.ToString("yyyy/MM/dd")

            stmt = "select * from p_0302 order by pck_name"
            If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
                Exit Sub
            End If
            n_table = do_sql.G_table
            DrDown_nstyle.Items.Clear()
            p = 0
            DrDown_nstyle.Items.Add("")
            DrDown_nstyle.Items(p).Value = ""
            p += 1
            For Each dr In n_table.Rows
                DrDown_nstyle.Items.Add(Trim(dr("pck_name").ToString.Trim))
                DrDown_nstyle.Items(p).Value = Trim(dr("pck_name").ToString.Trim)
                p += 1
            Next

        End If

    End Sub
    Sub add_stmt()
        If Txt_nARRDATE.Text = "" Then
            Txt_nARRDATE.Text = Now.ToString("yyyy/MM/dd")
        End If
        stmt = "SELECT P_03.EFORMSN, P_0301.DriveName, P_0301.DriveTel, P_0301.CarNumber, P_03.nARRIVEPLACE + '--' + P_03.nARRIVETO AS per,"
        stmt += "P_03.nPHONE, P_03.nENDPOINT, P_03.nREASON, P_0301.LeaveTime,SUBSTRING(P_0301.LeaveTime,1,2) LeaveHour,SUBSTRING(P_0301.LeaveTime,3,2) LeaveMin ,"
        stmt += "P_0301.ComeTime,SUBSTRING(P_0301.ComeTime,1,2) ComeHour,SUBSTRING(P_0301.ComeTime,3,2) ComeMin, P_0301.LeaveMilage, "
        stmt += "P_0301.ComeMilage, P_0301.RealMilage, P_03.nSTYLE,P_0301.JoinCar "
        stmt += " FROM P_03 INNER JOIN P_0301 ON P_03.EFORMSN = P_0301.EFORMSN"
        stmt += " where nARRDATE ='" + Txt_nARRDATE.Text + "'"
        If TXT_nREASON.Text <> "" Then
            stmt += " and P_03.nREASON like '" + TXT_nREASON.Text.Trim + "%'"
        End If
        If Txt_DriveName.Text <> "" Then
            stmt += " and P_0301.DriveName like '" + Txt_DriveName.Text.Trim + "%'"
        End If
        If DrDown_nstyle.SelectedItem.Text <> "" Then
            stmt += " and P_03.nSTYLE like '" + DrDown_nstyle.SelectedItem.Text + "%'"
        End If
        If Txt_CarNumber.Text <> "" Then
            stmt += " and P_0301.CarNumber like '" + Txt_CarNumber.Text.Trim + "%'"
        End If
        If Txt_nARRIVEPLACE.Text <> "" Then
            stmt += " and nARRIVEPLACE like '" + Txt_nARRIVEPLACE.Text.Trim + "%'"
        End If
        If Txt_nARRIVETO.Text <> "" Then
            stmt += " and nARRIVETO like '" + Txt_nARRIVETO.Text.Trim + "%'"
        End If
        SqlDataSource1.SelectCommand = stmt
        select_flag = True
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'GridView1.Rows(0).Cells(7).Text = Now.ToString("yyyy/MM/dd HH:mm:ss")
        'CType(GridView1.Rows(0).Cells(7).FindControl("TextBox1"), Label).Text = Now
        select_flag = True
        CType(GridView1.Rows(GridView1.EditIndex).Cells(7).FindControl("Label3"), Label).Text = Now.ToString("yyyy/MM/dd HH:mm:ss")

    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        select_flag = True
        CType(GridView1.Rows(GridView1.EditIndex).Cells(8).FindControl("Label5"), Label).Text = Now.ToString("yyyy/MM/dd HH:mm:ss")
    End Sub
   
    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged
        select_flag = True
        Txt_nARRDATE.Text = Calendar1.SelectedDate.ToString("yyyy/MM/dd")
        Div_grid.Visible = False
    End Sub

    Protected Sub Calendar1_VisibleMonthChanged(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.MonthChangedEventArgs) Handles Calendar1.VisibleMonthChanged
        select_flag = True
        display_flag = True
    End Sub

    Protected Sub DropDownList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        select_flag = True
        CType(GridView1.Rows(GridView1.EditIndex).Cells(1).FindControl("DropDownList6"), DropDownList).SelectedIndex = CType(GridView1.Rows(GridView1.EditIndex).Cells(1).FindControl("DropDownList1"), DropDownList).SelectedIndex
    End Sub

    Protected Sub DropDownList2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call add_col7()
    End Sub

    Protected Sub DropDownList3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call add_col7()
    End Sub
    Protected Sub DropDownList4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call add_col8()
    End Sub

    Protected Sub DropDownList5_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call add_col8()
    End Sub
    Sub add_col7()
        select_flag = True
        Dim mour_min As String
        mour_min = CType(GridView1.Rows(GridView1.EditIndex).Cells(7).FindControl("DropDownList2"), DropDownList).Text
        mour_min += CType(GridView1.Rows(GridView1.EditIndex).Cells(7).FindControl("DropDownList3"), DropDownList).Text
        CType(GridView1.Rows(GridView1.EditIndex).Cells(7).FindControl("Label3"), Label).Text = mour_min
    End Sub
    Sub add_col8()
        select_flag = True
        Dim mour_min As String
        mour_min = CType(GridView1.Rows(GridView1.EditIndex).Cells(8).FindControl("DropDownList4"), DropDownList).Text
        mour_min += CType(GridView1.Rows(GridView1.EditIndex).Cells(8).FindControl("DropDownList5"), DropDownList).Text
        CType(GridView1.Rows(GridView1.EditIndex).Cells(8).FindControl("Label5"), Label).Text = mour_min
    End Sub
    Protected Sub Page_LoadComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LoadComplete
        If select_flag = False Then
            Call add_stmt()
        End If
    End Sub

    Protected Sub ImagePrint_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImagePrint.Click
        Dim F_file As String
        Dim F_file2 As String
        Dim F_file_name As String
        Dim str_line As String = ""
        Dim tot_row As Integer = 0
        Dim second_file As Boolean = False
        Dim hi As Integer

        'stmt = "SELECT P_03.EFORMSN, P_0301.DriveName, P_0301.DriveTel, P_0301.CarNumber, P_03.nARRIVEPLACE + '--' + P_03.nARRIVETO AS per,"
        'stmt += "P_03.nPHONE, P_03.nENDPOINT, P_03.nREASON, P_0301.LeaveTime,SUBSTRING(P_0301.LeaveTime,1,2) LeaveHour,SUBSTRING(P_0301.LeaveTime,3,2) LeaveMin ,"
        'stmt += "P_0301.ComeTime,SUBSTRING(P_0301.ComeTime,1,2) ComeHour,SUBSTRING(P_0301.ComeTime,3,2) ComeMin, P_0301.LeaveMilage, "
        'stmt += "P_0301.ComeMilage, P_0301.RealMilage, P_03.nSTYLE "
        'stmt += " FROM P_03 INNER JOIN P_0301 ON P_03.EFORMSN = P_0301.EFORMSN "

        ''stmt += " union all SELECT P_03.EFORMSN, P_0301.DriveName, P_0301.DriveTel, P_0301.CarNumber, P_03.nARRIVEPLACE + '--' + P_03.nARRIVETO AS per,"
        ''stmt += "P_03.nPHONE, P_03.nENDPOINT, P_03.nREASON, P_0301.LeaveTime,SUBSTRING(P_0301.LeaveTime,1,2) LeaveHour,SUBSTRING(P_0301.LeaveTime,3,2) LeaveMin ,"
        ''stmt += "P_0301.ComeTime,SUBSTRING(P_0301.ComeTime,1,2) ComeHour,SUBSTRING(P_0301.ComeTime,3,2) ComeMin, P_0301.LeaveMilage, "
        ''stmt += "P_0301.ComeMilage, P_0301.RealMilage, P_03.nSTYLE "
        ''stmt += " FROM P_03 INNER JOIN P_0301 ON P_03.EFORMSN = P_0301.EFORMSN "
        ''stmt += " union all SELECT P_03.EFORMSN, P_0301.DriveName, P_0301.DriveTel, P_0301.CarNumber, P_03.nARRIVEPLACE + '--' + P_03.nARRIVETO AS per,"
        ''stmt += "P_03.nPHONE, P_03.nENDPOINT, P_03.nREASON, P_0301.LeaveTime,SUBSTRING(P_0301.LeaveTime,1,2) LeaveHour,SUBSTRING(P_0301.LeaveTime,3,2) LeaveMin ,"
        ''stmt += "P_0301.ComeTime,SUBSTRING(P_0301.ComeTime,1,2) ComeHour,SUBSTRING(P_0301.ComeTime,3,2) ComeMin, P_0301.LeaveMilage, "
        ''stmt += "P_0301.ComeMilage, P_0301.RealMilage, '大貨卡' As nSTYLE "
        ''stmt += " FROM P_03 INNER JOIN P_0301 ON P_03.EFORMSN = P_0301.EFORMSN "
        ''stmt += " union all SELECT P_03.EFORMSN, P_0301.DriveName, P_0301.DriveTel, P_0301.CarNumber, P_03.nARRIVEPLACE + '--' + P_03.nARRIVETO AS per,"
        ''stmt += "P_03.nPHONE, P_03.nENDPOINT, P_03.nREASON, P_0301.LeaveTime,SUBSTRING(P_0301.LeaveTime,1,2) LeaveHour,SUBSTRING(P_0301.LeaveTime,3,2) LeaveMin ,"
        ''stmt += "P_0301.ComeTime,SUBSTRING(P_0301.ComeTime,1,2) ComeHour,SUBSTRING(P_0301.ComeTime,3,2) ComeMin, P_0301.LeaveMilage, "
        ''stmt += "P_0301.ComeMilage, P_0301.RealMilage, '大貨卡' As nSTYLE "
        ''stmt += " FROM P_03 INNER JOIN P_0301 ON P_03.EFORMSN = P_0301.EFORMSN "
        ''stmt += " union all SELECT P_03.EFORMSN, P_0301.DriveName, P_0301.DriveTel, P_0301.CarNumber, P_03.nARRIVEPLACE + '--' + P_03.nARRIVETO AS per,"
        ''stmt += "P_03.nPHONE, P_03.nENDPOINT, P_03.nREASON, P_0301.LeaveTime,SUBSTRING(P_0301.LeaveTime,1,2) LeaveHour,SUBSTRING(P_0301.LeaveTime,3,2) LeaveMin ,"
        ''stmt += "P_0301.ComeTime,SUBSTRING(P_0301.ComeTime,1,2) ComeHour,SUBSTRING(P_0301.ComeTime,3,2) ComeMin, P_0301.LeaveMilage, "
        ''stmt += "P_0301.ComeMilage, P_0301.RealMilage, '中巴士' As nSTYLE "
        ''stmt += " FROM P_03 INNER JOIN P_0301 ON P_03.EFORMSN = P_0301.EFORMSN " '15

        ''stmt += " union all SELECT P_03.EFORMSN, P_0301.DriveName, P_0301.DriveTel, P_0301.CarNumber, P_03.nARRIVEPLACE + '--' + P_03.nARRIVETO AS per,"
        ''stmt += "P_03.nPHONE, P_03.nENDPOINT, P_03.nREASON, P_0301.LeaveTime,SUBSTRING(P_0301.LeaveTime,1,2) LeaveHour,SUBSTRING(P_0301.LeaveTime,3,2) LeaveMin ,"
        ''stmt += "P_0301.ComeTime,SUBSTRING(P_0301.ComeTime,1,2) ComeHour,SUBSTRING(P_0301.ComeTime,3,2) ComeMin, P_0301.LeaveMilage, "
        ''stmt += "P_0301.ComeMilage, P_0301.RealMilage, '大福特' As nSTYLE "
        ''stmt += " FROM P_03 INNER JOIN P_0301 ON P_03.EFORMSN = P_0301.EFORMSN "
        ''stmt += " union all SELECT P_03.EFORMSN, P_0301.DriveName, P_0301.DriveTel, P_0301.CarNumber, P_03.nARRIVEPLACE + '--' + P_03.nARRIVETO AS per,"
        ''stmt += "P_03.nPHONE, P_03.nENDPOINT, P_03.nREASON, P_0301.LeaveTime,SUBSTRING(P_0301.LeaveTime,1,2) LeaveHour,SUBSTRING(P_0301.LeaveTime,3,2) LeaveMin ,"
        ''stmt += "P_0301.ComeTime,SUBSTRING(P_0301.ComeTime,1,2) ComeHour,SUBSTRING(P_0301.ComeTime,3,2) ComeMin, P_0301.LeaveMilage, "
        ''stmt += "P_0301.ComeMilage, P_0301.RealMilage, '八人座' As nSTYLE "
        ''stmt += " FROM P_03 INNER JOIN P_0301 ON P_03.EFORMSN = P_0301.EFORMSN "
        ''stmt += " union all SELECT P_03.EFORMSN, P_0301.DriveName, P_0301.DriveTel, P_0301.CarNumber, P_03.nARRIVEPLACE + '--' + P_03.nARRIVETO AS per,"
        ''stmt += "P_03.nPHONE, P_03.nENDPOINT, P_03.nREASON, P_0301.LeaveTime,SUBSTRING(P_0301.LeaveTime,1,2) LeaveHour,SUBSTRING(P_0301.LeaveTime,3,2) LeaveMin ,"
        ''stmt += "P_0301.ComeTime,SUBSTRING(P_0301.ComeTime,1,2) ComeHour,SUBSTRING(P_0301.ComeTime,3,2) ComeMin, P_0301.LeaveMilage, "
        ''stmt += "P_0301.ComeMilage, P_0301.RealMilage, '小貨卡' As nSTYLE "
        ''stmt += " FROM P_03 INNER JOIN P_0301 ON P_03.EFORMSN = P_0301.EFORMSN " '24

        ''stmt += " union all SELECT P_03.EFORMSN, P_0301.DriveName, P_0301.DriveTel, P_0301.CarNumber, P_03.nARRIVEPLACE + '--' + P_03.nARRIVETO AS per,"
        ''stmt += "P_03.nPHONE, P_03.nENDPOINT, P_03.nREASON, P_0301.LeaveTime,SUBSTRING(P_0301.LeaveTime,1,2) LeaveHour,SUBSTRING(P_0301.LeaveTime,3,2) LeaveMin ,"
        ''stmt += "P_0301.ComeTime,SUBSTRING(P_0301.ComeTime,1,2) ComeHour,SUBSTRING(P_0301.ComeTime,3,2) ComeMin, P_0301.LeaveMilage, "
        ''stmt += "P_0301.ComeMilage, P_0301.RealMilage, '1.6' As nSTYLE "
        ''stmt += " FROM P_03 INNER JOIN P_0301 ON P_03.EFORMSN = P_0301.EFORMSN "
        ''stmt += " union all SELECT P_03.EFORMSN, P_0301.DriveName, P_0301.DriveTel, P_0301.CarNumber, P_03.nARRIVEPLACE + '--' + P_03.nARRIVETO AS per,"
        ''stmt += "P_03.nPHONE, P_03.nENDPOINT, P_03.nREASON, P_0301.LeaveTime,SUBSTRING(P_0301.LeaveTime,1,2) LeaveHour,SUBSTRING(P_0301.LeaveTime,3,2) LeaveMin ,"
        ''stmt += "P_0301.ComeTime,SUBSTRING(P_0301.ComeTime,1,2) ComeHour,SUBSTRING(P_0301.ComeTime,3,2) ComeMin, P_0301.LeaveMilage, "
        ''stmt += "P_0301.ComeMilage, P_0301.RealMilage, '2.0以上' As nSTYLE "
        ''stmt += " FROM P_03 INNER JOIN P_0301 ON P_03.EFORMSN = P_0301.EFORMSN "
        ''stmt += " union all SELECT P_03.EFORMSN, P_0301.DriveName, P_0301.DriveTel, P_0301.CarNumber, P_03.nARRIVEPLACE + '--' + P_03.nARRIVETO AS per,"
        ''stmt += "P_03.nPHONE, P_03.nENDPOINT, P_03.nREASON, P_0301.LeaveTime,SUBSTRING(P_0301.LeaveTime,1,2) LeaveHour,SUBSTRING(P_0301.LeaveTime,3,2) LeaveMin ,"
        ''stmt += "P_0301.ComeTime,SUBSTRING(P_0301.ComeTime,1,2) ComeHour,SUBSTRING(P_0301.ComeTime,3,2) ComeMin, P_0301.LeaveMilage, "
        ''stmt += "P_0301.ComeMilage, P_0301.RealMilage, '1.8' As nSTYLE "
        ''stmt += " FROM P_03 INNER JOIN P_0301 ON P_03.EFORMSN = P_0301.EFORMSN "
        ''stmt += " union all SELECT P_03.EFORMSN, P_0301.DriveName, P_0301.DriveTel, P_0301.CarNumber, P_03.nARRIVEPLACE + '--' + P_03.nARRIVETO AS per,"
        ''stmt += "P_03.nPHONE, P_03.nENDPOINT, P_03.nREASON, P_0301.LeaveTime,SUBSTRING(P_0301.LeaveTime,1,2) LeaveHour,SUBSTRING(P_0301.LeaveTime,3,2) LeaveMin ,"
        ''stmt += "P_0301.ComeTime,SUBSTRING(P_0301.ComeTime,1,2) ComeHour,SUBSTRING(P_0301.ComeTime,3,2) ComeMin, P_0301.LeaveMilage, "
        ''stmt += "P_0301.ComeMilage, P_0301.RealMilage, '租賃大巴士' As nSTYLE "
        ''stmt += " FROM P_03 INNER JOIN P_0301 ON P_03.EFORMSN = P_0301.EFORMSN "
        ''stmt += " union all SELECT P_03.EFORMSN, P_0301.DriveName, P_0301.DriveTel, P_0301.CarNumber, P_03.nARRIVEPLACE + '--' + P_03.nARRIVETO AS per,"
        ''stmt += "P_03.nPHONE, P_03.nENDPOINT, P_03.nREASON, P_0301.LeaveTime,SUBSTRING(P_0301.LeaveTime,1,2) LeaveHour,SUBSTRING(P_0301.LeaveTime,3,2) LeaveMin ,"
        ''stmt += "P_0301.ComeTime,SUBSTRING(P_0301.ComeTime,1,2) ComeHour,SUBSTRING(P_0301.ComeTime,3,2) ComeMin, P_0301.LeaveMilage, "
        ''stmt += "P_0301.ComeMilage, P_0301.RealMilage, '租賃中巴士' As nSTYLE "
        ''stmt += " FROM P_03 INNER JOIN P_0301 ON P_03.EFORMSN = P_0301.EFORMSN " '39

        ''stmt += " union all SELECT P_03.EFORMSN, P_0301.DriveName, P_0301.DriveTel, P_0301.CarNumber, P_03.nARRIVEPLACE + '--' + P_03.nARRIVETO AS per,"
        ''stmt += "P_03.nPHONE, P_03.nENDPOINT, P_03.nREASON, P_0301.LeaveTime,SUBSTRING(P_0301.LeaveTime,1,2) LeaveHour,SUBSTRING(P_0301.LeaveTime,3,2) LeaveMin ,"
        ''stmt += "P_0301.ComeTime,SUBSTRING(P_0301.ComeTime,1,2) ComeHour,SUBSTRING(P_0301.ComeTime,3,2) ComeMin, P_0301.LeaveMilage, "
        ''stmt += "P_0301.ComeMilage, P_0301.RealMilage, '租賃八人座' As nSTYLE "
        ''stmt += " FROM P_03 INNER JOIN P_0301 ON P_03.EFORMSN = P_0301.EFORMSN "
        ''stmt += " union all SELECT P_03.EFORMSN, P_0301.DriveName, P_0301.DriveTel, P_0301.CarNumber, P_03.nARRIVEPLACE + '--' + P_03.nARRIVETO AS per,"
        ''stmt += "P_03.nPHONE, P_03.nENDPOINT, P_03.nREASON, P_0301.LeaveTime,SUBSTRING(P_0301.LeaveTime,1,2) LeaveHour,SUBSTRING(P_0301.LeaveTime,3,2) LeaveMin ,"
        ''stmt += "P_0301.ComeTime,SUBSTRING(P_0301.ComeTime,1,2) ComeHour,SUBSTRING(P_0301.ComeTime,3,2) ComeMin, P_0301.LeaveMilage, "
        ''stmt += "P_0301.ComeMilage, P_0301.RealMilage, '租賃八人座' As nSTYLE "
        ''stmt += " FROM P_03 INNER JOIN P_0301 ON P_03.EFORMSN = P_0301.EFORMSN " '45

        'stmt += " where nARRDATE ='" + Txt_nARRDATE.Text + "'"
        'If TXT_nREASON.Text <> "" Then
        '    stmt += " and P_03.nREASON like '" + TXT_nREASON.Text + "%'"
        'End If
        'If Txt_DriveName.Text <> "" Then
        '    stmt += " and P_0301.DriveName like '" + Txt_DriveName.Text + "%'"
        'End If
        'If DrDown_nstyle.SelectedItem.Text <> "" Then
        '    stmt += " and P_03.nSTYLE like '" + DrDown_nstyle.SelectedItem.Text + "%'"
        'End If
        'If Txt_CarNumber.Text <> "" Then
        '    stmt += " and P_0301.CarNumber like '" + Txt_CarNumber.Text + "%'"
        'End If
        'If Txt_nARRIVEPLACE.Text <> "" Then
        '    stmt += " and nARRIVEPLACE like '" + Txt_nARRIVEPLACE.Text + "%'"
        'End If
        'If Txt_nARRIVETO.Text <> "" Then
        '    stmt += " and nARRIVETO like '" + Txt_nARRIVETO.Text + "%'"
        'End If
        stmt = SQLALL()
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Sub
        End If
        If do_sql.G_table.Rows.Count > 0 Then
            n_table = do_sql.G_table
            tot_row = n_table.Rows.Count
        Else
            do_sql.G_errmsg = "查無可印資料"
            Exit Sub
        End If


        If tot_row > 15 Then
            F_file_name = "rpt030020"
        Else
            F_file_name = "rpt030021"
        End If
        F_file2 = ""
        print_file = ""

        Dim line_no As Integer = 0
        Dim page_no As Integer = 0
        Dim block_name As String = ""
        Dim T_date As String = CStr(CInt(Now.ToString("yyyy")) - 1911) + "年"
        Dim user_name As String = ""
        T_date += Now.ToString("MM") + "月" + Now.ToString("dd") + "日"
        print_file = ""
        If do_sql.select_urname(do_sql.G_user_id) = False Then
            Exit Sub
        End If
        If do_sql.G_usr_table.Rows.Count > 0 Then
            user_name = do_sql.G_usr_table.Rows(0).Item("emp_chinese_name").ToString.Trim
        End If
        Dim y As Integer = 0
        Dim car_count(13) As Integer
        For hi = 0 To 13
            car_count(hi) = 0
        Next

        For hi = 0 To n_table.Rows.Count - 1
            If line_no = 0 Then
                page_no += 1
                If tot_row - hi > 15 Then
                    F_file_name = "rpt030020"
                Else
                    F_file_name = "rpt030021"
                    second_file = True
                End If
                F_file = MapPath("../../rpt/" + F_file_name + ".txt")
                If File.Exists(F_file) = False Then
                    do_sql.G_errmsg = "檔案不存在"
                    Exit Sub
                End If
                If print_file = "" Then
                    print_file = "../../drs/" + F_file_name + "-" + Now.ToString("mmssff") & ".drs"
                End If
                F_file2 = MapPath(print_file)
                '---------------------讀檔   輸出檔
                If page_no = 1 Then
                    Call do_sql.inc_file(F_file, F_file2, F_file_name)
                Else
                    Call do_sql.inc_file(F_file, F_file2, F_file_name, True)
                End If
                Call do_sql.print_sdata(F_file2, "/init " + F_file_name + ".txt 1", "/newpage null")

                Call do_sql.print_block(F_file2, "調度官", 0, 0, user_name)
                Call do_sql.print_block(F_file2, "日期", 0, 0, T_date)
                Call do_sql.print_block(F_file2, "頁次", 0, 0, page_no.ToString)
                y = 0
            End If

            line_no += 1
            block_name = "姓名"
            Call do_sql.print_block(F_file2, block_name, 0, y, n_table.Rows(hi).Item("DriveName").ToString())
            block_name = "車輛號碼"
            Call do_sql.print_block(F_file2, block_name, 0, y, n_table.Rows(hi).Item("CarNumber").ToString())
            block_name = "報到處"
            Call do_sql.print_block(F_file2, block_name, 0, y, n_table.Rows(hi).Item("per").ToString())
            block_name = "電話"
            Call do_sql.print_block(F_file2, block_name, 0, y, n_table.Rows(hi).Item("nPHONE").ToString())
            block_name = "地址"
            Call do_sql.print_block(F_file2, block_name, 0, y, n_table.Rows(hi).Item("nENDPOINT").ToString())
            block_name = "目的"
            Call do_sql.print_block(F_file2, block_name, 0, y, n_table.Rows(hi).Item("nREASON").ToString())
            block_name = "出場時間"
            Call do_sql.print_block(F_file2, block_name, 0, y, n_table.Rows(hi).Item("LeaveTime").ToString())
            block_name = "回場時間"
            Call do_sql.print_block(F_file2, block_name, 0, y, n_table.Rows(hi).Item("ComeTime").ToString())
            block_name = "出場里程"
            Call do_sql.print_block(F_file2, block_name, 0, y, n_table.Rows(hi).Item("LeaveMilage").ToString())
            block_name = "回場里程"
            Call do_sql.print_block(F_file2, block_name, 0, y, n_table.Rows(hi).Item("ComeMilage").ToString())
            block_name = "實際里程"
            Call do_sql.print_block(F_file2, block_name, 0, y, n_table.Rows(hi).Item("RealMilage").ToString())
            block_name = "車型"
            Call do_sql.print_block(F_file2, block_name, 0, y, n_table.Rows(hi).Item("nSTYLE").ToString())

            Select Case n_table.Rows(hi).Item("nSTYLE").ToString()
                Case "大巴士"
                    car_count(1) += 1
                Case "大貨卡"
                    car_count(2) += 1
                Case "中巴士"
                    car_count(3) += 1
                Case "大福特"
                    car_count(4) += 1
                Case "八人座"
                    car_count(5) += 1
                Case "小貨卡"
                    car_count(6) += 1
                Case "1.6轎車"
                    car_count(7) += 1
                Case "1.8轎車"
                    car_count(8) += 1
                Case "2.0轎車"
                    car_count(9) += 1
                Case "租賃大巴士"
                    car_count(10) += 1
                Case "租賃中巴士"
                    car_count(11) += 1
                Case "租賃八人座"
                    car_count(12) += 1
                Case Else
                    car_count(0) += 1
            End Select

            y += 9
            If line_no >= 26 Then
                line_no = 0
            End If
        Next
        If second_file = False Then
            page_no += 1
            F_file_name = "rpt030021"
            F_file = MapPath("../../rpt/" + F_file_name + ".txt")
            If File.Exists(F_file) = False Then
                do_sql.G_errmsg = "檔案不存在"
                Exit Sub
            End If
            If print_file = "" Then
                print_file = "../../drs/" + F_file_name + "-" + Now.ToString("mmssff") & ".drs"
            End If
            F_file2 = MapPath(print_file)
            '---------------------讀檔   輸出檔
            Call do_sql.inc_file(F_file, F_file2, F_file_name, True)
            Call do_sql.print_sdata(F_file2, "/init " + F_file_name + ".txt 1", "/newpage null")

            Call do_sql.print_block(F_file2, "調度官", 0, 0, user_name)
            Call do_sql.print_block(F_file2, "日期", 0, 0, T_date)
            Call do_sql.print_block(F_file2, "頁次", 0, 0, page_no.ToString)
        End If
        Dim tmp_print As String = ""
        block_name = "車型差勤人員"
        tmp_print = "民國" + T_date + "各車型差勤人員及車次統計"
        Call do_sql.print_block(F_file2, block_name, 0, 0, tmp_print)

        block_name = "大巴士"
        tmp_print = car_count(1).ToString() + "車次"
        Call do_sql.print_block(F_file2, block_name, 0, 0, tmp_print)
        block_name = "大貨卡"
        tmp_print = car_count(2).ToString() + "車次"
        Call do_sql.print_block(F_file2, block_name, 0, 0, tmp_print)
        block_name = "中巴士"
        tmp_print = car_count(3).ToString() + "車次"
        Call do_sql.print_block(F_file2, block_name, 0, 0, tmp_print)
        block_name = "大福特"
        tmp_print = car_count(4).ToString() + "車次"
        Call do_sql.print_block(F_file2, block_name, 0, 0, tmp_print)
        block_name = "八人座"
        tmp_print = car_count(5).ToString() + "車次"
        Call do_sql.print_block(F_file2, block_name, 0, 0, tmp_print)
        block_name = "小貨卡"
        tmp_print = car_count(6).ToString() + "車次"
        Call do_sql.print_block(F_file2, block_name, 0, 0, tmp_print)
        block_name = "1.6"
        tmp_print = car_count(7).ToString() + "車次"
        Call do_sql.print_block(F_file2, block_name, 0, 0, tmp_print)
        block_name = "1.8"
        tmp_print = car_count(8).ToString() + "車次"
        Call do_sql.print_block(F_file2, block_name, 0, 0, tmp_print)
        block_name = "2.0以上"
        tmp_print = car_count(9).ToString() + "車次"
        Call do_sql.print_block(F_file2, block_name, 0, 0, tmp_print)
        block_name = "租賃大巴士"
        tmp_print = car_count(10).ToString() + "車次"
        Call do_sql.print_block(F_file2, block_name, 0, 0, tmp_print)
        block_name = "租賃中巴士"
        tmp_print = car_count(11).ToString() + "車次"
        Call do_sql.print_block(F_file2, block_name, 0, 0, tmp_print)
        block_name = "租賃八人座"
        tmp_print = car_count(12).ToString() + "車次"
        Call do_sql.print_block(F_file2, block_name, 0, 0, tmp_print)

        block_name = "各車型差勤統計"
        tmp_print = tot_row.ToString() + "車次"
        Call do_sql.print_block(F_file2, block_name, 0, 0, tmp_print)
        block_name = "差勤人員統計"
        tmp_print = tot_row.ToString() + "員"
        Call do_sql.print_block(F_file2, block_name, 0, 0, tmp_print)
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated

        If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then
            '隱藏eformsn
            e.Row.Cells(14).Visible = False
        End If

    End Sub

    Protected Sub GridView1_RowUpdated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdatedEventArgs) Handles GridView1.RowUpdated

        'Mail通知申請人車號及駕駛人
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        Dim eformsn As String = ""
        Dim strDrive, strCarNum As String

        strDrive = CType(GridView1.Rows(strRowIndex).Cells(0).FindControl("DropDownList1"), DropDownList).SelectedItem.Text
        strCarNum = CType(GridView1.Rows(strRowIndex).Cells(0).FindControl("DropDownList9"), DropDownList).SelectedItem.Text

        '取得eformsn
        eformsn = GridView1.DataKeys(strRowIndex).Value

        '開啟連線
        Dim db As New SqlConnection(connstr)

        '表單流水號
        Dim P_NUM As String = ""

        '找出派車資料
        db.Open()
        Dim carCom As New SqlCommand("select * from P_03 where eformsn = '" & eformsn & "'", db)
        Dim RdvCar = carCom.ExecuteReader()
        If RdvCar.read() Then
            P_NUM = RdvCar.Item("P_NUM")
        End If
        db.Close()

        Dim FC As New C_FlowSend.C_FlowSend

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
        Dim peridCom As New SqlCommand("select PWIDNO from P_03 where eformsn = '" & eformsn & "'", db)
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

        '發送Mail給通知申請人派車資訊
        Dim MailBody As String = ""
        MailBody += "您所申請的派車申請單" & "<br>"
        MailBody += "表單流水號為:" & P_NUM & "<br>"
        MailBody += "駕駛人:" & strDrive & "<br>"
        MailBody += "車輛號碼:" & strCarNum & "<br>"

        ''判斷是否寄送Mail
        'If MailYN = "Y" And empemail <> "" Then
        '    FC.F_MailGO(SystemMail, "調度單位通知", SmtpHost, empemail, "車輛派遣通知", MailBody)
        'End If

        Call add_stmt()


    End Sub

    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GridView1.RowUpdating

        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string

        '取得RowIndex
        strRowIndex = e.RowIndex

        Dim strDrive, strCarNum As String
        Dim GoTimeSH, GoTimeSM, GoTimeBH, GoTimeBM As String

        strDrive = CType(GridView1.Rows(strRowIndex).Cells(0).FindControl("DropDownList1"), DropDownList).SelectedItem.Text
        strCarNum = CType(GridView1.Rows(strRowIndex).Cells(0).FindControl("DropDownList9"), DropDownList).SelectedItem.Text

        '出場時間
        GoTimeSH = CType(GridView1.Rows(strRowIndex).Cells(0).FindControl("DropDownList2"), DropDownList).SelectedItem.Text
        GoTimeSM = CType(GridView1.Rows(strRowIndex).Cells(0).FindControl("DropDownList3"), DropDownList).SelectedItem.Text

        '回場時間
        GoTimeBH = CType(GridView1.Rows(strRowIndex).Cells(0).FindControl("DropDownList4"), DropDownList).SelectedItem.Text
        GoTimeBM = CType(GridView1.Rows(strRowIndex).Cells(0).FindControl("DropDownList5"), DropDownList).SelectedItem.Text

        '有輸入駕駛或是車輛號碼則出場回場時間不可空白
        If strDrive <> "" Or strCarNum <> "" Then
            If GoTimeSH = "" Or GoTimeSM = "" Or GoTimeBH = "" Or GoTimeBM = "" Then

                e.Cancel = True

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('請選擇出場跟回場時間!!!');")
                Response.Write(" </script>")

            ElseIf GoTimeBH & GoTimeBM < GoTimeSH & GoTimeSM Then

                e.Cancel = True

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('回場時間不可小於出場時間!!!');")
                Response.Write(" </script>")

            Else

                '開啟連線
                Dim db As New SqlConnection(connstr)

                '駕駛重複出場
                Dim Drive_Flag As String = ""
                Dim DriveLT As String = ""
                Dim DriveCT As String = ""

                '找出派車資料
                db.Open()
                Dim DriveCom As New SqlCommand("select P_0301.LeaveTime,P_0301.ComeTime from P_03,P_0301 WHERE P_03.EFORMSN=P_0301.EFORMSN AND P_0301.DriveName='" & strDrive & "' AND (P_0301.DriveName IS NOT NULL) AND (DATEDIFF(Day, '" & Txt_nARRDATE.Text & "', nARRDATE) = 0)", db)
                Dim RdvDrive = DriveCom.ExecuteReader()
                If RdvDrive.read() Then
                    DriveLT = RdvDrive.Item("LeaveTime")
                    DriveCT = RdvDrive.Item("ComeTime")
                End If
                db.Close()

                '駕駛重複
                If (GoTimeSH & GoTimeSM > DriveLT) And (GoTimeSH & GoTimeSM < DriveCT) Then
                    Drive_Flag = "1"
                ElseIf (GoTimeSH & GoTimeSM < DriveLT) And (GoTimeBH & GoTimeBM > DriveLT) Then
                    Drive_Flag = "1"
                End If

                If Drive_Flag = "1" Then
                    e.Cancel = True

                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('駕駛重複選取出場!!!');")
                    Response.Write(" </script>")
                End If


                '車輛重複出場
                Dim Car_Flag As String = ""
                Dim CarLT As String = ""
                Dim CarCT As String = ""

                '找出派車資料
                db.Open()
                Dim CarCom As New SqlCommand("select P_0301.LeaveTime,P_0301.ComeTime from P_03,P_0301 WHERE P_03.EFORMSN=P_0301.EFORMSN AND P_0301.CarNumber='" & strCarNum & "' AND (P_0301.CarNumber IS NOT NULL) AND (DATEDIFF(Day, '" & Txt_nARRDATE.Text & "', nARRDATE) = 0)", db)
                Dim RdvCar = CarCom.ExecuteReader()
                If RdvCar.read() Then
                    CarLT = RdvCar.Item("LeaveTime")
                    CarCT = RdvCar.Item("ComeTime")
                End If
                db.Close()

                '車輛重複
                If (GoTimeSH & GoTimeSM > CarLT) And (GoTimeSH & GoTimeSM < CarCT) Then
                    Car_Flag = "1"
                ElseIf (GoTimeSH & GoTimeSM < CarLT) And (GoTimeBH & GoTimeBM > CarLT) Then
                    Car_Flag = "1"
                End If

                If Car_Flag = "1" Then
                    e.Cancel = True

                    Response.Write(" <script language='javascript'>")
                    Response.Write(" alert('車輛重複選取出場!!!');")
                    Response.Write(" </script>")
                End If

            End If

        End If

    End Sub

    Protected Sub DropDownList9_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        select_flag = True
    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        Call add_stmt()
    End Sub

    Protected Sub ImageButton1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "68px"
        Div_grid.Style("left") = "280px"

    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

        Div_grid.Visible = False

    End Sub

    Function SQLALL() As String

        stmt = "SELECT P_03.EFORMSN, P_0301.DriveName, P_0301.DriveTel, P_0301.CarNumber, P_03.nARRIVEPLACE + '--' + P_03.nARRIVETO AS per,"
        stmt += "P_03.nPHONE, P_03.nENDPOINT, P_03.nREASON, P_0301.LeaveTime,SUBSTRING(P_0301.LeaveTime,1,2) LeaveHour,SUBSTRING(P_0301.LeaveTime,3,2) LeaveMin ,"
        stmt += "P_0301.ComeTime,SUBSTRING(P_0301.ComeTime,1,2) ComeHour,SUBSTRING(P_0301.ComeTime,3,2) ComeMin, P_0301.LeaveMilage, "
        stmt += "P_0301.ComeMilage, P_0301.RealMilage, P_03.nSTYLE "
        stmt += " FROM P_03 INNER JOIN P_0301 ON P_03.EFORMSN = P_0301.EFORMSN "
        stmt += " where nARRDATE ='" + Txt_nARRDATE.Text + "'"

        If TXT_nREASON.Text <> "" Then
            stmt += " and P_03.nREASON like '" + TXT_nREASON.Text + "%'"
        End If
        If Txt_DriveName.Text <> "" Then
            stmt += " and P_0301.DriveName like '" + Txt_DriveName.Text + "%'"
        End If
        If DrDown_nstyle.SelectedItem.Text <> "" Then
            stmt += " and P_03.nSTYLE like '" + DrDown_nstyle.SelectedItem.Text + "%'"
        End If
        If Txt_CarNumber.Text <> "" Then
            stmt += " and P_0301.CarNumber like '" + Txt_CarNumber.Text + "%'"
        End If
        If Txt_nARRIVEPLACE.Text <> "" Then
            stmt += " and nARRIVEPLACE like '" + Txt_nARRIVEPLACE.Text + "%'"
        End If
        If Txt_nARRIVETO.Text <> "" Then
            stmt += " and nARRIVETO like '" + Txt_nARRIVETO.Text + "%'"
        End If

        SQLALL = stmt

    End Function

    Protected Sub Img_Export_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles Img_Export.Click
        Dim strsqlAll As String = SQLALL()

        '設定檔案路徑
        Dim path As String = "", filename As String = ""

        path = Server.MapPath("~/Drs/")
        filename = Date.Today.ToString("yyyyMMdd") & "_00200.csv"
        path = path & filename

        Dim colname As String = ""
        Dim data As String = ""
        Dim prn_stmt As String = ""

        Dim sw As New System.IO.StreamWriter(path, False, Encoding.GetEncoding("big5"))

        Dim T_date As String = CStr(CInt(Now.ToString("yyyy")) - 1911) + "年"
        Dim user_name As String = ""
        T_date += Now.ToString("MM") + "月" + Now.ToString("dd") + "日"
        print_file = ""
        If do_sql.select_urname(do_sql.G_user_id) = False Then
            Exit Sub
        End If
        If do_sql.G_usr_table.Rows.Count > 0 Then
            user_name = do_sql.G_usr_table.Rows(0).Item("emp_chinese_name").ToString.Trim
        End If

        prn_stmt = "國防部總務管理處中央調度室每日車輛派遣紀錄表"

        sw.WriteLine(prn_stmt)

        'prn_stmt = "調度官:," & user_name & ",,駐地:,,,日期:," & T_date & ",,," & ",,,"
        prn_stmt = "調度官:," & user_name & ",,駐地:,,,日期:,民國" & T_date & ",,," & ",,,"
        sw.WriteLine(prn_stmt)

        colname = "駕駛人姓名,車輛號碼,報到處何人,申派單位連絡電話,地址,目的,出場時間,回場時間,出場時路碼表里程,回場時路碼表里程,本次任務實際里程,備考(車型),"

        colname = Left(colname, Len(colname) - 1)

        sw.WriteLine(colname)


        Dim db As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

        db.Open()
        Dim dt2 As DataTable = New DataTable("")
        Dim da2 As SqlDataAdapter = New SqlDataAdapter(strsqlAll, db)
        da2.Fill(dt2)
        db.Close()
        Dim tmp_add As String = ""
        Dim car_count(13) As Integer
        Dim hi As Integer
        Dim tot_row As Integer = 0
        For hi = 0 To 13
            car_count(hi) = 0
        Next
        tot_row = dt2.Rows.Count
        For y As Integer = 0 To dt2.Rows.Count - 1

            data += dt2.Rows(y).Item("DriveName").ToString & "," & dt2.Rows(y).Item("CarNumber").ToString & ","
            data += dt2.Rows(y).Item("per").ToString & "," & dt2.Rows(y).Item("nPHONE").ToString & ","
            data += dt2.Rows(y).Item("nENDPOINT").ToString & "," & dt2.Rows(y).Item("nREASON").ToString & ","
            data += dt2.Rows(y).Item("LeaveTime").ToString & "," & dt2.Rows(y).Item("ComeTime").ToString & ","
            data += dt2.Rows(y).Item("LeaveMilage").ToString & "," & dt2.Rows(y).Item("ComeMilage").ToString & ","
            data += dt2.Rows(y).Item("RealMilage").ToString & "," & dt2.Rows(y).Item("nSTYLE").ToString & ","
            Select Case dt2.Rows(y).Item("nSTYLE").ToString
                Case "大巴士"
                    car_count(1) += 1
                Case "大貨卡"
                    car_count(2) += 1
                Case "中巴士"
                    car_count(3) += 1
                Case "大福特"
                    car_count(4) += 1
                Case "八人座"
                    car_count(5) += 1
                Case "小貨卡"
                    car_count(6) += 1
                Case "1.6轎車"
                    car_count(7) += 1
                Case "1.8轎車"
                    car_count(8) += 1
                Case "2.0轎車"
                    car_count(9) += 1
                Case "租賃大巴士"
                    car_count(10) += 1
                Case "租賃中巴士"
                    car_count(11) += 1
                Case "租賃八人座"
                    car_count(12) += 1
                Case Else
                    car_count(0) += 1
            End Select
            
           
            data = Left(data, Len(data) - 1)
            sw.WriteLine(data)
            data = ""

        Next
        sw.WriteLine("")
        prn_stmt = "民國" + T_date + "各車型差勤人員及車次統計,,,,,,各車型車輛狀況,,,,,,"
        sw.WriteLine(prn_stmt)
        prn_stmt = "大巴士:" + car_count(1).ToString & "車次,," + "大貨卡:" + car_count(2).ToString & "車次,," + "中巴士:" + car_count(3).ToString & "車次,," & ",,,,,,"
        sw.WriteLine(prn_stmt)
        prn_stmt = "大福特:" + car_count(4).ToString & "車次,," + "八人座:" + car_count(5).ToString & "車次,," + "小貨卡:" + car_count(6).ToString & "車次,," & ",,,,,,"
        sw.WriteLine(prn_stmt)
        prn_stmt = "1.6:" + car_count(7).ToString & "車次,," + "1.8:" + car_count(8).ToString & "車次,," + "2.0以上:" + car_count(9).ToString & "車次,," & ",,,,,,"
        sw.WriteLine(prn_stmt)
        prn_stmt = "租賃大巴士:" + car_count(10).ToString & "車次,," + "租賃中巴士:" + car_count(11).ToString & "車次,," + "租賃八人座:" + car_count(12).ToString & "車次,," & ",,,,,,"
        sw.WriteLine(prn_stmt)

        prn_stmt = "各車型差勤統計:共" + tot_row.ToString & "車次,,," + "差勤人員統計:共" + tot_row.ToString & "車次,,," & ",,,,,,"
        sw.WriteLine(prn_stmt)
        sw.WriteLine("")
        prn_stmt = "調度室呈核：中隊主官管核閱：大隊主官管核示：總管處核派主官："
        sw.WriteLine(prn_stmt)

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
