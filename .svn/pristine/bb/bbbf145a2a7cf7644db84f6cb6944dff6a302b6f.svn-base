Imports System.Data
Imports System.IO

Partial Class M_Source_01_MOA01014
    Inherits Page

    Public Shared SelAll As String      '是否全選
    Dim IsSuperior As Boolean = False
    Dim user_id, org_uid, org_name As String
    Public do_sql As New C_SQLFUN
    Dim ParentFlag As String = ""

    Public Function SQLALL(ByVal OrgSel As String, ByVal UserSel As String, ByVal SDate As String, ByVal EDate As String) As String

        Try
            Dim strsel As String = ""
            '整合SQL搜尋字串
            Dim strsql, strDate As String
            strsql = "SELECT P_NUM,PAUNIT,PATITLE,PANAME,PAIDNO,STARTTIME,STHOUR,ETHOUR,LOCATION,REASON FROM P_0107,EMPLOYEE WHERE PAIDNO=employee_id"
            '系統管理員
            If Session("Role") = "1" Then
                '組織搜尋
                If OrgSel = "" Then
                    strsel = " AND 1=1"
                Else
                    strsel = " AND ORG_UID='" & OrgSel & "'"
                End If
                '人員
                If UserSel <> "" Then
                    strsel += " AND PAIDNO = '" & UserSel & "'"
                End If
                '申請日期搜尋
                strDate = " AND (STARTTIME BETWEEN '" & SDate & "' AND '" & EDate & "')"
            Else
                If ParentFlag = "Y" Then
                    '主官管
                    '組織搜尋
                    If OrgSel <> "" Then
                        strsel = " AND ORG_UID='" & OrgSel & "'"
                    End If
                    '人員
                    If UserSel <> "" Then
                        strsel += " AND PAIDNO = '" & UserSel & "'"
                    End If
                Else
                    '人員
                    strsel = " AND PAIDNO = '" & Session("user_id").ToString() & "'"
                End If
                '申請日期搜尋
                strDate = " AND (STARTTIME BETWEEN '" & SDate & "' AND '" & EDate & "')"
            End If
            SQLALL = strsql & strsel & strDate
        Catch ex As Exception
            SQLALL = ""
        End Try
    End Function

    ''' <summary>
    ''' 取得排序的轉換字串
    ''' </summary>
    ''' <param name="sortDirection"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function getSortDirectionString(ByVal sortDirection As SortDirection) As String
        Dim newSortDirection As String = String.Empty
        If sortDirection = sortDirection.Ascending Then
            newSortDirection = "ASC"
        Else
            newSortDirection = "DESC"
        End If
        getSortDirectionString = newSortDirection
    End Function

    Protected Function GetSuperiorRight(ByVal User_ID As String) As Boolean
        Dim boolSuccess As Boolean = False
        boolSuccess = do_sql.db_exec("SELECT * FROM ADMINGROUP,EMPLOYEE WHERE ADMINGROUP.ORG_UID=EMPLOYEE.ORG_UID AND EMPLOYEE.EMPLOYEE_ID='" & User_ID & "'", do_sql.G_conn_string)
        If boolSuccess Then
            GetSuperiorRight = False
        End If
        GetSuperiorRight = True
    End Function

    Protected Function CheckCondition() As Boolean
        Dim boolReturn As Boolean = True
        ''檢查單位
        ''檢查姓名
        ''檢查起始日期
        If txtSdate.Text.Length = 0 Then boolReturn = False
        ''檢查終迄日期
        If txtEdate.Text.Length = 0 Then boolReturn = False

        CheckCondition = boolReturn
    End Function

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        'org_uid = CType(IIf(Session("ORG_UID") = Nothing, "", Session("ORG_UID").ToString()), String)
        If Session("ORG_UID") = Nothing Then org_uid = "" Else org_uid = Session("ORG_UID").ToString()
        If Session("user_id") = Nothing Then user_id = "" Else user_id = Session("user_id").ToString()

        If Not IsPostBack Then

            '找出上一級單位以下全部單位
            Dim Org_Down As New C_Public

            '顯示使用者資訊            
            lblUser.Text = Org_Down.GetUserNameByID(user_id)
            lblOrg.Text = Org_Down.GetOrgNameByID(org_uid)

            ''設定預設查詢日期
            If txtSdate.Text.Length = 0 Then txtSdate.Text = DateTime.Now.ToShortDateString()
            If txtEdate.Text.Length = 0 Then txtEdate.Text = DateTime.Now.ToShortDateString()


        End If
    End Sub

    Protected Sub ImgSearch_Click(sender As Object, e As ImageClickEventArgs) Handles ImgSearch.Click
        Dim strOrd As String

        If CheckCondition() Then
            strOrd = " ORDER BY STARTTIME DESC"

            SqlDataSource2.SelectCommand = SQLALL(org_uid, user_id, txtSdate.Text, txtEdate.Text) & strOrd
            Dim view As DataView = CType(SqlDataSource2.Select(New DataSourceSelectArguments), DataView)

            If view.Count > 0 Then AllChk.Enabled = True Else AllChk.Enabled = False
        End If
        'OrgChange = ""
    End Sub

    Protected Sub gvList_PreRender(sender As Object, e As System.EventArgs) Handles gvList.PreRender
        If Session("P01014PageNo") IsNot Nothing Then gvList.PageIndex = CType(Session("P01014PageNo"), Integer) - 1
        If Session("P01014PageSize") IsNot Nothing Then gvList.PageSize = CType(Session("P01014PageSize"), Integer)
    End Sub

    Protected Sub ddlPageNo_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        Session("P01014PageNo") = CType(sender, DropDownList).Text
        gvList.PageIndex = CInt(CType(sender, DropDownList).Text) - 1
        AllChk.Checked = False
        SqlDataSource2.SelectCommand = SQLALL(org_uid, user_id, txtSdate.Text, txtEdate.Text) & " ORDER BY STARTTIME DESC"
        'gvList.DataBind()
    End Sub

    Protected Sub ddlPageSize_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        Session("P01014PageSize") = CType(sender, DropDownList).Text
        gvList.PageSize = CInt(CType(sender, DropDownList).Text)
        AllChk.Checked = False
        SqlDataSource2.SelectCommand = SQLALL(org_uid, user_id, txtSdate.Text, txtEdate.Text) & " ORDER BY STARTTIME DESC"
    End Sub

    Protected Sub gvList_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvList.RowCommand
        Select Case e.CommandName
            Case "Edit"
                Dim Parameters = e.CommandArgument.ToString()
                Response.Redirect("MOA01013.ASPX?P_NUM=" & Parameters & "&ACTION=2")
            Case "Delete"
                Dim Parameters = e.CommandArgument.ToString()
                Dim strSql As String = "DELETE P_0107 WHERE P_NUM='" & Parameters & "'"
                If do_sql.db_exec(strSql, do_sql.G_conn_string) = True Then
                    'SqlDataSource2.SelectCommand = SQLALL(org_uid, user_id, txtSdate.Text, txtEdate.Text) & " ORDER BY STARTTIME DESC"
                    Response.Redirect("MOA01014.ASPX")
                End If
        End Select
    End Sub

    Protected Sub gvList_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvList.RowDataBound
        If e.Row.RowType = DataControlRowType.Pager Then
            ''設定GridView頁碼列
            Dim ddlPageNo As DropDownList = CType(e.Row.FindControl("ddlPageNo"), DropDownList)
            Dim lblPageCount As Label = CType(e.Row.FindControl("lblPageCount"), Label)
            Dim ddlPageSize As DropDownList = CType(e.Row.FindControl("ddlPageSize"), DropDownList)

            ddlPageNo.Items.Clear()
            For i As Integer = 1 To gvList.PageCount
                ddlPageNo.Items.Add(i)
            Next
            ddlPageNo.Text = CType((gvList.PageIndex + 1), String)
            lblPageCount.Text = lblPageCount.Text & gvList.PageCount.ToString()
            ddlPageSize.Text = CType(gvList.PageSize, String)
        ElseIf e.Row.RowType = DataControlRowType.DataRow Then
            Dim btnEdit As Button = CType(e.Row.FindControl("btnEdit"), Button)
            Dim btnDelete As Button = CType(e.Row.FindControl("btnDelete"), Button)
            btnEdit.CommandArgument = CType(e.Row.FindControl("hdnP_NUM"), HiddenField).Value
            btnDelete.CommandArgument = CType(e.Row.FindControl("hdnP_NUM"), HiddenField).Value
        End If
    End Sub

    Protected Sub ImgPrint_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgPrint.Click
        Dim F_file As String
        Dim F_file2 As String
        Dim F_file_name As String
        Dim strOrd As String = ""
        Dim strSql As String = ""
        Dim dtPrint As New DataTable
        Dim SelectedRowNo As String = ""

        Dim print_file As String
        Dim print As New C_Xprint

        'strSql = SQLALL(org_uid, user_id, txtSdate.Text, txtEdate.Text) & " ORDER BY STARTTIME DESC"
        For i As Integer = 0 To gvList.Rows.Count - 1
            If CType(gvList.Rows(i).Cells(0).FindControl("selchk"), CheckBox).Checked = True Then
                If SelectedRowNo.Length > 0 Then SelectedRowNo = SelectedRowNo & ","
                SelectedRowNo = SelectedRowNo & "'" & CType(gvList.Rows(i).Cells(0).FindControl("hdnP_NUM"), HiddenField).Value & "'"
            End If
        Next
        strSql = "SELECT * FROM P_0107 WHERE P_NUM IN (" & SelectedRowNo & ") ORDER BY " & IIf(gvList.SortExpression.Length = 0, "STARTTIME", gvList.SortExpression) & " " & IIf(gvList.SortExpression.Length = 0, "DESC", getSortDirectionString(gvList.SortDirection))

        If do_sql.db_sql(strSql, do_sql.G_conn_string) = False Then
            do_sql.G_errmsg = "查詢資料失敗"
            Exit Sub
        End If
        If do_sql.G_table.Rows.Count > 0 Then
            dtPrint = do_sql.G_table
        Else
            do_sql.G_errmsg = "查無可印資料"
            Exit Sub
        End If


        Dim prn_stmt As String

        F_file_name = "rpt01001"
        F_file = MapPath("../../rpt/" + F_file_name + ".txt")
        If File.Exists(F_file) = False Then
            do_sql.G_errmsg = "檔案不存在"
            Exit Sub
        End If

        Dim filename As String = F_file_name & Rnd() & ".drs"
        ''print_file = "../../drs/" + F_file_name + "-" + Now.ToString("mmssff") & ".drs"
        F_file2 = MapPath("../../drs/" & filename)

        print.C_Xprint("rpt01001.txt", filename)
        print.NewPage()

        Call do_sql.inc_file(F_file, F_file2, F_file_name)
        Dim nPage As Integer = 0
        Dim n_line As Integer = 0
        Dim yShift = 3
        If dtPrint.Rows.Count > 0 Then
            nPage = 1
            Call print.Add("總人數", dtPrint.Rows.Count.ToString(), 0, 0)
            For Each dr As DataRow In dtPrint.Rows
                If n_line = 14 AndAlso nPage = 1 Then
                    nPage = nPage + 1
                    print.NewPage("rpt01001_2.txt")
                    n_line = 0
                ElseIf n_line = 18 AndAlso nPage > 1 Then
                    nPage = nPage + 1
                    print.NewPage("rpt01001_2.txt")
                    n_line = 0
                End If

                'If n_line >= 14 Then                
                '    Call do_sql.print_sdata(F_file2, "/init " + F_file_name + ".txt 1", "/newpage null")
                '    prn_stmt = CDate(Sdate.Text).ToString("yyyy/MM/dd") & "~" & CDate(Edate.Text).ToString("yyyy/MM/dd")
                '    Call do_sql.print_block(F_file2, "日期:", 0, 0, prn_stmt)
                '    Call do_sql.print_block(F_file2, "加班:", 0, 0, DrDn_add.SelectedValue)
                '    n_line = 0
                'End If
                Call print.Add("單位名稱", dr("PAUNIT").ToString.Substring(0, dr("PAUNIT").ToString.IndexOf("(")), 0, 0)

                Call print.Add("級職", dr("PATITLE").ToString, 0, n_line * 13 + yShift)
                Call print.Add("姓名", dr("PANAME").ToString(), 0, n_line * 13 + yShift)
                Call print.Add("加班事由", dr("REASON").ToString, 0, n_line * 13 + yShift)
                Call print.Add("加班地點", dr("LOCATION").ToString, 0, n_line * 13 + yShift)
                Call print.Add("加班起日", CType(dr("STARTTIME"), DateTime).Year - 1911 & "年" & CType(dr("STARTTIME"), DateTime).Month & "月" & CType(dr("STARTTIME"), DateTime).Day & "日", 0, n_line * 13 + yShift)
                Call print.Add("加班起時", dr("STHOUR").ToString() & " 時起", 0, n_line * 13 + yShift)
                Call print.Add("加班迄時", dr("ETHOUR").ToString() & " 時止", 0, n_line * 13 + yShift)
                Call print.Add("申請人簽名", dr("PWNAME"), 0, n_line * 13 + yShift)

                n_line = n_line + 1
            Next

        End If


        print.EndFile()
        If (print.ErrMsg <> "") Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('" & print.ErrMsg & "');")
            Response.Write("</script>")
        Else
            Response.Write("<script language='javascript'>")
            Response.Write("window.onload = function() {")
            Response.Write("window.location.replace('../../drs/" & filename & "');")
            Response.Write("}")
            Response.Write("</script>")
        End If
    End Sub

    Protected Sub AllChk_CheckedChanged(sender As Object, e As System.EventArgs) Handles AllChk.CheckedChanged
        '是否全選
        Dim i As Integer = 0
        If AllChk.Checked = False Then
            For i = 0 To gvList.Rows.Count - 1
                CType(gvList.Rows(i).Cells(0).FindControl("selchk"), CheckBox).Checked = False
            Next
            '沒全選
            SelAll = ""
        Else
            For i = 0 To gvList.Rows.Count - 1
                CType(gvList.Rows(i).Cells(0).FindControl("selchk"), CheckBox).Checked = True
            Next
            '全選
            SelAll = "1"
        End If
    End Sub

    Protected Sub gvList_Sorting(sender As Object, e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvList.Sorting
        AllChk.Checked = False
        SqlDataSource2.SelectCommand = SQLALL(org_uid, user_id, txtSdate.Text, txtEdate.Text) & " ORDER BY " & e.SortExpression & " " & getSortDirectionString(e.SortDirection)
    End Sub
End Class
