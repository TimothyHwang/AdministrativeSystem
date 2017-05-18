
Partial Class M_Source_04_MOA04007
    Inherits System.Web.UI.Page
    Dim PAUNIT As String = ""
    Dim Sdate As String = ""
    Dim Edate As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        PAUNIT = Request("PAUNIT")
        Sdate = Request("Sdate")
        Edate = Request("Edate")

        If IsPostBack = False Then

            Dim strOrd As String = ""

            strOrd = " ORDER BY P_04.PAUNIT,P_04.PANAME"

            SqlDataSource2.SelectCommand = SQLALL(PAUNIT, Sdate, Edate) & strOrd

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


        strsql = "select P_04.PAUNIT,P_04.PANAME,P_04.nPHONE,P_04.nPLACE,P_04.nFIXITEM, "
        strsql += " case P_0401.pendflag when 'E' then '完成' else '' end As pendflag  "
        strsql += " from P_04 left outer join  P_0401 on P_04.eformsn = P_0401.neformsn  "
        strsql += " where P_04.pendflag='E' "
        If OrgSel <> "" Then
            strsql += " AND P_04.PAUNIT='" & OrgSel & "'"
        End If
        strsql += " and P_04.nAPPTIME >= '" & SDate & "' and  P_04.nAPPTIME <= '" & EDate & "'"

        SQLALL = strsql

    End Function
    Protected Sub BackBtn_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BackBtn.Click

        '回上頁
        'Server.Transfer("MOA04006.aspx")

        'Response.Write("<script language='javascript'>")
        'Response.Write("window.history.go('-2');")
        'Response.Write("</script>")
        Response.Redirect("MOA04006.aspx")
    End Sub
    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged

        '分頁
        Dim strOrd As String = ""

        strOrd = " ORDER BY P_04.PAUNIT,P_04.PANAME"

        SqlDataSource2.SelectCommand = SQLALL(PAUNIT, Sdate, Edate) & strOrd

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        '排序
        Dim strOrd As String = ""

        strOrd = " ORDER BY " & GridView1.SortExpression.ToString()

        SqlDataSource2.SelectCommand = SQLALL(PAUNIT, Sdate, Edate) & strOrd

    End Sub

    
End Class
