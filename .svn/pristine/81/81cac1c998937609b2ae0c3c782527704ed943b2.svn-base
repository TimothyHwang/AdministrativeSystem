Imports System.Data.sqlclient
Partial Class Source_06_MOA06004
    Inherits System.Web.UI.Page

    Dim infoFlag As String = ""
    Dim connstr As String
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

            If IsPostBack = False Then

                '找出上一級單位以下全部單位
                Dim Org_Down As New C_Public

                '判斷登入者權限
                If Session("Role") = "1" Then
                    SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
                ElseIf Session("Role") = "2" Then
                    SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(Session("PARENT_ORG_UID")) & ") ORDER BY ORG_NAME"
                Else
                    SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID ='" & org_uid & "' ORDER BY ORG_NAME"
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

                '新增連線
                Dim conn As New C_SQLFUN
                connstr = conn.G_conn_string

                '開啟連線
                Dim db As New SqlConnection(connstr)

                '判斷登入者是否為資訊媒體攜出入群組人員
                db.Open()
                Dim carCom As New SqlCommand("select * from ROLEGROUPITEM where employee_id = '" & user_id & "' AND (Group_Uid ='2128TF461K')", db)
                Dim RdvCar = carCom.ExecuteReader()
                If RdvCar.read() Then
                    infoFlag = "1"
                End If
                db.Close()

            End If

        End If

    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        If Not IsPostBack Then
            Dim tLItm As New ListItem("請選擇", "")

            '系統管理員組織加請選擇
            If Session("Role") = "1" Or infoFlag = "1" Then

                OrgSel.Items.Insert(0, tLItm)
                If OrgSel.Items.Count > 1 Then
                    OrgSel.Items(1).Selected = False
                End If

            End If

            '登入馬上查詢
            ImgSearch_Click(Nothing, Nothing)

        End If

    End Sub

    Public Function SQLALL(ByVal OrgSel, ByVal SDate, ByVal EDate)

        '整合SQL搜尋字串
        Dim strsql, strsel, strDate As String
        strsql = "SELECT P_06.PAUNIT, SUM(V_P06_InfoSum.Kind1) AS Kind1, SUM(V_P06_InfoSum.Kind2) AS Kind2, SUM(V_P06_InfoSum.Kind3) AS Kind3, SUM(V_P06_InfoSum.Kind4) AS Kind4, SUM(V_P06_InfoSum.Kind5) AS Kind5, SUM(V_P06_InfoSum.Kind6) AS Kind6 FROM P_06 INNER JOIN V_P06_InfoSum ON P_06.EFORMSN = V_P06_InfoSum.EFORMSN "
        strsel = ""

        '組織搜尋
        If OrgSel = "請選擇" Then
            strsel = " WHERE 1=1"
        Else
            strsel = " WHERE P_06.PAUNIT = '" & OrgSel & "'"
        End If

        '申請日期搜尋
        strDate = " AND (V_P06_InfoSum.nAPPTIME between '" & SDate & " 00:00:00 ' AND '" & EDate & " 23:59:59')"

        SQLALL = strsql & strsel & strDate

    End Function


    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click

        Dim strOrd, strgroup As String

        strgroup = " GROUP BY P_06.PAUNIT "

        strOrd = " ORDER BY P_06.PAUNIT "

        SqlDataSource2.SelectCommand = SQLALL(OrgSel.SelectedItem.Text, Sdate.Text, Edate.Text) & strgroup & strOrd


    End Sub

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged

        '分頁
        Dim strOrd, strgroup As String

        strgroup = " GROUP BY P_06.PAUNIT "

        '排序條件
        strOrd = " ORDER BY P_06.PAUNIT "

        SqlDataSource2.SelectCommand = SQLALL(OrgSel.SelectedItem.Text, Sdate.Text, Edate.Text) & strgroup & strOrd


    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        Dim strOrd, strgroup As String

        strgroup = " GROUP BY P_06.PAUNIT "

        '排序條件
        strOrd = " ORDER BY " & GridView1.SortExpression.ToString()

        SqlDataSource2.SelectCommand = SQLALL(OrgSel.SelectedItem.Text, Sdate.Text, Edate.Text) & strgroup & strOrd

    End Sub

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "70px"
        Div_grid.Style("left") = "450px"

        Calendar1.SelectedDate = Sdate.Text

    End Sub

    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click

        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "70px"
        Div_grid2.Style("left") = "560px"

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
