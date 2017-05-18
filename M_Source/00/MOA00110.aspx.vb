Imports System.Data.sqlclient
Partial Class Source_00_MOA00110
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String
    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click

        Dim strOrd As String

        '排序條件
        strOrd = " ORDER BY Login_Num DESC"

        SqlDataSource1.SelectCommand = SQLALL(LoginIP.Text, Sdate.Text, Edate.Text) & strOrd

    End Sub

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

                If LoginCheck.LoginCheck(user_id, "MOA00110") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00110.aspx")
                    Response.End()
                End If

                If IsPostBack = False Then

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

                    '一登入馬上查詢
                    ImgSearch_Click(Nothing, Nothing)

                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Public Function SQLALL(ByVal ipSel, ByVal SDate, ByVal EDate)

        '整合SQL搜尋字串
        Dim strsql, strDate, strip As String
        strip = ""
        strsql = "SELECT * FROM LoginLog WHERE 1=1 "

        'ip搜尋
        If LoginIP.Text <> "" Then
            strip = " AND (Login_IP like '%" & LoginIP.Text & "%')"
        End If

        '申請日期搜尋
        strDate = " AND (Login_Time between '" & SDate & " 00:00:00 ' AND '" & EDate & " 23:59:59')"

        SQLALL = strsql & strDate & strip

    End Function

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged

        '分頁
        Dim strOrd As String

        '排序條件
        strOrd = " ORDER BY Login_Num DESC"

        SqlDataSource1.SelectCommand = SQLALL(LoginIP.Text, Sdate.Text, Edate.Text) & strOrd


    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted

        '排序
        Dim strOrd As String

        '排序條件
        strOrd = " ORDER BY " & GridView1.SortExpression.ToString()

        SqlDataSource1.SelectCommand = SQLALL(LoginIP.Text, Sdate.Text, Edate.Text) & strOrd

    End Sub

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "70px"
        Div_grid.Style("left") = "90px"

        Calendar1.SelectedDate = Sdate.Text

    End Sub

    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click

        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "70px"
        Div_grid2.Style("left") = "205px"

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
