Imports System.Data.sqlclient
Partial Class Source_04_MOA04002
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        Search(False)
    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted
        Search(False)
    End Sub

    Sub Search(ByVal sort As Boolean)
        If (sort) Then
            GridView1.Sort("Block_Name,Block_Unit", SortDirection.Ascending)
            Return
        End If
        Dim sql As String
        sql = "SELECT Block_Name,Block_Unit,sum(Block_Amount) Block_Amount"
        sql += "  FROM P_0401 CROSS JOIN P_0402"
        sql += " where P_0401.EFORMSN=P_0402.EFORMSN"
        If (nAPPTIME1.Text <> "") Then sql += " and P_0401.nAPPTIME>=CONVERT(datetime,'" + nAPPTIME1.Text + " 00:00:00')" '申請時間
        If (nAPPTIME2.Text <> "") Then sql += " and P_0401.nAPPTIME<=CONVERT(datetime,'" + nAPPTIME2.Text + " 23:59:59')" '申請時間
        If (Block_Name.Text <> "") Then sql += " and P_0402.Block_Name like '%" + Trim(Block_Name.Text) + "%'" '物品
        If (Block_Unit.Text <> "") Then sql += " and P_0402.Block_Unit like '%" + Trim(Block_Unit.Text) + "%'" '單位
        sql += " group by Block_Name,Block_Unit"

        SqlDataSource2.SelectCommand = sql
    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        Search(True)
    End Sub

    Protected Sub ImgClearAll_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgClearAll.Click
        nAPPTIME1.Text = ""
        nAPPTIME2.Text = ""
        Block_Name.Text = ""
        Block_Unit.Text = ""
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

                Dim nAPPTIME1 As TextBox = Me.FindControl("nAPPTIME1")
                Dim nAPPTIME2 As TextBox = Me.FindControl("nAPPTIME2")
                Dim dt As Date = Now()

                If (nAPPTIME1.Text = "") Then
                    nAPPTIME1.Text = dt.AddDays(-14).Date
                End If
                If (nAPPTIME2.Text = "") Then
                    nAPPTIME2.Text = dt.Date
                End If

                '登入馬上查詢
                ImgSearch_Click(Nothing, Nothing)

            End If

        End If

    End Sub

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "70px"
        Div_grid.Style("left") = "60px"

        Calendar1.SelectedDate = nAPPTIME1.Text

    End Sub

    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click

        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "70px"
        Div_grid2.Style("left") = "170px"

        Calendar2.SelectedDate = nAPPTIME2.Text

    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

        nAPPTIME1.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False

    End Sub

    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged

        nAPPTIME2.Text = Calendar2.SelectedDate.Date
        Div_grid2.Visible = False

    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

        Div_grid.Visible = False

    End Sub

    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click

        Div_grid2.Visible = False

    End Sub
End Class
