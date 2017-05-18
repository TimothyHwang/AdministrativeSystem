Imports System.Data.sqlclient
Partial Class Source_07_MOA07002
    Inherits System.Web.UI.Page

    Public do_sql As New C_SQLFUN
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

            Dim nAPPTIME1 As TextBox = Me.FindControl("nAPPTIME1")
            Dim nAPPTIME2 As TextBox = Me.FindControl("nAPPTIME2")
            Dim dt As Date = Now()

            If (nAPPTIME1.Text = "") Then
                nAPPTIME1.Text = dt.AddDays(-14).Date
            End If
            If (nAPPTIME2.Text = "") Then
                nAPPTIME2.Text = dt.Date
            End If

            '找出上一級單位以下全部單位
            Dim Org_Down As New C_Public

            '判斷登入者權限
            If Session("Role") = "1" Then
                SqlDataSource1.SelectCommand = "select '' as ORG_UID,'' as ORG_NAME union SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
            ElseIf Session("Role") = "2" Then
                SqlDataSource1.SelectCommand = "select '' as ORG_UID,'' as ORG_NAME union SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(Session("PARENT_ORG_UID")) & ") ORDER BY ORG_NAME"
            Else
                SqlDataSource1.SelectCommand = "select '' as ORG_UID,'' as ORG_NAME union SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID ='" & org_uid & "' ORDER BY ORG_NAME"
            End If

            '登入馬上查詢
            ImgSearch_Click(Nothing, Nothing)

        End If


    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Server.Transfer("MOA07003.aspx?EFORMSN=" & GridView1.SelectedValue)
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        Search(False)
    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted
        Search(False)
    End Sub

    Sub Search(ByVal sort As Boolean)
        If (sort) Then
            GridView1.Sort("nAPPTIME", SortDirection.Descending)
            Return
        End If
        Dim sql As String
        sql = "SELECT distinct P_NUM,PAUNIT,PANAME, convert(nvarchar,nAPPTIME,111) nAPPTIME, nTel,nSeat, P_07.EFORMSN EFORMSN"
        sql += "  FROM P_07 CROSS JOIN P_0701"
        sql += " where P_07.EFORMSN=P_0701.EFORMSN"
        If (PAUNIT.Text <> "") Then sql += " and P_07.PAUNIT like '%" + Trim(PAUNIT.Text) + "%'" '申請人單位
        If (nAssetNum.Text <> "") Then sql += " and P_0701.nAssetNum like '%" + Trim(nAssetNum.Text) + "%'" '財產編號
        If (nLabel.Text <> "") Then sql += " and P_0701.nLabel like '%" + Trim(nLabel.Text) + "%'" '標籤
        If (nLabelNum.Text <> "") Then sql += " and P_0701.nLabelNum like '%" + Trim(nLabelNum.Text) + "%'" '標籤號碼
        If (nAPPTIME1.Text <> "") Then sql += " and P_07.nAPPTIME>=CONVERT(datetime,'" + nAPPTIME1.Text + " 00:00:00')" '申請時間
        If (nAPPTIME2.Text <> "") Then sql += " and P_07.nAPPTIME<=CONVERT(datetime,'" + nAPPTIME2.Text + " 23:59:59')" '申請時間
        If (nAssetName.Text <> "") Then sql += " and P_0701.nAssetName like '%" + Trim(nAssetName.Text) + "%'" '財產名稱
        If (nREASON.Text <> "") Then sql += " and P_0701.nREASON like '%" + Trim(nREASON.Text) + "%'" '問題類別

        SqlDataSource2.SelectCommand = sql
    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        Search(True)
    End Sub

    Protected Sub ImgClearAll_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgClearAll.Click
        PAUNIT.Text = ""
        nAssetNum.Text = ""
        nLabel.Text = ""
        nLabelNum.Text = ""
        nAPPTIME1.Text = ""
        nAPPTIME2.Text = ""
        nAssetName.Text = ""
        nREASON.Text = ""
    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        '登入馬上查詢
        ImgSearch_Click(Nothing, Nothing)

    End Sub

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "90px"
        Div_grid.Style("left") = "80px"

        Calendar1.SelectedDate = nAPPTIME1.Text

    End Sub

    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click

        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "90px"
        Div_grid2.Style("left") = "180px"

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
