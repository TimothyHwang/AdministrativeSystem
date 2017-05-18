Imports System.Data.sqlclient
Partial Class Source_05_MOA05003
    Inherits System.Web.UI.Page

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Server.Transfer("MOA05004.aspx?EFORMSN=" & GridView1.SelectedValue & "&SelDate=" & nRECDATE.Text)
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        Search(False)
    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted
        Search(False)
    End Sub

    Sub Search(ByVal sort As Boolean)

        If (sort) Then
            GridView1.Sort("nRECDATE", SortDirection.Ascending)
            Return
        End If

        Dim sql As String
        sql = "SELECT distinct P_NUM,convert(nvarchar,nRECDATE,111) nRECDATE,nSTARTTIME+'~'+nENDTIME nSTARTTIME, nREASON,PAUNIT, PANAME,nPHONE,nSTATUS,nRECEXIT, P_05.EFORMSN EFORMSN"
        sql += "  FROM P_05 CROSS JOIN P_0501"
        sql += " where P_05.EFORMSN=P_0501.EFORMSN AND PENDFLAG='E'"

        If (nRECDATE.Text <> "") Then sql += " and DateDiff(day,P_05.nRECDATE,CONVERT(datetime,'" + nRECDATE.Text + "'))=0" '會客時間
        If (nName.Text <> "") Then sql += " and P_0501.nName like '%" + Trim(nName.Text) + "%'" '姓名
        If (nID.Text <> "") Then sql += " and P_0501.nID like '%" + Trim(nID.Text) + "%'" '身分證字號
        If (nSTATUS.Text <> "") Then sql += " and P_05.nSTATUS like '%" + Trim(nSTATUS.Text) + "%'" '狀態
        'If (nREASON.Text <> "") Then sql += " and P_05.nREASON like '%" + Trim(nREASON.Text) + "%'" '事由
        If (PAUNIT.Text <> "") Then sql += " and P_05.PAUNIT like '%" + Trim(PAUNIT.Text) + "%'" '單位
        If (PANAME.Text <> "") Then sql += " and P_05.PANAME like '%" + Trim(PANAME.Text) + "%'" '被會人
        If (nPHONE.Text <> "") Then sql += " and P_05.nPHONE like '%" + Trim(nPHONE.Text) + "%'" '電話
        If (DDLnRECEXIT.Text <> "") Then sql += " and P_05.nRECEXIT = '" + Trim(DDLnRECEXIT.Text) + "'" '會客入口

        SqlDataSource1.SelectCommand = sql
    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        Search(True)
    End Sub

    Protected Function Chg_nSTATUS(ByVal str As String) As String
        Try
            Dim tmpStr = Eval(str)
            If tmpStr = "0" Then
                tmpStr = "尚未進營"
            ElseIf tmpStr = "1" Then
                tmpStr = "正在營中"
            ElseIf tmpStr = "2" Then
                tmpStr = "已經出營"
            End If
            Chg_nSTATUS = tmpStr
        Catch ex As Exception
            Chg_nSTATUS = ""
        End Try
    End Function

    Protected Sub ImgClearAll_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgClearAll.Click
        Dim dt As Date = Now()
        nRECDATE.Text = dt.Date
        nName.Text = ""
        nID.Text = ""
        nSTATUS.Text = ""
        'nREASON.Text = ""
        PAUNIT.Text = ""
        PANAME.Text = ""
        nPHONE.Text = ""
        DDLnRECEXIT.Text = ""
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            If IsPostBack = False Then

                Dim nRECDATE As TextBox = Me.FindControl("nRECDATE")
                Dim dt As Date = Now()

                If (nRECDATE.Text = "") Then
                    nRECDATE.Text = dt.Date
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        '增加請選擇選項
        If Not IsPostBack Then
            Dim tLItm As New ListItem("請選擇", "")
            DDLnRECEXIT.Items.Insert(0, tLItm)
            If DDLnRECEXIT.Items.Count > 1 Then
                DDLnRECEXIT.Items(1).Selected = False
            End If

            nSTATUS.Items.Insert(0, tLItm)
            If nSTATUS.Items.Count > 1 Then
                nSTATUS.Items(1).Selected = False
            End If

            '登入馬上查詢
            ImgSearch_Click(Nothing, Nothing)

        End If
    End Sub

End Class
