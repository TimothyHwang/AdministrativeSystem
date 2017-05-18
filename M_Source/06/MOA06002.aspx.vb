Imports System.Data.sqlclient
Partial Class Source_06_MOA06002
    Inherits System.Web.UI.Page
    Dim connstr As String
    Dim infoFlag As String = ""
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

                '先設定申請起始日期
                Dim dt As Date = Now()
                If (nAPPTIME1.Text = "") Then
                    nAPPTIME1.Text = dt.AddDays(-14).Date
                End If

                If (nAPPTIME2.Text = "") Then
                    nAPPTIME2.Text = dt.Date
                End If

                '先設定出入起始日期
                If (nDATE1.Text = "") Then
                    nDATE1.Text = dt.AddDays(-14).Date
                End If

                If (nDATE2.Text = "") Then
                    nDATE2.Text = dt.Date
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

                '找出上一級單位以下全部單位
                Dim Org_Down As New C_Public

                '判斷登入者權限
                If Session("Role") = "1" Then
                    SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
                Else
                    SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP WHERE ORG_UID IN (" & Org_Down.getchildorg(org_uid) & ") ORDER BY ORG_NAME"
                End If


                '登入馬上查詢
                ImgSearch_Click(Nothing, Nothing)

            End If

        End If

    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Server.Transfer("MOA06003.aspx?EFORMSN=" & GridView1.SelectedValue)
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        Search(False)
    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted
        Search(False)
    End Sub

    Sub Search(ByVal sort As Boolean)

        If (sort) Then
            GridView1.Sort("nDATE", SortDirection.Descending)
            Return
        End If

        Dim sql As String
        Dim strsql As String = ""
        Dim ParentFlag As String = ""

        sql = "SELECT distinct P_NUM,PAUNIT,PAIDNO, convert(nvarchar,nDATE,111) nDATE,convert(nvarchar,nAPPTIME,111) nAPPTIME, nREASON, nPLACE, P_06.EFORMSN EFORMSN"
        sql += "  FROM P_06 CROSS JOIN P_0601"
        sql += " where P_06.EFORMSN=P_0601.EFORMSN"
        If (nAPPTIME1.Text <> "") Then sql += " and P_06.nAPPTIME>=CONVERT(datetime,'" + nAPPTIME1.Text + " 00:00:00')" '申請時間
        If (nAPPTIME2.Text <> "") Then sql += " and P_06.nAPPTIME<=CONVERT(datetime,'" + nAPPTIME2.Text + " 23:59:59')" '申請時間
        If (PANAME.Text <> "") Then sql += " and P_06.PANAME like '%" + Trim(PANAME.Text) + "%'" '申請人姓名
        If (PAUNIT.Text <> "") Then sql += " and P_06.PAUNIT like '%" + Trim(PAUNIT.Text) + "%'" '申請人單位
        'If (nMName.Text <> "") Then sql += " and P_0601.nMName like '%" + Trim(nMName.Text) + "%'" '名稱機型

        If (nDATE1.Text <> "") Then sql += " and P_06.nDATE>=CONVERT(datetime,'" + nDATE1.Text + " 00:00:00')" '申請出入日期
        If (nDATE2.Text <> "") Then sql += " and P_06.nDATE<=CONVERT(datetime,'" + nDATE2.Text + " 23:59:59')" '申請出入日期
        If (nClass.Text <> "") Then sql += " and P_0601.nClass = '" + Trim(nClass.Text) + "'" '機密等級
        If (nREASON.Text <> "") Then sql += " and P_06.nREASON like '%" + Trim(nREASON.Text) + "%'" '申請事由
        If (nPLACE.Text <> "") Then sql += " and P_06.nPLACE like '%" + Trim(nPLACE.Text) + "%'" '地點

        '判斷登入者權限
        If Session("Role") = "1" Then
        Else

            Dim conn As New C_SQLFUN
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

            If ParentFlag <> "Y" Then
                strsql = " and PAIDNO='" & Session("user_id") & "'"
            End If

        End If

        SqlDataSource2.SelectCommand = sql & strsql
    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        Search(True)
    End Sub

    Protected Sub ImgClearAll_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgClearAll.Click
        nAPPTIME1.Text = ""
        nAPPTIME2.Text = ""
        PANAME.Text = ""
        'PAUNIT.SelectedItem.Text = ""
        'nMName.Text = ""
        nDATE1.Text = ""
        nDATE2.Text = ""
        nClass.Text = ""
        nREASON.Text = ""
        nPLACE.Text = ""
    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        If Not IsPostBack Then
            Dim tLItm As New ListItem("請選擇", "")

            '系統管理員組織加請選擇
            If Session("Role") = "1" Or infoFlag = "1" Then

                PAUNIT.Items.Insert(0, tLItm)
                If PAUNIT.Items.Count > 1 Then
                    PAUNIT.Items(1).Selected = False
                End If

            End If

        End If

    End Sub

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "80px"
        Div_grid.Style("left") = "75px"

        'Calendar1.SelectedDate = Txt_nHandleDate.Text

    End Sub

    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click

        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "80px"
        Div_grid2.Style("left") = "175px"

        'Calendar2.SelectedDate = Txt_nCheckDate.Text

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

    Protected Sub ImgDate3_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate3.Click

        Div_grid3.Visible = True
        Div_grid3.Style("Top") = "115px"
        Div_grid3.Style("left") = "75px"

        'Calendar3.SelectedDate = Txt_nDoleDate.Text

    End Sub

    Protected Sub ImgDate4_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate4.Click

        Div_grid4.Visible = True
        Div_grid4.Style("Top") = "115px"
        Div_grid4.Style("left") = "175px"

        'Calendar4.SelectedDate = Txt_nFinishDate.Text

    End Sub

    Protected Sub Calendar3_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar3.SelectionChanged

        nDATE1.Text = Calendar3.SelectedDate.Date
        Div_grid3.Visible = False

    End Sub

    Protected Sub Calendar4_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar4.SelectionChanged

        nDATE2.Text = Calendar4.SelectedDate.Date
        Div_grid4.Visible = False

    End Sub

    Protected Sub btnClose3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose3.Click

        Div_grid3.Visible = False

    End Sub

    Protected Sub btnClose4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose4.Click

        Div_grid4.Visible = False

    End Sub
End Class
