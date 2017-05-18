Imports System.Data.sqlclient
Partial Class Source_04_MOA04004
    Inherits System.Web.UI.Page
    Dim connstr As String = ""
    Dim HouseFlag As String = ""
    Dim SelFlag As Integer = 1
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

            '維修時間
            Dim nFIXDATE1 As TextBox = Me.FindControl("nFIXDATE1")
            Dim nFIXDATE2 As TextBox = Me.FindControl("nFIXDATE2")
            Dim dt As Date = Now()

            If (nFIXDATE1.Text = "") Then
                nFIXDATE1.Text = dt.AddDays(-14).Date
            End If
            If (nFIXDATE2.Text = "") Then
                nFIXDATE2.Text = dt.Date
            End If

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

            '新增連線
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string

            '開啟連線
            Dim db As New SqlConnection(connstr)

            '判斷登入者是否為房屋水電相關群組人員
            db.Open()
            Dim carCom As New SqlCommand("select * from ROLEGROUPITEM where employee_id = '" & Session("user_id") & "' AND (Group_Uid ='5K6V1X2I65' OR Group_Uid ='W334922GX1')", db)
            Dim RdvCar = carCom.ExecuteReader()
            If RdvCar.read() Then
                HouseFlag = "1"
            End If
            db.Close()

        End If

    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        Search(False)
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated

        If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then

            '隱藏流水號
            e.Row.Cells(0).Visible = False

        End If

    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

        Dim P_Num As Integer = GridView1.Rows(GridView1.SelectedIndex).Cells(0).Text

        If SelFlag = 1 Then

            '轉換到查詢派工日期詳細頁面
            Server.Transfer("MOA04005.aspx?P_Num=" & P_Num)

        ElseIf SelFlag = 2 Then


            '開啟連線
            Dim db As New SqlConnection(connstr)

            '確定完工
            db.Open()
            Dim insCom As New SqlCommand("UPDATE P_04 SET nFinalDate = getdate() WHERE(P_NUM = '" & P_Num & "')", db)
            insCom.ExecuteNonQuery()
            db.Close()

        End If

        Search(True)

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

        sql = "SELECT distinct PAUNIT, convert(nvarchar,nAPPTIME,111) nAPPTIME, PANAME, PAUNIT,convert(nvarchar,nFIXDATE,111) nFIXDATE,P_NUM,nPLACE,nFIXITEM,nPHONE,nFinalDate"
        sql += "  FROM P_04"
        sql += " where 1=1 "

        If (PANAME.Text <> "") Then sql += " and P_04.PANAME like '%" + Trim(PANAME.Text) + "%'" '申請人姓名
        If (nPHONE.Text <> "") Then sql += " and P_04.nPHONE like '%" + Trim(nPHONE.Text) + "%'" '聯絡電話

        If Session("Role") = "3" Then
            sql += " and PAIDNO='" & Session("user_id") & "'"
        Else
            If (PAUNIT.Text <> "") Then sql += " and P_04.PAUNIT like '%" + Trim(PAUNIT.Text) + "%'" '申請人單位
        End If

        If (nFIXDATE1.Text <> "") Then sql += " and P_04.nFIXDATE>=CONVERT(datetime,'" + nFIXDATE1.Text + " 00:00:00')" '維修時間
        If (nFIXDATE2.Text <> "") Then sql += " and P_04.nFIXDATE<=CONVERT(datetime,'" + nFIXDATE2.Text + " 23:59:59')" '維修時間
        If (nPLACE.Text <> "") Then sql += " and P_04.nPLACE like '%" + Trim(nPLACE.Text) + "%'" '請修地點
        If (nFIXITEM.Text <> "") Then sql += " and P_04.nFIXITEM like '%" + Trim(nFIXITEM.Text) + "%'" '請修事項

        SqlDataSource2.SelectCommand = sql
    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        Search(True)
    End Sub

    Protected Sub ImgClearAll_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgClearAll.Click
        
        PANAME.Text = ""
        'PAUNIT.SelectedItem.Text = ""
        nPHONE.Text = ""
        nFIXDATE1.Text = ""
        nFIXDATE2.Text = ""
        nFIXITEM.Text = ""
        nPLACE.Text = ""

    End Sub

    Protected Sub Page_PreRenderComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRenderComplete

        If Not IsPostBack Then
            Dim tLItm As New ListItem("請選擇", "")

            '系統管理員組織加請選擇
            If Session("Role") = "1" Or HouseFlag = "1" Then

                PAUNIT.Items.Insert(0, tLItm)
                If PAUNIT.Items.Count > 1 Then
                    PAUNIT.Items(1).Selected = False
                End If

            End If

        End If

        Search(False)

    End Sub

    Protected Sub ImgList_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        '查詢詳細日期
        SelFlag = 1

    End Sub

    Protected Sub ImgFinal_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        '確認完工
        SelFlag = 2

    End Sub

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "70px"
        Div_grid.Style("left") = "70px"

        Calendar1.SelectedDate = nFIXDATE1.Text

    End Sub

    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click

        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "70px"
        Div_grid2.Style("left") = "130px"

        Calendar2.SelectedDate = nFIXDATE2.Text

    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

        nFIXDATE1.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False

    End Sub

    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged

        nFIXDATE2.Text = Calendar2.SelectedDate.Date
        Div_grid2.Visible = False

    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

        Div_grid.Visible = False

    End Sub

    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click

        Div_grid2.Visible = False

    End Sub
End Class
