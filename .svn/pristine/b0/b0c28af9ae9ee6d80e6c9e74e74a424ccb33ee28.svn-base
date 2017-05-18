Imports System.Data.SqlClient
Imports System.Data


Partial Class M_Source_MOA04106
    Inherits System.Web.UI.Page
    Dim connstr As String = ""
    Dim HouseFlag As String = ""
    Dim SelFlag As Integer = 1
    Dim dt As Date = Now()
    Public default_QueryRange As String = " where (nStartDATE between (Convert(varchar(16),DateAdd(Day, DateDiff(Day, 14, getdate()), 0),120)) and getdate()) " '以開工日期
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

            '系統預設顯示兩周內的報修單(by nAPPTIME 申請日期)
            Me.SqlDataSource2.SelectCommand = SqlDataSource2.SelectCommand.Insert _
                                                           ( _
                                                                 SqlDataSource2.SelectCommand.IndexOf(" Order by"), _
                                                                 default_QueryRange _
                                                           )
        End If
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        Me.DataRecords_DataRebinding()
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
            Server.Transfer("MOA04107.aspx?P_Num=" & P_Num)
            'ElseIf SelFlag = 2 Then
            '    '開啟連線
            '    Dim db As New SqlConnection(connstr)
            '    '確定完工
            '    db.Open()
            '    Dim insCom As New SqlCommand("UPDATE P_0415 SET nFinalDate = getdate() WHERE(P_NUM = '" & P_Num & "')", db)
            '    insCom.ExecuteNonQuery()
            '    db.Close()
        End If
        Me.DataRecords_DataRebinding()
    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted
        Me.DataRecords_DataRebinding()
    End Sub

    Private Sub DataRecords_DataRebinding()
        SqlDataSource2.SelectParameters.Clear()
        Dim where_conditions As String = " where (1=1) "
        If Not String.IsNullOrEmpty(PANAME.Text.Trim()) Then
            '申請人姓名
            where_conditions += " And a.PANAME Like ('%' + @PANAME + '%')"
            SqlDataSource2.SelectParameters.Add(New Parameter("PANAME", DbType.String, PANAME.Text.Trim()))
        End If

        If Not String.IsNullOrEmpty(nPHONE.Text.Trim()) Then
            '聯絡電話
            where_conditions += " And a.nPHONE Like ('%' + @nPHONE + '%')"
            SqlDataSource2.SelectParameters.Add(New Parameter("nPHONE", DbType.String, nPHONE.Text.Trim()))
        End If
        '申請人單位
        If Not String.IsNullOrEmpty(PAUNIT.SelectedValue) Then
            where_conditions += " And a.PAUNIT =@PAUNIT"
            SqlDataSource2.SelectParameters.Add(New Parameter("PAUNIT", DbType.String, PAUNIT.SelectedValue))
        End If

        'If Session("Role") = "3" Then
        '    where_conditions += " and a.PAIDNO='" & Session("user_id") & "'"
        'Else

        'End If

        If Not String.IsNullOrEmpty(nStartDATE1.Text.Trim()) And Not String.IsNullOrEmpty(nStartDATE2.Text.Trim()) Then
            '申請時間
            where_conditions += " And (a.nStartDATE between @nStartDATE1 +' 00:00:00' and @nStartDATE2 + ' 23:59:59')"
            SqlDataSource2.SelectParameters.Add(New Parameter("nStartDATE1", DbType.String, nStartDATE1.Text.Trim()))
            SqlDataSource2.SelectParameters.Add(New Parameter("nStartDATE2", DbType.String, nStartDATE2.Text.Trim()))
        End If
        If Not String.IsNullOrEmpty(nPLACE.Text.Trim()) Then
            '請修地點
            where_conditions += " And b.bd_name+'/'+c.fl_name+'/'+d.rnum_name like '%' + @nPLACE + '%'"
            SqlDataSource2.SelectParameters.Add(New Parameter("nPLACE", DbType.String, nPLACE.Text.Trim()))
        End If
        If Not String.IsNullOrEmpty(nFIXITEM.Text.Trim()) Then
            '請修事項
            where_conditions += " And a.nFIXITEM like '%' + @nFIXITEM + '%'"
            SqlDataSource2.SelectParameters.Add(New Parameter("nFIXITEM", DbType.String, nFIXITEM.Text.Trim()))
        End If
        If ddl_Status.SelectedValue <> "0" Then
            '請修事項
            where_conditions += " And a.FlowStatus = @FlowStatus"
            SqlDataSource2.SelectParameters.Add(New Parameter("FlowStatus", DbType.String, ddl_Status.SelectedValue))
        End If

        If where_conditions <> String.Empty Then

            Me.SqlDataSource2.SelectCommand = SqlDataSource2.SelectCommand.Replace(default_QueryRange, "")
            Me.SqlDataSource2.SelectCommand = SqlDataSource2.SelectCommand.Insert _
                                                             ( _
                                                                   SqlDataSource2.SelectCommand.IndexOf(" Order by"), _
                                                                   where_conditions _
                                                             )
        End If
    End Sub
    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        Me.DataRecords_DataRebinding()
    End Sub

    Protected Sub ImgClearAll_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgClearAll.Click
        ddl_Status.SelectedIndex = 0
        PANAME.Text = ""
        PAUNIT.SelectedValue = ""
        nPHONE.Text = ""
        nStartDATE1.Text = ""
        nStartDATE2.Text = ""
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
        ' Me.DataRecords_DataRebinding()
    End Sub

    Protected Sub ImgList_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        '查詢詳細日期
        SelFlag = 1
    End Sub

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click
        Div_grid.Visible = True
        Div_grid.Style("Top") = "70px"
        Div_grid.Style("left") = "70px"
        If nStartDATE1.Text.Trim() = "" Then
            Calendar1.SelectedDate = dt.AddDays(-14).Date
        Else
            Calendar1.SelectedDate = nStartDATE1.Text
        End If
        nStartDATE1.Text = Calendar1.SelectedDate.Date
    End Sub

    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click
        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "70px"
        Div_grid2.Style("left") = "130px"
        If (nStartDATE2.Text.Trim() = "") Then
            Calendar2.SelectedDate = dt.Date
        Else
            Calendar2.SelectedDate = nStartDATE2.Text
        End If
        nStartDATE2.Text = Calendar2.SelectedDate.Date
    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged
        nStartDATE1.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False
    End Sub

    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged
        nStartDATE2.Text = Calendar2.SelectedDate.Date
        Div_grid2.Visible = False
    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click
        Div_grid.Visible = False
    End Sub

    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click
        Div_grid2.Visible = False
    End Sub
    Public Function StatusName(ByVal FlowStatus As String) As String
        Select Case FlowStatus
            Case "1"
                Return "新送件"
            Case "2"
                Return "處理中"
            Case "3"
                Return "待料中"
            Case "4"
                Return "完工"
            Case Else
                Return String.Empty
        End Select
    End Function
    Public Function CheckShowDATE(ByVal nDataDATE As String) As String
        Dim sResult As String = String.Empty
        If nDataDATE.Trim().Length <> 0 Then
            Try
                Dim dt As New DateTime
                dt = DateTime.Parse(nDataDATE)
                Dim sDataDT As String = dt.ToString("yyyy/MM/dd")
                If DateTime.Parse(sDataDT) >= DateTime.Parse("2000/01/01") Then
                    sResult = DateTime.Parse(sDataDT).ToString("yyyy/MM/dd")
                End If
            Catch ex As Exception
                sResult = "error"
            End Try
        End If
        Return sResult
    End Function
End Class
