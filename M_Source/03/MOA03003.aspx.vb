Imports System.Data
Imports System.Data.SqlClient

Partial Class Source_03_MOA03003
    Inherits System.Web.UI.Page

    Dim EFORMSN As String
    Dim sql As String
    Dim UpPrt As Boolean = True
    Dim user_id, org_uid As String

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        Search(False)
    End Sub

    '更新或列印
    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        If (UpPrt) Then '更新
            Dim BtnCarStatus As Button = Me.GridView1.Rows.Item(GridView1.SelectedIndex).FindControl("BtnCarStatus")

            SqlDataSource1.UpdateParameters("EFORMSN").DefaultValue = Me.GridView1.SelectedDataKey.Item(0)
            If BtnCarStatus.Text = "出場" Then
                SqlDataSource1.UpdateParameters("CarStatus").DefaultValue = "已出場"
            Else
                SqlDataSource1.UpdateParameters("CarStatus").DefaultValue = BtnCarStatus.Text
            End If
            SqlDataSource1.Update()
            Search(False)
        Else '列印
            EFORMSN = GridView1.SelectedDataKey.Item(0)
            Dim filename As String = "prt_030030" & Rnd() & ".drs"
            Dim print As New C_Xprint
            print.C_Xprint("rpt030030.txt", filename)
            print.NewPage()
            Prt_03(print) '增加列印內容
            print.EndFile()
            If (print.ErrMsg <> "") Then '產生列印檔過程有錯誤
                Response.Write("<script language='javascript'>")
                Response.Write("alert('" & print.ErrMsg & "');")
                Response.Write("</script>")
            Else '顯示報表檔
                Response.Write("<script language='javascript'>")
                Response.Write("window.onload = function() {")
                Response.Write("window.location.replace('../../drs/" & filename & "');")
                Response.Write("}")
                Response.Write("</script>")
            End If
        End If
    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted
        Search(False)
    End Sub

    Sub Search(ByVal sort As Boolean) 'GridView1之顯示資料
        If (sort) Then
            GridView1.Sort("DriveName,DriveTel", SortDirection.Ascending)
            Return
        End If
        Dim sql, sqlord As String

        sql = "SELECT distinct P_NUM,DriveName,DriveTel,nAPPLYTIME,nSTYLE,CarNumber,nARRIVEPLACE,nUSEDATE,nSTUSEHOUR,nSTUSEMIN,nEDUSEDATE,nEDUSEHOUR,nEDUSEMIN,nARRDATE,"
        sql += "CarStatus,P_0301.EFORMSN EFORMSN,LeaveTime,ComeTime,"
        sql += "SUBSTRING(LeaveTime,1,2) LeaveTime1,SUBSTRING(LeaveTime,3,2) LeaveTime2,SUBSTRING(ComeTime,1,2) ComeTime1,SUBSTRING(ComeTime,3,2) ComeTime2"
        sql += "  FROM P_03 CROSS JOIN P_0301"
        sql += " where P_03.EFORMSN=P_0301.EFORMSN"
        If (DriveName.Text <> "") Then sql += " and P_0301.DriveName like '%" + Trim(DriveName.Text) + "%'" '駕駛人姓名
        If (DriveTel.Text <> "") Then sql += " and P_0301.DriveTel like '%" + Trim(DriveTel.Text) + "%'" '駕駛人電話
        If (CarNumber.Text <> "") Then sql += " and P_0301.CarNumber like '%" + Trim(CarNumber.Text) + "%'" '車輛號碼
        If (nARRDATE1.Text <> "") Then sql += " and P_03.nARRDATE>=CONVERT(datetime,'" + Trim(nARRDATE1.Text) + " 00:00:00')" '車輛報到日期
        If (nARRDATE2.Text <> "") Then sql += " and P_03.nARRDATE<=CONVERT(datetime,'" + Trim(nARRDATE2.Text) + " 23:59:59')" '車輛報到日期
        If (CarStatus.Text <> "") Then sql += " and P_0301.CarStatus like '%" + Trim(CarStatus.Text) + "%'" '狀態
        If (nARRIVEPLACE.Text <> "") Then sql += " and P_03.nARRIVEPLACE like'%" + Trim(nARRIVEPLACE.Text) + "%'" '車輛報到地點
        If (nSTYLE.Text <> "") Then sql += " and P_03.nSTYLE='" + Trim(nSTYLE.Text) + "'" '車輛品名型式

        '車輛報到時間
        If (nARRDATE1.Text <> "" And nARRDATE2.Text <> "") Then
            sql += " and P_03.nEDUSEDATE>=CONVERT(datetime,'" & nARRDATE1.Text & " 00:00:00') and P_03.nUSEDATE<=CONVERT(datetime,'" & nARRDATE2.Text & " 23:59:59')"
        ElseIf (nARRDATE1.Text <> "") Then
            sql += " and P_03.nEDUSEDATE>=CONVERT(datetime,'" & nARRDATE1.Text & " 00:00:00') and P_03.nUSEDATE<=CONVERT(datetime,'" & nARRDATE1.Text & " 23:59:59')"
        ElseIf (nARRDATE2.Text <> "") Then
            sql += " and P_03.nEDUSEDATE>=CONVERT(datetime,'" & nARRDATE2.Text & " 00:00:00') and P_03.nUSEDATE<=CONVERT(datetime,'" & nARRDATE2.Text & " 23:59:59')"
        End If

        sqlord = " ORDER BY nAPPLYTIME DESC "

        SqlDataSource1.SelectCommand = sql & sqlord
    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        Search(True)
    End Sub

    '車輛報到時間之顯示格式
    Protected Function ShowDate(ByVal sd As String, ByVal sh As String, ByVal sm As String, ByVal ed As String, ByVal eh As String, ByVal em As String) As String
        Dim Sdate As String
        Dim Edate As String
        Dim tmpSTR As String
        Try
            Sdate = CDate(Eval(sd) & " " & Eval(sh) & ":" & Eval(sm)).ToString("yyyy/MM/dd")
        Catch ex As Exception
            Sdate = ""
        End Try
        Try
            Edate = CDate(Eval(ed) & " " & Eval(eh) & ":" & Eval(em)).ToString("yyyy/MM/dd")
        Catch ex As Exception
            Edate = ""
        End Try

        tmpSTR = Sdate & "~" & Edate
        If (tmpSTR.Equals("~")) Then
            tmpSTR = ""
        End If

        ShowDate = tmpSTR
    End Function

    Protected Function ShowCarStatus() As String '狀態若為null顯示為未出場
        Dim CarStatus As String = "未出場"
        Try
            If (Not Eval("CarStatus").Equals("")) Then
                CarStatus = Eval("CarStatus")
            End If
        Catch ex As Exception
        End Try
        ShowCarStatus = CarStatus
    End Function

    Protected Function CarStatusColor() As System.Drawing.Color '狀態顯示之Color
        Dim tmpColor As System.Drawing.Color = Drawing.Color.Black
        Dim CarStatus As String = "未出場"
        Dim dDate As Date = Now()
        Dim Sdate As Date
        Dim Edate As Date

        Try
            If (Not Eval("CarStatus").Equals("")) Then
                CarStatus = Eval("CarStatus")
            End If
        Catch ex As Exception
        End Try

        Try
            If (CarStatus.Equals("未出場")) Then
                Try
                    Sdate = CDate(Eval("nUSEDATE") & " " & Eval("LeaveTime1") & ":" & Eval("LeaveTime2"))
                    If (dDate > Sdate) Then
                        tmpColor = Drawing.Color.Red
                    End If
                Catch ex As Exception
                End Try
            ElseIf (CarStatus.Equals("已出場")) Then
                Try
                    Edate = CDate(Eval("nEDUSEDATE") & " " & Eval("ComeTime1") & ":" & Eval("ComeTime2"))
                    If (dDate > Edate) Then
                        tmpColor = Drawing.Color.Red
                    End If
                Catch ex As Exception
                End Try
            End If
        Catch ex As Exception
        End Try
        CarStatusColor = tmpColor
    End Function

    Protected Function CarStatusText() As String '狀態更新Button之Text
        Dim tmpText As String = "未出場"
        Dim CarStatus As String = "未出場"

        Try
            If (Not Eval("CarStatus").Equals("")) Then
                CarStatus = Eval("CarStatus")
            End If
        Catch ex As Exception

        End Try
        If (CarStatus.Equals("未出場")) Then
            tmpText = "出場" '已出場
        ElseIf (CarStatus.Equals("已出場")) Then
            tmpText = "回場"
        End If

        CarStatusText = tmpText
    End Function

    Protected Function CarStatusBtn() As Boolean '狀態更新之Button是否顯示
        Dim tmpBtn As Boolean = True
        Dim CarStatus As String = "未出場"

        Try
            If (Not Eval("CarStatus").Equals("")) Then
                CarStatus = Eval("CarStatus")
            End If
        Catch ex As Exception

        End Try
        If (CarStatus.Equals("回場")) Then
            tmpBtn = False
        End If

        CarStatusBtn = tmpBtn
    End Function

    Protected Sub Prt_03(ByVal prt As C_Xprint) '列印
        Dim dv As DataView
        Dim Sdate As String
        Dim Edate As String
        Dim tmpSTR As String

        sql += "SELECT *"
        sql += " FROM P_03,P_0301"
        sql += " where P_03.EFORMSN='" + EFORMSN + "'"
        sql += " and P_03.EFORMSN=P_0301.EFORMSN"
        SqlDataSource3.SelectCommand = sql
        dv = CType(SqlDataSource3.Select(DataSourceSelectArguments.Empty), DataView)

        If (dv.Count() > 0) Then
            prt.Add("填表人單位", dv.Table.Rows(0)("PWUNIT").ToString, 0, 0)
            prt.Add("填表人級職", dv.Table.Rows(0)("PWTITLE").ToString, 0, 0)
            prt.Add("填表人姓名", dv.Table.Rows(0)("PWNAME").ToString, 0, 0)
            prt.Add("申請人單位", dv.Table.Rows(0)("PAUNIT").ToString, 0, 0)
            prt.Add("申請人姓名", dv.Table.Rows(0)("PANAME").ToString, 0, 0)
            prt.Add("申請人級職", dv.Table.Rows(0)("PATITLE").ToString, 0, 0)
            prt.Add("申請時間", CDate(dv.Table.Rows(0)("nAPPLYTIME")).ToString("yyyy/MM/dd hh:mm"), 0, 0)
            prt.Add("聯絡電話", dv.Table.Rows(0)("nPHONE").ToString, 0, 0)
            prt.Add("事由", dv.Table.Rows(0)("nREASON").ToString, 0, 0)
            prt.Add("人員項目", dv.Table.Rows(0)("nITEM").ToString, 0, 0)
            prt.Add("車輛報到地點", dv.Table.Rows(0)("nARRIVEPLACE").ToString, 0, 0)
            prt.Add("向何人報到", dv.Table.Rows(0)("nARRIVETO").ToString, 0, 0)
            prt.Add("車次數", dv.Table.Rows(0)("nCARNUM").ToString, 0, 0)
            prt.Add("使用車輛數", dv.Table.Rows(0)("nUSECARNUM").ToString, 0, 0)
            prt.Add("起點", dv.Table.Rows(0)("nSTARTPOINT").ToString, 0, 0)
            prt.Add("目的地", dv.Table.Rows(0)("nENDPOINT").ToString, 0, 0)
            Try
                Sdate = CDate(CDate(dv.Table.Rows(0)("nARRDATE").ToString).ToString("yyyy/MM/dd") & " " & dv.Table.Rows(0)("nSTHOUR").ToString & ":" & dv.Table.Rows(0)("nEDHOUR").ToString).ToString("yyyy/MM/dd hh:mm")
            Catch ex As Exception
                Sdate = ""
            End Try
            prt.Add("車輛報到日期", Sdate, 0, 0)
            prt.Add("車輛狀態", dv.Table.Rows(0)("nSTATUS").ToString, 0, 0)
            prt.Add("車輛品名型式", dv.Table.Rows(0)("nSTYLE").ToString, 0, 0)
            Try
                Sdate = CDate(CDate(dv.Table.Rows(0)("nUSEDATE").ToString).ToString("yyyy/MM/dd") & " " & dv.Table.Rows(0)("nSTUSEHOUR").ToString & ":" & dv.Table.Rows(0)("nSTUSEMIN").ToString).ToString("yyyy/MM/dd hh:mm")
            Catch ex As Exception
                Sdate = ""
            End Try
            Try
                Edate = CDate(CDate(dv.Table.Rows(0)("nEDUSEDATE").ToString).ToString("yyyy/MM/dd") & " " & dv.Table.Rows(0)("nEDUSEHOUR").ToString & ":" & dv.Table.Rows(0)("nEDUSEMIN").ToString).ToString("yyyy/MM/dd hh:mm")
            Catch ex As Exception
                Edate = ""
            End Try
            tmpSTR = Sdate & "~" & Edate
            If (tmpSTR.Equals("~")) Then
                tmpSTR = ""
            End If
            prt.Add("任務使用時間", tmpSTR, 0, 0)
            prt.Add("駕駛人姓名", dv.Table.Rows(0)("DriveName").ToString, 0, 0)
            prt.Add("駕駛人電話", dv.Table.Rows(0)("DriveTel").ToString, 0, 0)
            prt.Add("車輛號碼", dv.Table.Rows(0)("CarNumber").ToString, 0, 0)
            prt.Add("出場時間", dv.Table.Rows(0)("LeaveTime").ToString, 0, 0)
            prt.Add("回場時間", dv.Table.Rows(0)("ComeTime").ToString, 0, 0)
            prt.Add("出場里數", dv.Table.Rows(0)("LeaveMilage").ToString, 0, 0)
            prt.Add("回場里數", dv.Table.Rows(0)("ComeMilage").ToString, 0, 0)
            prt.Add("任務實際里程", dv.Table.Rows(0)("RealMilage").ToString, 0, 0)
            prt.Add("狀態", dv.Table.Rows(0)("CarStatus").ToString, 0, 0)
        End If

        sql = "SELECT GoLocal,EndLocal"
        sql += " FROM P_0305"
        sql += " where EFORMSN='" + EFORMSN + "'"
        sql += " order by Local_Num"
        SqlDataSource3.SelectCommand = sql
        dv = CType(SqlDataSource3.Select(DataSourceSelectArguments.Empty), DataView)

        Dim cnt As Integer = dv.Count()
        Dim i As Integer = 0
        Dim row As Integer = 0
        Dim h As Integer = 10
        Do While (i < cnt And i < 8)
            prt.Add("行程起點", dv.Table.Rows(i)("GoLocal"), 0, row * h)
            prt.Add("行程迄點", dv.Table.Rows(i)("EndLocal"), 0, row * h)
            i += 1
            row += 1
        Loop
    End Sub

    Protected Sub ImgPrint_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        UpPrt = False
    End Sub

    Protected Sub ImgClearAll_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgClearAll.Click
        
        DriveName.Text = ""
        DriveTel.Text = ""
        CarNumber.Text = ""
        CarStatus.Text = ""
        nARRIVEPLACE.Text = ""
        nSTYLE.Text = ""

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

                '設定Default日期
                Dim nARRDATE1 As TextBox = Me.FindControl("nARRDATE1")
                Dim nARRDATE2 As TextBox = Me.FindControl("nARRDATE2")

                Dim dt As Date = Now()
                If (nARRDATE1.Text = "") Then
                    nARRDATE1.Text = dt.AddDays(-7).Date
                End If

                If (nARRDATE2.Text = "") Then
                    nARRDATE2.Text = dt.AddDays(7).Date
                End If

                '登入馬上查詢
                ImgSearch_Click(Nothing, Nothing)
            End If

        End If


    End Sub

    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "70px"
        Div_grid.Style("left") = "100px"

        Calendar1.SelectedDate = nARRDATE1.Text

    End Sub

    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click

        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "70px"
        Div_grid2.Style("left") = "210px"

        Calendar2.SelectedDate = nARRDATE2.Text

    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

        nARRDATE1.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False

    End Sub

    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged

        nARRDATE2.Text = Calendar2.SelectedDate.Date
        Div_grid2.Visible = False

    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

        Div_grid.Visible = False

    End Sub

    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click

        Div_grid2.Visible = False

    End Sub
End Class
