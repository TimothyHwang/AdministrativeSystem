Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing

Partial Class M_Source_10_MOA10006
    Inherits Page
    Const iRowCount = 10
    Const iColCount = 3
    Const TableWidth = 100 '1024
    Const TableHeight = 760
    Public MarqueeMsg As String
    ReadOnly sql_function As New C_SQLFUN
    ReadOnly connstr As String = sql_function.G_conn_string
    Dim db As New SqlConnection(connstr)

#Region "Custom Function"
    ''**********************  以下為Custom Function  **************************
    ''' <summary>
    ''' 建立主官在營狀態資料表格
    ''' </summary>
    ''' <param name="phr"></param>
    ''' <remarks></remarks>
    Protected Sub MakeStatusData(ByVal phr As PlaceHolder)
        ''改為縱向排序
        ''取出相對應的table
        Dim arrDataTableList As ArrayList = New ArrayList()
        Dim strSql As String = ""
        For i As Integer = 0 To iColCount - 1
            strSql = "SELECT TOP " & iRowCount & " P_NUM,NAME,STATUS FROM P_10 WHERE P_NUM NOT IN (SELECT TOP " & iRowCount * i & " P_NUM FROM P_10 ORDER BY RANKID) ORDER BY RANKID"
            db.Open()
            Dim DR As SqlDataReader = New SqlCommand(strSql, db).ExecuteReader
            Dim dt = New DataTable
            If DR.HasRows Then
                dt.Load(DR)
            End If
            db.Close()
            arrDataTableList.Add(dt)
        Next

        Dim strResponse = ""
        ''

        strResponse = "<table width=""100%"" Style=""border:0px solid #fff"" Height=""" & (TableHeight - 50) & """ BorderColor=""white""" & """ CellSpacing=""0"" CellPadding=""0"">"
        strResponse += "<tr>"
        For i As Integer = 0 To iColCount - 1
            Dim iItemCount = 0
            strResponse += "<td valign=""top"">"
            strResponse += "<table width=""100%"" CellSpacing=""0"" CellPadding=""0"">"
            If i < arrDataTableList.Count Then
                For Each row As DataRow In (CType(arrDataTableList(i), DataTable)).Rows
                    Dim sTemp As String = ""
                    strResponse += "<tr>"
                    If (row("Status").ToString = "0") Then
                        strResponse += "<td class=""imgtd"" width=""60"" height=""60"">&nbsp;</td>"

                        strResponse += "<td style=""border: 1px solid #FFF;"" align=""center""><table width=""100%""><tr><td class=""nametd"">"
                        For letter = 0 To Len(row("Name").ToString) - 1
                            sTemp = sTemp + "" + row("Name").ToString.Substring(letter, 1)
                        Next
                        strResponse += sTemp & "</td></tr></table></td>"
                    Else
                        strResponse += "<td class=""imgtd"" width=""60"" height=""60"" align=""center""><img src=""" & ResolveUrl("~/image/OnlineCircle.png") & """ width=""40""/></td>"

                        strResponse += "<td style=""border: 1px solid #FFF;"" align=""center""><table width=""100%""><tr><td class=""nametd"">"
                        For letter = 0 To Len(row("Name").ToString) - 1
                            sTemp = sTemp + "" + row("Name").ToString.Substring(letter, 1)
                        Next
                        strResponse += sTemp & "</td></tr></table></td>"
                    End If
                    iItemCount += 1

                    'strResponse += "</td>"
                    strResponse += "</tr>"
                Next
                If iItemCount < iRowCount Then
                    For iCount As Integer = 0 To (iRowCount - iItemCount) - 1
                        strResponse += "<tr><td class=""imgtd"" width=""70"" height=""70"">&nbsp;</td>"
                        strResponse += "<td style=""border: 1px solid #FFF;"" align=""center"">&nbsp;</td></tr>"
                    Next
                End If
            End If
            strResponse += "</table>"
            strResponse += "</td>"
        Next
        strResponse += "</tr>"
        strResponse += "</table>"
        Dim ltl As New Literal
        ltl.Text = strResponse
        phr.Controls.Add(ltl)

        'db.Open()
        'Const strSql As String = "SELECT TOP 27 P_NUM,NAME,STATUS FROM P_10 ORDER BY RANKID"
        'Dim DR As SqlDataReader = New SqlCommand(strSql, db).ExecuteReader
        'Dim dt = New DataTable
        'If DR.HasRows Then
        '    dt.Load(DR)
        'End If
        'db.Close()
        'Dim iTotaldRowCount As Integer = dt.Rows.Count
        'Dim iUsedRowCount As Integer = 0
        'Dim strResponse = ""

        'strResponse = "<table width=""100%"" Style=""border:1px solid #fff"" Height=""" & (TableHeight - 50) & """ BorderColor=""white""" & """ CellSpacing=""0"" CellPadding=""0"">"
        'For i = 0 To iRowCount - 1
        '    strResponse += "<tr>"
        '    For j = 0 To iColCount - 1
        '        iUsedRowCount += 1
        '        If iUsedRowCount <= iTotaldRowCount Then
        '            Dim sTemp As String = ""
        '            Dim s As String = dt.Rows(iUsedRowCount - 1).Item("Name").ToString.ToString()
        '            If (dt.Rows(iUsedRowCount - 1).Item("Status").ToString = "0") Then
        '                strResponse += "<td class=""imgtd"" width=""70"" height=""70"">&nbsp;</td>"
        '                For letter = 0 To Len(s) - 1
        '                    sTemp = sTemp + " " + s.Substring(letter, 1)
        '                Next
        '                strResponse += "<td class=""nametd"" width=""270""> " & sTemp & " </td>"
        '            Else
        '                strResponse += "<td class=""imgtd"" width=""70"" height=""70"" align=""center""><img src=""" & ResolveUrl("~/image/OnlineCircle.png") & """ width=""40""/></td>"
        '                For letter = 0 To Len(s) - 1
        '                    sTemp = sTemp + " " + s.Substring(letter, 1)
        '                Next
        '                strResponse += "<td class=""nametd"" width=""270""> " & sTemp & " </td>"
        '            End If
        '        Else
        '            strResponse += "<td width=""70"" height=""70"">&nbsp;</td>"
        '            strResponse += "<td width=""270"">&nbsp;</td>"
        '        End If
        'Next
        'strResponse += "</tr>"
        'Next
        'strResponse += "</table>"
        'Dim ltl As New Literal
        'ltl.Text = strResponse
        'phr.Controls.Add(ltl)
    End Sub
    'Protected Sub MakeStatusData(ByVal phr As PlaceHolder)
    '    db.Open()
    '    Const strSql As String = "SELECT TOP 27 P_NUM,NAME,STATUS FROM P_10 ORDER BY RANKID"
    '    Dim DR As SqlDataReader = New SqlCommand(strSql, db).ExecuteReader
    '    Dim dt = New DataTable
    '    If DR.HasRows Then
    '        dt.Load(DR)
    '    End If
    '    db.Close()
    '    Dim tbl As Table
    '    Dim RowCell As TableRow
    '    Dim Cell As TableCell
    '    Dim iRecCount As Integer

    '    tbl = New Table()
    '    'tbl.Width = Unit.Pixel(TableWidth)
    '    tbl.Width = Unit.Percentage(TableWidth)
    '    tbl.Height = TableHeight - 50

    '    tbl.BorderColor = Color.White
    '    tbl.BorderWidth = 1
    '    tbl.BorderStyle = BorderStyle.Solid
    '    tbl.CellSpacing = 0
    '    tbl.CellPadding = 0

    '    For i = 0 To iRowCount - 1
    '        RowCell = New TableRow()
    '        RowCell.Height = Unit.Pixel(70)
    '        For j = 0 To iColCount - 1
    '            Cell = New TableCell()
    '            Cell.Height = RowCell.Height
    '            Cell.BorderColor = Color.White
    '            Cell.BorderWidth = 1
    '            Cell.Width = Unit.Percentage(33)
    '            Cell.HorizontalAlign = HorizontalAlign.Justify
    '            Cell.VerticalAlign = VerticalAlign.Middle

    '            Dim tblManager As Table = New Table
    '            Dim rowManager As TableRow = New TableRow()
    '            iRecCount += 1

    '            Dim cellManager As TableCell
    '            If Not iRecCount > dt.Rows.Count Then
    '                cellManager = New TableCell()

    '                cellManager.CssClass = "imgtd"
    '                cellManager.Width = Unit.Pixel(50)
    '                cellManager.HorizontalAlign = HorizontalAlign.Center
    '                cellManager.VerticalAlign = VerticalAlign.Middle
    '                cellManager.BorderColor = Color.White
    '                cellManager.BorderWidth = 1
    '                cellManager.Height = RowCell.Height

    '                Dim img As New WebControls.Image
    '                img.Height = Unit.Pixel(40)
    '                img.Width = Unit.Pixel(40)
    '                If (dt.Rows(iRecCount - 1).Item("Status").ToString = 0) Then
    '                    img.ImageUrl = ResolveUrl("~/image/ManagerOffline.gif")
    '                Else
    '                    img.ImageUrl = ResolveUrl("~/image/OnlineCircle.png")
    '                End If
    '                cellManager.Controls.Add(img)
    '                rowManager.Cells.Add(cellManager)

    '                cellManager = New TableCell()

    '                cellManager.CssClass = "nametd"
    '                cellManager.Width = Unit.Percentage(100)
    '                'cellManager.HorizontalAlign = HorizontalAlign.Justify
    '                cellManager.Height = RowCell.Height

    '                Dim lblName = New Label()
    '                'lblName.Width = Unit.Percentage(100)
    '                Dim s As String = dt.Rows(iRecCount - 1).Item("Name").ToString.ToString()
    '                For letter = 0 To Len(s) - 1
    '                    lblName.Text = lblName.Text + " " + s.Substring(letter, 1)
    '                Next
    '                'lblName.Text = dt.Rows(iRecCount - 1).Item("Name").ToString.ToString()
    '                'Dim lblName = New Literal
    '                ''lblName.Width = Unit.Percentage(100)
    '                'lblName.Text = dt.Rows(iRecCount - 1).Item("Name").ToString.ToString()

    '                cellManager.Controls.Add(lblName)
    '            Else
    '                cellManager = New TableCell()

    '                cellManager.Text = "&nbsp;"
    '                cellManager.ColumnSpan = 1
    '                cellManager.CssClass = "imgtd"
    '            End If
    '            rowManager.Cells.Add(cellManager)
    '            tblManager.Rows.Add(rowManager)
    '            Cell.Controls.Add(tblManager)
    '            RowCell.Cells.Add(Cell)
    '        Next
    '        tbl.Rows.Add(RowCell)
    '    Next
    '    phr.Controls.Add(tbl)
    'End Sub
    ''' <summary>
    ''' 建立主官在營跑馬燈資料
    ''' </summary>
    ''' <param name="lbl"></param>
    ''' <remarks></remarks>
    Protected Sub MakeMarqueeMessage(ByVal lbl As Literal)

        Const strSql As String = "SELECT CONFIG_VALUE FROM SYSCONFIG WHERE CONFIG_VAR='P_10Marquee'"
        db.Open()
        Dim DR As SqlDataReader = New SqlCommand(strSql, db).ExecuteReader
        If DR.HasRows Then
            If DR.Read() Then
                Dim msg As StringBuilder = New StringBuilder()
                'msg.Append("<ul id=""marquee"" class=""marquee"">")
                Dim strTemp As String
                strTemp = DR("CONFIG_VALUE").ToString()
                Dim arrTemp As String() = strTemp.Replace(vbCrLf, ",").Split(CType(",", Char))

                For Each s As String In arrTemp
                    msg.Append("<li><b><font style=""font-family:華康中圓體;"">" & s & "</font></b></li>")
                Next

                'msg.Append("</ul>")
                lbl.Text = msg.ToString()
            End If
        End If
        db.Close()
    End Sub

    
#End Region

#Region "Form Function"
    ''**********************  以下為Form Function  ****************************
    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not IsPostBack Then
            MakeStatusData(phrStatus)
            MakeMarqueeMessage(lblMarqueeMsg)
        End If
        Page.Header.DataBind()
    End Sub

    Protected Sub Refresh_Tick(sender As Object, e As EventArgs) Handles Refresh.Tick
        MakeStatusData(phrStatus)
    End Sub
#End Region

End Class
