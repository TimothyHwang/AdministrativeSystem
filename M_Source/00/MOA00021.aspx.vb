Imports System.Data.SqlClient
Imports System.Data

Partial Class Source_00_MOA00021
    Inherits System.Web.UI.Page
    Dim conn As New C_SQLFUN
    Dim connstr As String = conn.G_conn_string
    Dim db As New SqlConnection(connstr)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '判斷哪張表單
        Dim PagePathAll As String = ""
        PagePathAll = Request.QueryString("PagePathAll")

        '讀取eformsn
        Dim eformsn As String = ""
        eformsn = Request.QueryString("eformsn")

        '讀取eformid
        Dim eformid As String = ""
        eformid = Request.QueryString("eformid")

        '讀取read_only
        Dim read_only As String = ""
        read_only = Request.QueryString("read_only")

        PlaceHolder1.Controls.Add(New LiteralControl("<table cellSpacing='0' cellPadding='0' width='100%' border='0'>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<tr>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<td>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<table cellSpacing='0' cellPadding='0' border='0'>"))
        PlaceHolder1.Controls.Add(New LiteralControl("<tr>"))

        If eformid = "YAqBTxRP8P" Then '請假申請單需有上載附件功能
            PlaceHolder1.Controls.Add(New LiteralControl("<td style='height: 19px'><A onmouseover=TitleShow('1') href='../" & PagePathAll & ".aspx?eformsn=" & eformsn & "&eformid=" & eformid & "&read_only=" & read_only & "' target='mainFrame'><IMG height='19' src='../../image/topbn-01r.gif' width='122' border='0' name='Image1'></A></td>"))
            PlaceHolder1.Controls.Add(New LiteralControl("<td style='height: 19px'><A onmouseover=TitleShow('2') href='../00/MOA00105.aspx?eformsn=" & eformsn & "&eformid=" & eformid & "&read_only=" & read_only & "' target='mainFrame'><IMG height='19' src='../../image/topbn-02r.gif' width='122' border='0' name='Image2'></A></td>"))
            PlaceHolder1.Controls.Add(New LiteralControl("<td style='width: 123px;height: 19px'><A onmouseover=TitleShow('3') href='../00/MOA00005.aspx?eformid=" & eformid & "&eformrole=1&FlowDesignMode=1' target='mainFrame'><IMG height='19' src='../../image/topbn-03r.gif' width='122' border='0' name='Image3'></A></td>"))
        Else
            PlaceHolder1.Controls.Add(New LiteralControl("<td style='height: 19px'><A onmouseover=TitleShow('1') href='../" & PagePathAll & ".aspx?eformsn=" & eformsn & "&eformid=" & eformid & "&read_only=" & read_only & "' target='mainFrame'><IMG height='19' src='../../image/topbn-01r.gif' border='0' name='Image1'></A></td>"))
            PlaceHolder1.Controls.Add(New LiteralControl("<td style='height: 19px'><A onmouseover=TitleShow('3') href='../00/MOA00005.aspx?eformid=" & eformid & "&eformrole=1&FlowDesignMode=1' target='mainFrame'><IMG height='19' src='../../image/topbn-03r.gif' border='0' name='Image3'></A></td>"))
        End If

        PlaceHolder1.Controls.Add(New LiteralControl("</tr>"))
        PlaceHolder1.Controls.Add(New LiteralControl("</table>"))
        PlaceHolder1.Controls.Add(New LiteralControl("</td>"))
        PlaceHolder1.Controls.Add(New LiteralControl("</tr>"))
        PlaceHolder1.Controls.Add(New LiteralControl("</table>"))

        'javascript
        PlaceHolder1.Controls.Add(New LiteralControl("<script language='JavaScript' type='text/JavaScript'>"))

        '先顯示選擇表單

        PlaceHolder1.Controls.Add(New LiteralControl("MM_swapImage('Image1','','../../image/topbn-01.gif',1);"))

        PlaceHolder1.Controls.Add(New LiteralControl("function TitleShow(x) {"))

        PlaceHolder1.Controls.Add(New LiteralControl("if (x == '1'){"))
        PlaceHolder1.Controls.Add(New LiteralControl("MM_swapImage('Image1','','../../image/topbn-01.gif',1);"))
        PlaceHolder1.Controls.Add(New LiteralControl("MM_swapImage('Image2','','../../image/topbn-02r.gif',1);"))
        PlaceHolder1.Controls.Add(New LiteralControl("MM_swapImage('Image3','','../../image/topbn-03r.gif',1);"))
        PlaceHolder1.Controls.Add(New LiteralControl("}else "))

        If eformid = "YAqBTxRP8P" Then '請假申請單需有上載附件功能

            PlaceHolder1.Controls.Add(New LiteralControl("if (x == '2'){"))
            PlaceHolder1.Controls.Add(New LiteralControl("MM_swapImage('Image1','','../../image/topbn-01r.gif',1);"))
            PlaceHolder1.Controls.Add(New LiteralControl("MM_swapImage('Image2','','../../image/topbn-02.gif',1);"))
            PlaceHolder1.Controls.Add(New LiteralControl("MM_swapImage('Image3','','../../image/topbn-03r.gif',1);"))
            PlaceHolder1.Controls.Add(New LiteralControl("}else "))

            PlaceHolder1.Controls.Add(New LiteralControl("if (x == '3'){"))
            PlaceHolder1.Controls.Add(New LiteralControl("MM_swapImage('Image1','','../../image/topbn-01r.gif',1);"))
            PlaceHolder1.Controls.Add(New LiteralControl("MM_swapImage('Image2','','../../image/topbn-02r.gif',1);"))
            PlaceHolder1.Controls.Add(New LiteralControl("MM_swapImage('Image3','','../../image/topbn-03.gif',1);"))
            PlaceHolder1.Controls.Add(New LiteralControl("}}"))

        Else

            PlaceHolder1.Controls.Add(New LiteralControl("if (x == '3'){"))
            PlaceHolder1.Controls.Add(New LiteralControl("MM_swapImage('Image1','','../../image/topbn-01r.gif',1);"))
            PlaceHolder1.Controls.Add(New LiteralControl("MM_swapImage('Image3','','../../image/topbn-03.gif',1);"))
            PlaceHolder1.Controls.Add(New LiteralControl("}}"))

        End If

        PlaceHolder1.Controls.Add(New LiteralControl("</script>"))


    End Sub

    '取得表單種類ID
    Private Function GetEFormId(ByVal eformName As String) As String
        db.Open()
        Dim sqlcomm As New SqlCommand("SELECT eformid FROM EFORMS WHERE frm_chinese_name = '" + eformName + "'", db)
        GetEFormId = sqlcomm.ExecuteScalar()
        db.Close()
    End Function

End Class
