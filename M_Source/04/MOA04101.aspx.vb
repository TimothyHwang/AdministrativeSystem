Imports System.Data.SqlClient

Partial Class M_Source_04_MOA04101
    Inherits System.Web.UI.Page
    Public streformsn As String
    Public stritcode As String
    Public do_sql As New C_SQLFUN

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        If Trim(It_Name.Text) = "" Then
            Response.Redirect("../04/MOA04101.aspx?eformsn=" + streformsn + "&itcode=" + stritcode)
        Else
            SqlDataSource1.SelectCommand = "SELECT * FROM [P_0407] WHERE ([it_name] LIKE '%" + Trim(It_Name.Text) + "%')"
        End If
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType.Equals(DataControlRowType.DataRow) Then
            Dim lblIt_stock As Label
            Dim lblIt_applied As Label
            Dim lblIt_code As Label
            Dim txtIt_apply As TextBox
            lblIt_stock = e.Row.Cells(4).FindControl("lblIt_stock")
            lblIt_applied = e.Row.Cells(5).FindControl("lblIt_applied")
            txtIt_apply = e.Row.Cells(6).FindControl("txtIt_apply")
            lblIt_code = e.Row.Cells(0).FindControl("lblIt_code")
            lblIt_stock.Text = getstock(lblIt_code.Text)
            lblIt_applied.Text = getapplied(lblIt_code.Text)
            txtIt_apply.Attributes.Add("onblur", "return applyCheck(this,'" + lblIt_stock.Text + "','" + lblIt_applied.Text + "');")
        End If
    End Sub

    Protected Function getstock(ByVal _itcode As String) As Integer
        Dim conn As New C_SQLFUN
        Dim connstr As String = conn.G_conn_string
        Dim strSQL As String
        Dim dr As SqlDataReader
        Dim dt As SqlCommand
        Dim StockNum As Integer = 0

        '開啟連線
        Dim db As New SqlConnection(connstr)

        db.Open()
        strSQL = "select count(*) as StockNum from P_0414 " & _
                 "where shcode like '" + _itcode + "%' and " & _
                 "shtype='0' and UseCheck ='0'"
        dt = New SqlCommand(strSQL, db)
        dr = dt.ExecuteReader()

        If dr.Read() Then
            StockNum = dr.Item("StockNum")
        End If
        db.Close()
        Return StockNum
    End Function
    Protected Function getapplied(ByVal _itcode As String) As Integer
        Dim conn As New C_SQLFUN
        Dim connstr As String = conn.G_conn_string
        Dim strSQL As String
        Dim dr As SqlDataReader
        Dim dt As SqlCommand
        Dim AppliedNum As Integer = 0

        '開啟連線
        Dim db As New SqlConnection(connstr)

        db.Open()
        strSQL = "select count(*) as AppliedNum from P_0414 " & _
                 "where shcode like '" + _itcode + "%' and " & _
                 "shtype='0' and UseCheck ='1' and Job_Num='" + streformsn + "'"
        dt = New SqlCommand(strSQL, db)
        dr = dt.ExecuteReader()

        If dr.Read() Then
            AppliedNum = dr.Item("AppliedNum")
        End If
        db.Close()
        Return AppliedNum
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        streformsn = Request.QueryString("eformsn")
        stritcode = Request.QueryString("itcode")
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim N, CancelNum As Integer
        Dim tt As String
        Dim txtIt_apply As TextBox
        Dim lblIt_code As Label
        Dim stmt As String
        Dim user_id As String
        Dim appflag As Integer = 0

        '取得登入者帳號
        If Page.User.Identity.Name.ToString.IndexOf("\") > 0 Then

            Dim LoginAll As String = Page.User.Identity.Name.ToString

            Dim LoginID() As String = Split(LoginAll, "\")

            user_id = LoginID(1)
        Else
            user_id = Page.User.Identity.Name.ToString
        End If

        For N = 0 To GridView1.Rows.Count - 1

            lblIt_code = GridView1.Rows(N).Cells(0).FindControl("lblIt_code")
            tt = lblIt_code.Text
            txtIt_apply = GridView1.Rows(N).Cells(6).FindControl("txtIt_apply")
            If Trim(txtIt_apply.Text) <> "" Then
                '更新工單編號及改變為待出庫
                If Convert.ToInt16(Trim(txtIt_apply.Text)) > 0 Then
                    stmt = "update P_0414 set Job_Num='" + streformsn + "',UseCheck ='1',AppDate=getdate(),AppKeyIn='" + user_id + "'" & _
                            " where shcode in " & _
                            "(select top " + Trim(txtIt_apply.Text) + " shcode from P_0414 where shcode like '" + lblIt_code.Text + "%' and shtype='0' and UseCheck ='0')"

                    If do_sql.db_exec(stmt, do_sql.G_conn_string) = False Then
                        Exit Sub
                    End If
                    appflag = 1
                End If
                '輸入負為退回多申請
                If Convert.ToInt16(Trim(txtIt_apply.Text)) < 0 Then
                    CancelNum = Convert.ToInt16(Trim(txtIt_apply.Text)) * -1
                    stmt = "update P_0414 set Job_Num='',UseCheck ='0',AppDate=null,AppKeyIn=''" & _
                           " where shcode in " & _
                           "(select top " + CancelNum.ToString + " shcode from P_0414 where shcode like '" + lblIt_code.Text + "%' and shtype='0' and UseCheck ='1' and Job_Num='" + streformsn + "')"

                    If do_sql.db_exec(stmt, do_sql.G_conn_string) = False Then
                        Exit Sub
                    End If
                End If
            End If
        Next

        If appflag = 1 Then
            stmt = "update P_0415 set nAppStockDate=getdate(),nAppStockStatus='N' where eformsn='" + streformsn + "'"
            If do_sql.db_exec(stmt, do_sql.G_conn_string) = False Then
                Exit Sub
            End If
        End If

        '重新整理頁面
        Response.Write("<script language='javascript'>")
        Response.Write(" window.dialogArguments.location='../04/MOA04100.aspx?eformsn=" + streformsn + "&eformid=F9MBD7O97G&read_only=2';")
        Response.Write("window.close();")
        Response.Write("</script>")
    End Sub

End Class
