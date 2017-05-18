Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.IO
Partial Class M_Source_04_MOA04019_2

    Inherits System.Web.UI.Page
    Dim chk As New C_CheckFun


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If CType(Request.QueryString("shcode"), String) Is Nothing Or CType(Request.QueryString("usecheck"), String) Is Nothing Then
            Response.Write("<script language='javascript'>")
            Response.Write("alert('您的參數已遺失，請重新操作，感謝您！');")
            Response.Write("window.close();")
            Response.Write("</script>")
        Else
            Dim it_code As String = CType(Request.QueryString("shcode"), String)
            Dim usecheck As String = CType(Request.QueryString("usecheck"), String)
            If it_code.Length <> 6 And usecheck.Length <> 1 Then
                Response.Write("<script language='javascript'>")
                Response.Write("alert('您的參數不正確，請重新操作，感謝您！');")
                Response.Write("window.close();")
                Response.Write("</script>")
            End If

            lbl_it_code.Text = it_code
            lbl_usecheck.Text = usecheck
            Dim checkstatus As String
            Select Case usecheck
                Case "0"
                    checkstatus = "庫存"
                Case "1"
                    checkstatus = "待出庫"
                Case "2"
                    checkstatus = "出庫"
            End Select

            lblTitle.Text = "倉儲資料-物料詳細資料(" + checkstatus + ")"
        End If
    End Sub

    Public Function showeditdelLB(ByVal usecheck As String) As Boolean
        If usecheck = "2" Then
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub GridView1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If e.Row.FindControl("ddl_seat_num") IsNot Nothing Then
                Dim ddl_seat_num As DropDownList = CType(e.Row.FindControl("ddl_seat_num"), DropDownList)
                sqlds_Seat_Num.SelectCommand = "SELECT [seat_num]+'_'+[seat_name] as seatnum,[seat_num], [seat_name] FROM [P_0417] ORDER BY [seat_num]"
                ddl_seat_num.DataSource = sqlds_Seat_Num
                ddl_seat_num.DataTextField = "seatnum"
                ddl_seat_num.DataValueField = "seat_num"
                ddl_seat_num.DataBind()
                ddl_seat_num.SelectedValue = Convert.ToString(DataBinder.Eval(e.Row.DataItem, "seat_num"))
            End If
        End If
    End Sub

    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GridView1.RowUpdating
        '修改時欄位不可空白
        Dim ddl_seat_num As DropDownList = CType(GridView1.Rows(e.RowIndex).FindControl("ddl_seat_num"), DropDownList)
        Dim tb_expired_y As TextBox = CType(GridView1.Rows(e.RowIndex).FindControl("tb_expired_y"), TextBox)
        Dim tb_shcost As TextBox = CType(GridView1.Rows(e.RowIndex).FindControl("tb_shcost"), TextBox)
        Try
            If (tb_expired_y.Text.Trim().Length > 0) Then
                chk.CheckDataInt(tb_expired_y.Text, 1, "修改<有效期>時：", "必需大於0")
            End If
            If (tb_shcost.Text.Trim().Length > 0) Then
                chk.CheckDataInt(tb_shcost.Text, 1, "修改<物料價格>時：", "必需大於0")
            End If
            sqldatalist.UpdateParameters("Seat_num").DefaultValue = ddl_seat_num.SelectedValue
        Catch ex As Exception
            e.Cancel = True
            lbErrMsg.Visible = True
            lbErrMsg.Text = ex.Message
        End Try
    End Sub

    Protected Sub GridView1_RowUpdated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdatedEventArgs) Handles GridView1.RowUpdated
        lbErrMsg.Visible = False
        lbErrMsg.Text = ""
    End Sub

    Protected Sub GridView1_RowCancelingEdit(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCancelEditEventArgs) Handles GridView1.RowCancelingEdit
        lbErrMsg.Visible = False
        lbErrMsg.Text = ""
    End Sub
End Class
