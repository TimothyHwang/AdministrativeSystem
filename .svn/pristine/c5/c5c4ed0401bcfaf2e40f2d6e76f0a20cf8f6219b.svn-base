Imports System.Data.SqlClient

Partial Class M_Source_10_MOA10005
    Inherits Page

    ReadOnly sql_function As New C_SQLFUN
    ReadOnly connstr As String = sql_function.G_conn_string
    Dim db As New SqlConnection(connstr)

#Region "Custom Function"
    ''**********************  以下為Custom Function  **************************
#End Region

#Region "Form Function"
    ''**********************  以下為  Form Function  **************************
    Protected Sub gvManagerEmployee_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gvManagerEmployee.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                Dim lblManager As Label = CType(e.Row.FindControl("lblManager"), Label)
                Dim ibnOnline As ImageButton = CType(e.Row.FindControl("ibnOnStatus"), ImageButton)
                Dim ibnOffline As ImageButton = CType(e.Row.FindControl("ibnOffStatus"), ImageButton)
                Dim lblStatus As Label = CType(e.Row.FindControl("lblStatus"), Label)

                If lblManager Is Nothing Then Return

                If lblStatus IsNot Nothing Then
                    ibnOnline.Visible = lblStatus.Text = "1"
                    ibnOffline.Visible = lblStatus.Text = "0"
                End If
        End Select
    End Sub

    Protected Sub gvManagerEmployee_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles gvManagerEmployee.RowCommand
        Dim pk_index As Integer
        Dim pk As Integer
        Dim strSql As String
        Select Case e.CommandName
            Case "Add"
            Case "Offline"
                pk_index = CInt(e.CommandArgument)
                Int32.TryParse(gvManagerEmployee.DataKeys(pk_index).Value.ToString(), pk)
                db.Open()
                strSql = "UPDATE P_10 SET STATUS=1 WHERE P_NUM=" & pk
                Call New SqlCommand(strSql, db).ExecuteNonQuery()
                db.Close()
            Case "Online"
                pk_index = CInt(e.CommandArgument)
                Int32.TryParse(gvManagerEmployee.DataKeys(pk_index).Value.ToString(), pk)
                db.Open()
                strSql = "UPDATE P_10 SET STATUS=0 WHERE P_NUM=" & pk
                Call New SqlCommand(strSql, db).ExecuteNonQuery()
                db.Close()
        End Select
        gvManagerEmployee.DataSourceID = Nothing
        gvManagerEmployee.DataSourceID = SqlDataSourceP_10.ID
        gvManagerEmployee.DataBind()

    End Sub


    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Page.Header.DataBind()
    End Sub

#End Region


End Class
