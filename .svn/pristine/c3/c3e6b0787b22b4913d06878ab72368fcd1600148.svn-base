Imports System.Data
Imports System.Data.SqlClient

Partial Class M_Source_00_MOA00014
    Inherits System.Web.UI.Page

    Dim user_id, eformsn, connstr As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
            Try
                Dim conn As New C_SQLFUN
                connstr = conn.G_conn_string

                If Session("user_id") = "" Then
                    user_id = Page.User.Identity.Name.Substring(Page.User.Identity.Name.IndexOf("\") + 1)
                Else
                    user_id = Session("user_id")
                End If
                eformsn = Request.QueryString("eformsn")

            Catch ex As Exception
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('" + ex.Message + "');")
                Response.Write(" </script>")
            End Try
    End Sub

    Protected Sub btnSend_Click(sender As Object, e As System.EventArgs) Handles btnSend.Click
        Dim trans As SqlTransaction = Nothing
        Dim departmentCode As String = GetDepartmentCode(DDLUser.SelectedValue)
        Try
            Dim db As New SqlConnection(connstr)
            Dim ds As New DataSet()

            db.Open()
            trans = db.BeginTransaction()
            Dim selectComm As New SqlCommand("SELECT TOP(1) * FROM flowctl WHERE  eformsn = @eformsn ORDER BY flowsn DESC", db, trans)
            selectComm.Parameters.Add("@eformsn", Data.SqlDbType.VarChar, 16).Value = eformsn.Trim()
            Dim da As New SqlDataAdapter(selectComm)
            da.Fill(ds)
            Dim dt As DataTable = ds.Tables(0)

            Dim strInsert As String = "INSERT INTO flowctl(eformid,eformrole,eformsn,stepsid,steps,empuid,emp_chinese_name"
            strInsert = strInsert + ",group_name,gonogo,nextstep,goback,nextuser,orgname,important,backto_steps,recdate"
            strInsert = strInsert + ",is_testmode,duty_for,appdate,filler,subtype,deptcode,createdate) "
            strInsert = strInsert + "VALUES(@eformid,@eformrole,@eformsn,@stepsid,@steps,@empuid,@emp_chinese_name"
            strInsert = strInsert + ",@group_name,'?',@nextstep,@goback,@nextuser,@orgname,@important,@backto_steps"
            strInsert = strInsert + ",GETDATE(),@is_testmode,@duty_for,GETDATE(),@filler,@subtype,@deptcode,GETDATE())"
            Dim insertComm As New SqlCommand(strInsert, db, trans)
            insertComm.Parameters.Add("@eformid", Data.SqlDbType.VarChar, 10).Value = dt.Rows(0)("eformid")
            insertComm.Parameters.Add("@eformrole", Data.SqlDbType.Int).Value = dt.Rows(0)("eformrole")
            insertComm.Parameters.Add("@eformsn", Data.SqlDbType.VarChar, 16).Value = dt.Rows(0)("eformsn")
            insertComm.Parameters.Add("@stepsid", Data.SqlDbType.Int).Value = dt.Rows(0)("stepsid")
            insertComm.Parameters.Add("@steps", Data.SqlDbType.Int).Value = dt.Rows(0)("steps")
            insertComm.Parameters.Add("@empuid", Data.SqlDbType.VarChar, 10).Value = DDLUser.SelectedValue.Trim()
            insertComm.Parameters.Add("@emp_chinese_name", Data.SqlDbType.VarChar, 50).Value = DDLUser.SelectedItem.Text.Trim()
            insertComm.Parameters.Add("@group_name", Data.SqlDbType.VarChar, 50).Value = dt.Rows(0)("group_name")
            insertComm.Parameters.Add("@nextstep", Data.SqlDbType.Int).Value = dt.Rows(0)("nextstep")
            insertComm.Parameters.Add("@goback", Data.SqlDbType.Int).Value = dt.Rows(0)("goback")
            insertComm.Parameters.Add("@nextuser", Data.SqlDbType.VarChar, 10).Value = dt.Rows(0)("nextuser")
            insertComm.Parameters.Add("@orgname", Data.SqlDbType.VarChar, 50).Value = dt.Rows(0)("orgname")
            insertComm.Parameters.Add("@important", Data.SqlDbType.VarChar, 1).Value = dt.Rows(0)("important")
            insertComm.Parameters.Add("@backto_steps", Data.SqlDbType.Int).Value = dt.Rows(0)("backto_steps")
            insertComm.Parameters.Add("@is_testmode", Data.SqlDbType.Char, 1).Value = dt.Rows(0)("is_testmode")
            insertComm.Parameters.Add("@duty_for", Data.SqlDbType.VarChar, 10).Value = dt.Rows(0)("duty_for")
            insertComm.Parameters.Add("@filler", Data.SqlDbType.VarChar, 10).Value = dt.Rows(0)("filler")
            insertComm.Parameters.Add("@subtype", Data.SqlDbType.VarChar, 4).Value = dt.Rows(0)("subtype")
            insertComm.Parameters.Add("@deptcode", Data.SqlDbType.VarChar, 10).Value = departmentCode
            insertComm.ExecuteNonQuery()

            Dim updateComm As New SqlCommand("UPDATE flowctl SET gonogo = 'T' WHERE flowsn = @flowsn", db, trans)
            updateComm.Parameters.Add("@flowsn", Data.SqlDbType.Decimal).Value = dt.Rows(0)("flowsn")
            updateComm.ExecuteNonQuery()

            trans.Commit()
            db.Close()
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('表單已呈轉給上一級主管');")
            '重新整理頁面
            Response.Write(" window.dialogArguments.location='../00/MOA00010.aspx';")
            Response.Write(" window.close();")
            Response.Write(" </script>")

        Catch ex As Exception
            trans.Rollback()
            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('" + ex.Message + "');")
            Response.Write(" </script>")
        End Try
    End Sub

    Private Function GetDepartmentCode(ByVal employeeID As String) As String
        Dim db As New SqlConnection(connstr)
        Dim departmentCode As String = ""

        db.Open()
        Dim comm As New SqlCommand("SELECT ORG_UID FROM EMPLOYEE WHERE employee_id = @employee_id", db)
        comm.Parameters.Add("@employee_id", Data.SqlDbType.VarChar, 10).Value = employeeID.Trim()
        departmentCode = comm.ExecuteScalar()
        comm.Dispose()
        db.Close()

        Return departmentCode
    End Function

End Class
