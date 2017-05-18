Imports System.Data
Imports System.Data.SqlClient
Namespace M_Source._00
    Partial Class M_Source_00_MOA00015
        Inherits Page        

        Dim user_id As String
        Dim SendVal As String
        Dim eformid, employee_id, eformsn, eformrole, signer As String

        Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
            Try

                If Session("user_id") = "" Then
                    user_id = Page.User.Identity.Name.Substring(Page.User.Identity.Name.IndexOf("\", StringComparison.Ordinal) + 1)
                Else
                    user_id = Session("user_id").ToString()
                End If
                SendVal = Request.QueryString("SendVal")
                If SendVal.Length > 0 Then
                    Dim arr As String() = Split(SendVal, ",")
                    eformid = arr(0)
                    employee_id = arr(1)
                    eformsn = arr(2)
                    eformrole = arr(3)
                    signer = arr(4)
                End If
                If Not IsPostBack Then
                    ''設定下拉選項
                    Dim tool As New C_Public
                    Dim arrSelectionUser As String() = Split(tool.GetUnitIDsByStepName(user_id, "單位資訊官", ",", "", 3), ",")
                    Dim dtSource As New DataTable
                    Dim col1 As New DataColumn("employee_id")
                    Dim col2 As New DataColumn("emp_chinese_name")
                    col1.DataType = Type.GetType("System.String")
                    col2.DataType = Type.GetType("System.String")
                    dtSource.Columns.Add(col1)
                    dtSource.Columns.Add(col2)
                    For Each s As String In arrSelectionUser
                        Dim row As DataRow = dtSource.NewRow()
                        row(col1) = s
                        row(col2) = tool.GetUserNameByID(s)
                        dtSource.Rows.Add(row)
                    Next
                    DDLUser.DataSource = dtSource
                    DDLUser.DataBind()
                End If
            Catch ex As Exception
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('" + ex.Message + "');")
                Response.Write(" </script>")
            End Try
        End Sub

        Protected Sub btnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click
            Dim trans As SqlTransaction = Nothing
            Dim departmentCode As String = GetDepartmentCode(DDLUser.SelectedValue)
            Try
                Dim strSql As String = ""
                Dim DR As SqlDataReader
                Dim DC As New SQLDBControl


                strSql = "SELECT TOP(1) * FROM flowctl WHERE  eformsn = '" + eformsn + "' ORDER BY flowsn DESC"
                DR = DC.CreateReader(strSql)

                ''已有流程                
                If DR.Read() Then
                    'strSql += "INSERT INTO flowctl(eformid,eformrole,eformsn,stepsid,steps,empuid,emp_chinese_name"
                    'strSql += ",group_name,gonogo,nextstep,goback,nextuser,orgname,important,backto_steps,recdate"
                    'strSql += ",is_testmode,duty_for,appdate,filler,subtype,deptcode,createdate) "
                    'strSql += "VALUES("
                    'strSql += "'" + DR("eformid").ToString() + "'"
                    'strSql += ",'" + DR("eformrole").ToString() + "'"
                    'strSql += ",'" + DR("eformsn").ToString() + "'"
                    'strSql += ",'" + DR("stepsid").ToString() + "'"
                    'strSql += ",'" + DR("steps").ToString() + "'"
                    'strSql += ",'" + DDLUser.SelectedValue.Trim() + "'"
                    'strSql += ",'" + DDLUser.SelectedItem.Text.Trim() + "'"
                    'strSql += ",'" + DR("group_name").ToString() + "'"
                    'strSql += ",'" + DR("'?'").ToString() + "'"
                    'strSql += ",'" + DR("nextstep").ToString() + "'"
                    'strSql += ",'" + DR("goback").ToString() + "'"
                    'strSql += ",'" + DR("nextuser").ToString() + "'"
                    'strSql += ",'" + DR("orgname").ToString() + "'"
                    'strSql += ",'" + DR("important").ToString() + "'"
                    'strSql += ",'" + DR("backto_steps").ToString() + "'"
                    'strSql += ",'" + DR("GETDATE()").ToString() + "'"
                    'strSql += ",'" + DR("is_testmode").ToString() + "'"
                    'strSql += ",'" + DR("duty_for").ToString() + "'"
                    'strSql += ",GETDATE()"
                    'strSql += ",'" + DR("filler").ToString() + "'"
                    'strSql += ",'" + DR("subtype").ToString() + "'"
                    'strSql += ",'" + departmentCode + "'"
                    'strSql += ",GETDATE()"

                    'DC.TransStart()
                    'DC.ExecuteTransSQL(strSql)

                    'strSql = "UPDATE flowctl SET gonogo = 'T' WHERE flowsn = '" + DR("flowsn").ToString() + "'"
                    'DC.ExecuteTransSQL(strSql)

                    'DC.TransCommit()

                    'Response.Write(" <script language='javascript'>")
                    'Response.Write(" alert('表單已呈轉給單位資訊官');")
                    ''重新整理頁面
                    'Response.Write(" window.dialogArguments.location='../00/MOA00010.aspx';")
                    'Response.Write(" window.close();")
                    'Response.Write(" </script>")
                Else ''新表單
                    Dim SendValOA = eformid & "," & user_id & "," & eformsn & "," & "1" & "," & DDLUser.SelectedValue
                    Dim Val_P As String
                    Dim FC As New CFlowSend
                    Dim do_sql As New C_SQLFUN

                    '表單審核
                    Val_P = FC.F_Send(SendValOA, do_sql.G_conn_string).ToString()
                    Response.Redirect("../00/MOA00007.aspx?val=" & Val_P & "&PageUp=New")
                    do_sql.G_errmsg = "存檔成功"
                End If
                DC.Dispose()
            Catch ex As Exception

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('" + ex.Message + "');")
                Response.Write(" </script>")
            End Try
        End Sub

        Private Function GetDepartmentCode(ByVal employeeID As String) As String
            Dim db As New SqlConnection(New C_SQLFUN().G_conn_string)
            Dim departmentCode As String = ""

            db.Open()
            Dim comm As New SqlCommand("SELECT ORG_UID FROM EMPLOYEE WHERE employee_id = @employee_id", db)
            comm.Parameters.Add("@employee_id", SqlDbType.VarChar, 10).Value = employeeID.Trim()
            departmentCode = comm.ExecuteScalar().ToString()
            comm.Dispose()
            db.Close()

            Return departmentCode
        End Function
    End Class
End Namespace