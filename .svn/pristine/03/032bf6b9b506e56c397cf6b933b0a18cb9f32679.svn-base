Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.Common
Imports System.Collections.Generic

Partial Class M_Source_08_MOA08006
    Inherits System.Web.UI.Page
    Dim user_id, org_uid As String
    Dim chk As New C_CheckFun
    Dim sql_function As New C_SQLFUN
    Dim CF As New CFlowSend
    Dim scripts As New StringBuilder

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        user_id = Session("user_id")
        org_uid = Session("ORG_UID")
        ErrMsg.Text = ""
        Try
            'session被清空回首頁
            If user_id = "" Or org_uid = "" Then
                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")
            Else
                '判斷登入者權限
                Dim LoginCheck As New C_Public
                If LoginCheck.LoginCheck(user_id, "MOA08006") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA08006.aspx")
                    Response.End()
                End If
                ViewState("printerManagerGuidId") = GetROLEGROUPGroupUid("影印機管理者") '取得影印機管理者群組ID
                ViewState("employeeIdList") = QueryROLEGROUPITEMByGroupUid(ViewState("printerManagerGuidId").ToString()) '取得影印機管理者群組人員ID名單
            End If
        Catch ex As Exception
            ErrMsg.Text = ex.Message
        End Try
    End Sub

    Sub On_Inserted(ByVal sender As Object, ByVal e As SqlDataSourceStatusEventArgs)
        ErrMsg.Text = "新增成功!!!"
        GV_Printer.DataBind()
    End Sub

    Protected Sub ImgInsert_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim bl_insertResult As Boolean = True
        Dim sErrMsg As String = String.Empty
        Try
            Dim OrgSel As DropDownList = Me.DetailsView1.FindControl("OrgSel")
            Dim UserSel As DropDownList = Me.DetailsView1.FindControl("UserSel")
            Dim Printer_No As TextBox = Me.DetailsView1.FindControl("tb_Printer_No")
            Dim Printer_Name As TextBox = Me.DetailsView1.FindControl("tb_Printer_Name")
            Dim memo As TextBox = Me.DetailsView1.FindControl("tb_memo")
            If UserSel.SelectedValue.ToString().Trim() = "" Then
                bl_insertResult = False
                sErrMsg = "請先選擇管理人員姓名~"
            End If
            chk.CheckDataLen(Printer_No.Text, 6, "新增時：<機器號碼>", True)
            chk.CheckDataLen(Printer_Name.Text, 50, "新增時：<機器名稱>", True)

            '判斷一組機器號碼不可有不同的機器名稱
            Dim dt_Printer As New DataTable("dtPrinter")
            Dim query As String = "SELECT [Printer_Name] FROM [P_0803] WHERE ([Printer_No] = @Printer_No) "
            Dim Adapter As SqlDataAdapter = New SqlDataAdapter(query, sql_function.G_conn_string)
            Adapter.SelectCommand.Parameters.Add(New SqlParameter("Printer_No", SqlDbType.VarChar)).Value = Printer_No.Text.Trim()
            Adapter.Fill(dt_Printer)
            If Not dt_Printer Is Nothing And dt_Printer.Rows.Count > 0 Then
                If Printer_Name.Text.Trim() <> dt_Printer.Rows(0)("Printer_Name").ToString() Then
                    bl_insertResult = False
                    sErrMsg = "同一個機器號碼已存在其設定的機器名稱，請重新確認再新增！"
                End If
            End If

            '判斷相同人員相同Printer是否已建立過關係
            If CheckExist(UserSel.SelectedValue, Printer_No.Text, 0) Then
                bl_insertResult = False
                sErrMsg = "相同人員同一台機器號碼只能建立乙次關係哦~"
            Else
                Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
                Dim btnInsert As ImageButton
                btnInsert = Me.DetailsView1.FindControl("ImgInsert")
                btnInsert.CommandName = "insert"

                Dim employeeIdList As List(Of String) = ViewState("employeeIdList") '影印機管理者群組內成員ID名單
                If Not employeeIdList.Contains(UserSel.SelectedValue) Then '若員工ID不在影印機管理者群組內
                    '將該員工新增至影印機管理者群組
                    Dim insertRoleGroupItemResult As String = insertROLEGROUPITEM(UserSel.SelectedValue, ViewState("printerManagerGuidId").ToString())
                    If Not insertRoleGroupItemResult.Equals("OK") Then
                        Throw New Exception(insertRoleGroupItemResult)
                    End If
                    employeeIdList.Add(UserSel.SelectedValue)
                    ViewState("employeeIdList") = employeeIdList
                End If
            End If
        Catch ex As Exception
            bl_insertResult = False
            sErrMsg = ex.Message
        End Try

        If bl_insertResult = False Then
            chk.AlertSussTranscation(sErrMsg, "MOA08006.aspx")
        End If
    End Sub

    '加入使用者所屬群組
    Private Function insertROLEGROUPITEM(ByVal employeeId As String, ByVal groupUid As String) As String
        Dim result As String = "OK"
        Try
            Dim connstr As String
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string
            Dim db As New SqlConnection(connstr)
            db.Open()

            Dim insCom As New SqlCommand("INSERT INTO ROLEGROUPITEM([Group_Uid],[employee_id]) VALUES(@Group_Uid,@employee_id)", db)
            insCom.Parameters.Add(New SqlParameter("@Group_Uid", SqlDbType.VarChar, 10)).Value = groupUid
            insCom.Parameters.Add(New SqlParameter("@employee_id", SqlDbType.VarChar, 10)).Value = employeeId
            insCom.ExecuteNonQuery()

            db.Close()
        Catch ex As Exception
            result = ex.Message
        End Try
        insertROLEGROUPITEM = result
    End Function

    '取得權限群組人員ID名單
    Private Function QueryROLEGROUPITEMByGroupUid(ByVal Group_Uid As String) As List(Of String)
        Dim employeeIdList As New List(Of String)
        Dim dt As DataTable = Nothing
        Dim connstr As String
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string
        Dim db As New SqlConnection(connstr)
        db.Open()

        Dim strSQL = "SELECT * FROM ROLEGROUPITEM WHERE Group_Uid = '" + Group_Uid + "'"
        Dim ds As New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter(strSQL, db)
        da.Fill(ds)
        dt = ds.Tables(0)

        db.Close()
        For Each dr As DataRow In dt.Rows
            employeeIdList.Add(dr("employee_id").ToString())
        Next
        QueryROLEGROUPITEMByGroupUid = employeeIdList
    End Function

    '取得權限群組Guid_Uid
    Private Function GetROLEGROUPGroupUid(ByVal Group_Name As String) As String
        Dim printerManagerGroupUid As String = ""
        Dim connstr As String
        Dim conn As New C_SQLFUN
        connstr = conn.G_conn_string
        Dim db As New SqlConnection(connstr)
        db.Open()

        Dim getIdCommand As New SqlCommand("SELECT Group_Uid FROM ROLEGROUP WHERE Group_Name = @Group_Name", db)
        getIdCommand.Parameters.Add(New SqlParameter("@Group_Name", SqlDbType.NVarChar, 30)).Value = Group_Name
        GetROLEGROUPGroupUid = getIdCommand.ExecuteScalar()

        db.Close()
    End Function

    Protected Sub GV_Printer_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GV_Printer.RowUpdating
        Dim ErrMsg As Label = Me.FindControl("ErrMsg")
        Dim lb_Guid_ID As Label = CType(GV_Printer.Rows(e.RowIndex).FindControl("lb_Guid_ID"), Label)
        Dim tb_NewPrinter_No As TextBox = CType(GV_Printer.Rows(e.RowIndex).FindControl("tb_Printer_No"), TextBox)
        Dim tb_Printer_Name As TextBox = CType(GV_Printer.Rows(e.RowIndex).FindControl("tb_PrinterName"), TextBox)
        Dim ddl_Employee As DropDownList = CType(GV_Printer.Rows(e.RowIndex).FindControl("ddl_Employee"), DropDownList)
        Dim ddl_ORG_UID As DropDownList = CType(GV_Printer.Rows(e.RowIndex).FindControl("ddl_ORG_UID"), DropDownList)
        Dim tb_memo As TextBox = CType(GV_Printer.Rows(e.RowIndex).FindControl("tb_memo"), TextBox)

        If tb_Printer_Name.Text.Trim().Length = 0 Or tb_NewPrinter_No.Text.Trim().Length = 0 Then
            ErrMsg.Text += "修改時：<機器名稱>或<機器號碼>長度不可空白"
            e.Cancel = True
        End If

        If tb_NewPrinter_No.Text.Trim().Length > 6 Then
            ErrMsg.Text += "修改時：<機器號碼>長度不可大於6碼"
            e.Cancel = True
        End If

        If tb_Printer_Name.Text.Trim().Length > 50 Then
            ErrMsg.Text += "修改時：<機器名稱>長度不可大於50個字元"
            e.Cancel = True
        End If

        If ddl_Employee.Items.Count = 0 Then
            ErrMsg.Text += "修改時：請選擇<管理人員>"
            e.Cancel = True
        End If

        '判斷一組機器號碼不可有不同的機器名稱
        Dim dt_Printer As New DataTable("dtPrinter")
        Dim query As String = "SELECT [Printer_Name], [employee_id] FROM [P_0803] WHERE ([Printer_No] = @Printer_No) "
        Dim Adapter As SqlDataAdapter = New SqlDataAdapter(query, sql_function.G_conn_string)
        Adapter.SelectCommand.Parameters.Add(New SqlParameter("Printer_No", SqlDbType.VarChar)).Value = tb_NewPrinter_No.Text.Trim()
        Adapter.Fill(dt_Printer)
        If Not dt_Printer Is Nothing And dt_Printer.Rows.Count > 0 Then
            If tb_Printer_Name.Text.Trim() <> dt_Printer.Rows(0)("Printer_Name").ToString() Then
                ErrMsg.Text += "同一個機器號碼已存在其設定的機器名稱，請重新確認再修改！~"
                e.Cancel = True
            End If
        End If

        If CheckExist(ddl_Employee.SelectedValue, tb_NewPrinter_No.Text, Integer.Parse(lb_Guid_ID.Text)) Then
            ErrMsg.Text += "修改時：相同人員同一台機器號碼只能建立乙次關係哦~"
            e.Cancel = True
        End If

        Dim employeeIdOld As String = dt_Printer.Rows(0)("employee_id").ToString() '修改前舊的人員ID
        '刪除修改前舊人員的影印機管理者權限,修改後新人員的權限增加,做在GV_Printer_RowDataBound
        Dim deleteResult As String = DeleteEmployeeRoleGroup(employeeIdOld, "影印機管理者")
        If Not deleteResult.Equals("OK") Then
            ErrMsg.Text += deleteResult
            e.Cancel = True
        End If

        e.NewValues.Add("ORG_UID", ddl_ORG_UID.SelectedValue)
        e.NewValues.Add("employee_id", ddl_Employee.SelectedValue)
    End Sub

    '刪除人員權限群組
    Private Function DeleteEmployeeRoleGroup(ByVal employee_id As String, ByVal Group_Name As String) As String
        Dim result As String = "OK"
        Try
            Dim connstr As String
            Dim conn As New C_SQLFUN
            connstr = conn.G_conn_string
            Dim db As New SqlConnection(connstr)
            db.Open()

            Dim getIdCommand As New SqlCommand("SELECT Group_Uid FROM ROLEGROUP WHERE Group_Name = @Group_Name", db)
            getIdCommand.Parameters.Add(New SqlParameter("@Group_Name", SqlDbType.NVarChar, 30)).Value = Group_Name
            Dim groupId = getIdCommand.ExecuteScalar()

            Dim deleteCommand As New SqlCommand("DELETE FROM ROLEGROUPITEM WHERE Group_Uid = @Group_Uid AND employee_id = @employee_id", db)
            deleteCommand.Parameters.Add(New SqlParameter("@Group_Uid", SqlDbType.NVarChar, 30)).Value = groupId
            deleteCommand.Parameters.Add(New SqlParameter("@employee_id", SqlDbType.NVarChar, 30)).Value = employee_id
            deleteCommand.ExecuteNonQuery()

            db.Close()
        Catch ex As Exception
            result = ex.Message
        End Try
        DeleteEmployeeRoleGroup = result
    End Function

    '顯示單位中文名稱
    Public Function ShowORG_Name(ByVal ORG_UID As String) As String
        Dim sResult As String = String.Empty
        If ORG_UID.Trim().Length <> 0 Then
            sResult = CF.getORG_Name(ORG_UID, sql_function.G_conn_string)
        End If
        Return sResult
    End Function

    '顯示單位人員中文名稱
    Public Function Showemp_chinese_name(ByVal employee_id As String) As String
        Dim sResult As String = String.Empty
        If employee_id.Trim().Length <> 0 Then
            sResult = CF.getEmp_chinese_name(employee_id, sql_function.G_conn_string)
        End If
        Return sResult
    End Function

    '判斷相同人員相同Printer是否已建立過關係,單一人員只能對應同一台Printer乙次 return: true=已對應過該Printer false=尚未對應過該Printer
    Private Function CheckExist(ByVal employee_id As String, ByVal Printer_No As String, ByVal Guid_ID As Integer) As Boolean
        Dim bl_Exist As Boolean = False
        Dim CountPrint_Num As Integer = 0
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))
        Dim sQuery As String = "select Guid_ID from dbo.P_0803 with(nolock) where Printer_No=@Printer_No and employee_id=@employee_id"
        If Guid_ID <> 0 Then
            sQuery += " and Guid_ID <> @Guid_ID"
        End If

        command.CommandType = CommandType.Text
        command.CommandText = sQuery
        command.Parameters.Add(New SqlParameter("Printer_No", SqlDbType.VarChar, 6)).Value = Printer_No
        command.Parameters.Add(New SqlParameter("employee_id", SqlDbType.VarChar, 10)).Value = employee_id
        If Guid_ID <> 0 Then
            command.Parameters.Add(New SqlParameter("Guid_ID", SqlDbType.Int)).Value = Guid_ID
        End If
        Try
            command.Connection.Open()
            Dim obj As Object = command.ExecuteScalar()
            If Not obj Is Nothing Then
                bl_Exist = True
            End If
        Catch ex As Exception
            bl_Exist = True
            ErrMsg.Text = ex.Message
        Finally
            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If
            command.Dispose()
            command = Nothing
        End Try

        Return bl_Exist
    End Function

    Protected Sub GV_Printer_RowEditing(sender As Object, e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GV_Printer.RowEditing
        GV_Printer.EditIndex = e.NewEditIndex
        Session("gvEditIndex") = e.NewEditIndex
    End Sub

    Protected Sub GV_Printer_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GV_Printer.RowDataBound
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            Dim lb_ORG_UID As Label = DirectCast(e.Row.FindControl("lb_ORG_UID"), Label)
            Dim lb_employee_id As Label = DirectCast(e.Row.FindControl("lb_employee_id"), Label)
            Dim ddl_Employee As DropDownList = DirectCast(e.Row.FindControl("ddl_Employee"), DropDownList)
            Try
                Dim ds_employee As New DataSet("Employees")
                Dim query As String = "SELECT [employee_id], [emp_chinese_name], [ORG_UID] FROM [EMPLOYEE] WHERE ([ORG_UID] = @ORG_UID) ORDER BY [emp_chinese_name]"
                Dim Adapter As SqlDataAdapter = New SqlDataAdapter(query, sql_function.G_conn_string)
                Adapter.SelectCommand.Parameters.Add(New SqlParameter("ORG_UID", SqlDbType.VarChar)).Value = lb_ORG_UID.Text
                Adapter.Fill(ds_employee)
                ddl_Employee.Items.Clear()
                ddl_Employee.DataSource = ds_employee
                ddl_Employee.DataTextField = "emp_chinese_name"
                ddl_Employee.DataValueField = "employee_id"
                ddl_Employee.DataBind()
                ddl_Employee.SelectedValue = lb_employee_id.Text

                Dim employeeIdList As List(Of String) = ViewState("employeeIdList") '影印機管理者群組內成員ID名單
                If Not employeeIdList.Contains(lb_employee_id.Text) Then '若此欄員工ID不在影印機管理者群組內
                    '將該員工新增至影印機管理者群組
                    Dim insertResult As String = insertROLEGROUPITEM(lb_employee_id.Text, ViewState("printerManagerGuidId").ToString())
                    If Not insertResult.Equals("OK") Then
                        Throw New Exception(insertResult)
                    End If
                    employeeIdList.Add(lb_employee_id.Text)
                    ViewState("employeeIdList") = employeeIdList
                End If
            Catch ex As Exception
                ErrMsg.Text = "(1)." + ex.Message()
            End Try
        End If
    End Sub

    Protected Sub ddl_ORG_UID_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        Dim lb_ORG_UID As Label = CType(GV_Printer.Rows(Session("gvEditIndex")).Cells(1).FindControl("lb_ORG_UID"), Label)
        Dim lb_employee_id As Label = CType(GV_Printer.Rows(Session("gvEditIndex")).Cells(1).FindControl("lb_employee_id"), Label)
        If CType(GV_Printer.Rows(Session("gvEditIndex")).Cells(1).FindControl("ddl_ORG_UID"), DropDownList).SelectedValue <> 0 Then
            lb_ORG_UID.Text = CType(GV_Printer.Rows(Session("gvEditIndex")).Cells(1).FindControl("ddl_ORG_UID"), DropDownList).SelectedValue
        End If
        Dim ddl_Employee As DropDownList = CType(GV_Printer.Rows(Session("gvEditIndex")).Cells(2).FindControl("ddl_Employee"), DropDownList)
        ddl_Employee.Items.Clear()
        Try
            Dim ds_employee As New DataSet("Employees")
            Dim query As String = "SELECT [employee_id], [emp_chinese_name], [ORG_UID] FROM [EMPLOYEE] WHERE ([ORG_UID] = @ORG_UID) ORDER BY [emp_chinese_name]"
            Dim Adapter As SqlDataAdapter = New SqlDataAdapter(query, sql_function.G_conn_string)
            Adapter.SelectCommand.Parameters.Add(New SqlParameter("ORG_UID", SqlDbType.VarChar)).Value = lb_ORG_UID.Text
            Adapter.Fill(ds_employee)
            ddl_Employee.DataSource = ds_employee
            ddl_Employee.DataTextField = "emp_chinese_name"
            ddl_Employee.DataValueField = "employee_id"
            ddl_Employee.DataBind()
        Catch ex As Exception
            ErrMsg.Text = "(2)." + ex.Message()
        End Try
    End Sub

End Class
