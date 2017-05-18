Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration
Imports System.Collections.Generic

Public Class C_SQLFUN
    Public G_table As New DataTable
    Public G_usr_table As New DataTable
    'Public G_conn_string As String = "server=BURSRV;uid=sa;pwd=sa2005;database=mndps1"
    Public G_conn_string As String = ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString '"Data Source=bursrv;Initial Catalog=MOA;Persist Security Info=True;User ID=sysadm;Password=sysadm"
    Public G_Trans As System.Data.SqlClient.SqlTransaction
    Public G_Conn As New System.Data.SqlClient.SqlConnection
    Public G_errmsg As String = ""  '錯誤訊息
    Public G_user_id As String = ""  '使用時須G_user_id = session("user_id")

    '查詢SQL 
    'stmt為 sql查詢字串
    'ConnectionString為連接資料庫字串
    Public Function db_sql(ByVal stmt As String, ByVal ConnectionString As String) As Boolean
        db_sql = False
        G_errmsg = ""
        On Error GoTo error_db_sql
        Dim conn As SqlConnection
        Dim comm As New SqlCommand
        Dim dy As New DataSet
        conn = New SqlConnection(ConnectionString)

        Dim adapter As SqlDataAdapter = New SqlDataAdapter(stmt, conn)
        adapter.Fill(dy)
        G_table = dy.Tables(0)
        conn.Close()
        db_sql = True
        Exit Function
error_db_sql:
        G_errmsg = "sql errormsg1:" + Err.Description
        G_errmsg = G_errmsg.Replace(Chr(13), "")
        G_errmsg = G_errmsg.Replace(Chr(10), "")
        G_errmsg = G_errmsg.Replace("'", "")
        'MsgBox("sql errormsg1:" + Err.Description, MsgBoxStyle.OkOnly, "")
        conn.Close()
    End Function
    '查詢SQL 
    'stmt為 sql查詢字串
    'Trans為tranctions
    Public Function db_sql(ByVal stmt As String, ByVal Trans As System.Data.SqlClient.SqlTransaction) As Boolean
        db_sql = False
        G_errmsg = ""
        On Error GoTo error_db_sql
        Dim Conn As SqlConnection = Trans.Connection
        Dim comm As New SqlCommand
        Dim dy As New DataSet

        comm.Connection = Conn
        comm.CommandType = CommandType.Text
        comm.Transaction = Trans
        comm.CommandText = stmt
        Dim adapter As SqlDataAdapter = New SqlDataAdapter(comm)
        adapter.Fill(dy)
        G_table = dy.Tables(0)
        db_sql = True
        Exit Function
error_db_sql:
        G_errmsg = "sql errormsg2:" + Err.Description
        G_errmsg = G_errmsg.Replace(Chr(13), "")
        G_errmsg = G_errmsg.Replace(Chr(10), "")
        G_errmsg = G_errmsg.Replace("'", "")
        'MsgBox("sql errormsg2:" + Err.Description, MsgBoxStyle.OkOnly, "")
    End Function
    '執行insert,update,delete 等 
    'stmt為 insert,update,delete 之SQL字串
    'Trans為tranctions
    Public Function db_exec(ByVal stmt As String, ByVal Trans As System.Data.SqlClient.SqlTransaction) As Boolean
        On Error GoTo error_db_exec
        db_exec = False
        G_errmsg = ""
        Dim conn As SqlConnection = Trans.Connection
        Dim comm As New SqlCommand

        comm.Connection = conn
        comm.CommandType = CommandType.Text
        comm.CommandText = stmt
        comm.Transaction = Trans
        comm.ExecuteNonQuery()
        db_exec = True
        Exit Function
error_db_exec:
        G_errmsg = "sql errormsg3:" + Err.Description
        G_errmsg = G_errmsg.Replace(Chr(13), "")
        G_errmsg = G_errmsg.Replace(Chr(10), "")
        G_errmsg = G_errmsg.Replace("'", "")
        'MsgBox("sql errormsg3:" + Err.Description, MsgBoxStyle.OkOnly, "")
        'conn.Close()
        Exit Function
    End Function
    '執行insert,update,delete 等 
    'stmt為 insert,update,delete 之SQL字串
    'ConnectionString為連接資料庫字串
    Public Function db_exec(ByVal stmt As String, ByVal ConnectionString As String) As Boolean
        On Error GoTo error_db_exec
        db_exec = False
        G_errmsg = ""
        Dim conn As SqlConnection
        Dim comm As New SqlCommand
        conn = New SqlConnection(ConnectionString)
        comm.Connection = conn
        comm.CommandType = CommandType.Text
        comm.CommandText = stmt
        comm.Connection.Open()
        comm.ExecuteNonQuery()
        conn.Close()
        db_exec = True
        Exit Function
error_db_exec:
        G_errmsg = "sql errormsg4:" + Err.Description
        G_errmsg = G_errmsg.Replace(Chr(13), "")
        G_errmsg = G_errmsg.Replace(Chr(10), "")
        G_errmsg = G_errmsg.Replace("'", "")
        'MsgBox("sql errormsg4:" + Err.Description, MsgBoxStyle.OkOnly, "")
        conn.Close()
        Exit Function
    End Function
    '傳欲用的connection字傳,使建立Transaction G_Trans
    'ConnectionString為連接資料庫字串
    Function begin_tran(ByVal ConnectionString As String) As Boolean
        On Error GoTo err_begin_tran
        begin_tran = False
        G_errmsg = ""
        G_Conn.ConnectionString = ConnectionString
        G_Conn.Open()
        G_Trans = G_Conn.BeginTransaction(IsolationLevel.ReadCommitted)
        begin_tran = True
        Exit Function
err_begin_tran:
        G_errmsg = "begin_tran errormsg5:" + Err.Description
        G_errmsg = G_errmsg.Replace(Chr(13), "")
        G_errmsg = G_errmsg.Replace(Chr(10), "")
        G_errmsg = G_errmsg.Replace("'", "")
        'MsgBox("begin_tran errormsg5:" + Err.Description, MsgBoxStyle.OkOnly, "")
        Exit Function
    End Function
    '對Transaction G_Trans做commit, 並關畢connection
    Sub commit_tran()
        Try
            G_Trans.Commit()
            G_Conn.Close()
        Catch ex As Exception
            G_errmsg = "commit_tran errormsg6:" + ex.Message
            G_errmsg = G_errmsg.Replace(Chr(13), "")
            G_errmsg = G_errmsg.Replace(Chr(10), "")
            G_errmsg = G_errmsg.Replace("'", "")
            'MsgBox("commit_tran errormsg6:" + ex.Message)
        End Try
    End Sub
    '對Transaction G_Trans做Rollback, 並關畢connection
    Sub rollback_tran()
        Try
            G_Trans.Rollback()
            G_Conn.Close()
        Catch ex As Exception
            G_errmsg = "rollback msg7:" + ex.Message
            G_errmsg = G_errmsg.Replace(Chr(13), "")
            G_errmsg = G_errmsg.Replace(Chr(10), "")
            G_errmsg = G_errmsg.Replace("'", "")
            'MsgBox("rollback msg7:" + ex.Message)
        End Try
    End Sub
    '找人員資訊
    Function select_urname(ByVal user_id As String) As Boolean
        Dim stmt As String
        select_urname = False
        'stmt = "select A.*,B.ORG_NAME,C.title_name from (Employee A left outer join AdminGroup B on A.ORG_UID=B.ORG_UID) "
        'stmt += " left outer join Titles C on A.title_id=C.title_id "
        'V_EmpInfo
        stmt = "select * from V_EmpInfo"
        stmt += " where employee_id ='" + user_id + "'"
        If db_sql(stmt, G_conn_string) = False Then
            Exit Function
        End If
        G_usr_table = G_table
        If G_usr_table.Rows.Count = 0 Then
            Exit Function
        End If
        select_urname = True
    End Function
    '檢查身份證字號
    Function checkIDNO(ByVal ID As String) As Boolean
        Dim X1 As Integer
        Dim X2 As Integer
        Dim num_Total As Integer
        Dim G_N_J As Integer
        Dim G_N_I As Integer

        checkIDNO = False
        ID = UCase(Trim(ID))
        If Len(ID) < 10 Then
            Exit Function
        End If
        Select Case Mid$(ID, 1, 1)
            Case "A", "B", "C", "D", "E", "F", "G", "H", "J", "K"
                X1 = 1
            Case "L", "M", "N", "P", "Q", "R", "S", "T", "U", "V"
                X1 = 2
            Case "X", "Y", "W", "Z", "I", "O"
                X1 = 3
            Case Else
                Exit Function
        End Select
        Select Case Mid$(ID, 1, 1)
            Case "A", "L", "X"
                X2 = 0
            Case "B", "M", "Y"
                X2 = 1
            Case "C", "N", "W"
                X2 = 2
            Case "D", "P", "Z"
                X2 = 3
            Case "E", "Q", "I"
                X2 = 4
            Case "F", "R", "O"
                X2 = 5
            Case "G", "S"
                X2 = 6
            Case "H", "T"
                X2 = 7
            Case "J", "U"
                X2 = 8
            Case "K", "V"
                X2 = 9
        End Select
        If IsNumeric(Mid$(ID, 2, 10)) = False Then
            Exit Function
        End If
        If Mid(ID, 2, 1) <> "1" And Mid(ID, 2, 1) <> "2" Then
            Exit Function
        End If
        num_Total = X1 + 9 * X2
        G_N_J = 8
        For G_N_I = 2 To 9
            num_Total = num_Total + Val(Mid(ID, G_N_I, 1)) * G_N_J
            G_N_J = G_N_J - 1
        Next
        num_Total = num_Total + Val(Mid(ID, 10, 1))
        If (num_Total Mod 10) = 0 Then
            checkIDNO = True
        End If
    End Function

    Sub set_drop_down(ByVal drop_box As DropDownList, ByVal value As String)
        Dim box_index As Integer = drop_box.Items.Count
        Dim ci As Integer
        For ci = 0 To box_index - 1
            If drop_box.Items(ci).Value = value Then
                drop_box.SelectedIndex = ci
                Exit Sub
            End If
        Next
    End Sub

    Sub inc_file(ByVal read_file_name As String, ByVal write_file_name As String, ByVal F_file_name As String)
        Call inc_file(read_file_name, write_file_name, F_file_name, False)
    End Sub

    Sub inc_file(ByVal read_file_name As String, ByVal write_file_name As String, ByVal F_file_name As String, ByVal append_flag As Boolean)
        Dim str_line As String
        FileOpen(1, read_file_name, OpenMode.Input, OpenAccess.Read)
        If append_flag = True Then
            FileOpen(2, write_file_name, OpenMode.Append, OpenAccess.Write)
        Else
            FileOpen(2, write_file_name, OpenMode.Output, OpenAccess.Write)
        End If

        PrintLine(2, "/dtd " + F_file_name + ".txt")
        Do While EOF(1) = False
            str_line = LineInput(1)
            PrintLine(2, str_line)
        Loop

        FileClose(1)
        PrintLine(2, "/dtd " + F_file_name + ".txt")
        FileClose(2)

    End Sub

    Sub print_sdata(ByVal write_file_name As String, ByVal f_data As String, ByVal f_data2 As String)
        FileOpen(2, write_file_name, OpenMode.Append, OpenAccess.Write)
        PrintLine(2, f_data)
        If f_data2 <> "" Then
            PrintLine(2, f_data2)
        End If
        FileClose(2)
    End Sub

    Sub print_block(ByVal write_file_name As String, ByVal block_name As String, ByVal x1 As String, ByVal y1 As String, ByVal block_value As String)
        FileOpen(2, write_file_name, OpenMode.Append, OpenAccess.Write)
        PrintLine(2, "/add " + block_name + " " + x1 + " " + y1)
        PrintLine(2, block_value)
        PrintLine(2, "/add " + block_name + " " + x1 + " " + y1)
        FileClose(2)
    End Sub
    'write_file_name 輸出檔
    'block_name 套印變數名
    'x1 座標x
    'y1 座標y
    '欲印之值
    'x2 離原點x座標
    'y2 離原點y座標
    '印幾次
    Sub print_block(ByVal write_file_name As String, ByVal block_name As String, ByVal x1 As String, ByVal y1 As String, ByVal block_value As String, ByVal x2 As String, ByVal y2 As String, ByVal x_count As Integer)
        Dim di As Integer
        Dim x_long As Integer = CInt(x1)
        Dim y_long As Integer = CInt(y1)
        FileOpen(2, write_file_name, OpenMode.Append, OpenAccess.Write)

        For di = 1 To x_count
            PrintLine(2, "/add " + block_name + " " + x_long.ToString + " " + y_long.ToString)
            PrintLine(2, block_value)
            PrintLine(2, "/add " + block_name + " " + x_long.ToString + " " + y_long.ToString)
            x_long += CInt(x2)
            y_long += CInt(y2)
        Next
        FileClose(2)
    End Sub
    '前補零
    'value_data 欲補字串
    'value_len  總長度
    Function add_zero(ByVal value_data As String, ByVal value_len As Integer) As String
        add_zero = value_data
        Dim tmp_for As String
        If value_data.Length < value_len Then
            tmp_for = ""
            For pi As Integer = 1 To value_len - value_data.Length
                tmp_for += "0"
            Next
            add_zero = tmp_for + value_data
        End If

    End Function
End Class

Public Class SQLDBControl
    Implements IDisposable

    Public ConnStr As String
    Private Conn As SqlConnection
    Private TransCMD As SqlTransaction

    Public Function CreateFunctionReader(ByVal sFunctionName As String, ByVal ht As Hashtable) As SqlDataReader
        Conn.ConnectionString = ConnStr
        Conn.Open()
        Dim cmd As New SqlCommand()
        cmd.CommandText = sFunctionName
        cmd.CommandType = CommandType.Text
        cmd.Connection = Conn
        ''cmd.CommandText = "SELECT " + sFunctionName + " ("
        For Each de As DictionaryEntry In ht
            cmd.Parameters.Add(New SqlParameter("@" + de.Key.ToString(), de.Value.ToString()))
        Next
        Try
            Return cmd.ExecuteReader()
        Catch ex As SqlException
            Throw New System.Exception("An exception has occurred." + ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' SQL Connect 類別
    ''' </summary>
    ''' <param name="Server"></param>
    Public Sub New(ByVal Server As String)
        Conn = New SqlConnection
        If (Server.ToUpper() = "PIFC") Then

        Else
            ConnStr = WebConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString
        End If
    End Sub

    Public Sub New()
        Conn = New SqlConnection
        ConnStr = WebConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString
    End Sub

    Public Function CreateDataTable(ByVal SQL As String) As DataTable
        Dim dt As New DataTable()
        Dim DR As SqlDataReader
        Dim Cmd As New SqlCommand()
        Conn.ConnectionString = ConnStr
        Conn.Open()
        Cmd.CommandText = SQL
        Cmd.Connection = Conn
        Try
            DR = Cmd.ExecuteReader()
            dt.Load(DR)
            Return dt
        Catch
            Throw New System.Exception("An exception has occurred.")
        End Try
    End Function

    ''' <summary>
    ''' 將查詢結果丟至DataReader
    ''' </summary>
    ''' <param name="SQL"></param>
    ''' <returns></returns>
    Public Function CreateReader(ByVal SQL As String) As SqlDataReader
        Conn.ConnectionString = ConnStr
        Conn.Open()
        Dim Cmd As New SqlCommand()
        Cmd.CommandText = SQL
        Cmd.Connection = Conn
        Try
            Return Cmd.ExecuteReader()
        Catch EX As Exception
            Throw New System.Exception("An exception has occurred." + EX.Message)
        End Try
    End Function

    ''' <summary>
    ''' 將查詢結果丟至DataReader
    ''' </summary>
    ''' <param name="SQL"></param>
    ''' <returns></returns>
    Public Function CreateTransReader(ByVal SQL As String) As SqlDataReader
        
        Dim Cmd As New SqlCommand()
        Cmd.CommandText = SQL
        Cmd.Connection = Conn
        Try
            Return Cmd.ExecuteReader()
        Catch EX As Exception
            Throw New Exception("An exception has occurred." + EX.Message)
        End Try
    End Function

    Public Function CreateAdapter(ByVal SQL As String) As SqlDataAdapter
        Conn.ConnectionString = ConnStr
        Conn.Open()
        Dim cmd As New SqlCommand()
        cmd.CommandText = SQL
        cmd.Connection = Conn
        Dim ada As New SqlDataAdapter(cmd)
        Try
            Return ada
        Catch ex As Exception
            Throw New System.Exception("An exception has occurred." + ex.Message)
        End Try
    End Function
    ''' <summary>
    ''' 帶參數HashTable型的DataReader
    ''' </summary>
    ''' <param name="SQL"></param>
    ''' <param name="ht"></param>
    ''' <returns></returns>
    Public Function CreateReaderWithParameters(ByVal SQL As String, ByVal ht As Hashtable) As SqlDataReader

        Conn.ConnectionString = ConnStr
        Conn.Open()
        Dim Cmd As New SqlCommand()
        Cmd.CommandText = SQL
        Cmd.CommandType = CommandType.Text

        For Each de As DictionaryEntry In ht
            Cmd.Parameters.Add(New SqlParameter("@" + de.Key.ToString(), de.Value.ToString()))
        Next
        Cmd.Connection = Conn
        Try
            Return Cmd.ExecuteReader()
        Catch ex As Exception
            Throw New System.Exception("An exception has occurred." + ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' 單執行SQL Command(Update、Insert)無回傳值
    ''' </summary>
    ''' <param name="SQL"></param>
    ''' <returns></returns>
    Public Function ExecuteSQL(ByVal SQL As String) As Boolean
        Dim boolReturn As Boolean = False
        Conn.ConnectionString = ConnStr
        Conn.Open()
        Dim Cmd As New SqlCommand()
        Cmd.CommandText = SQL
        Cmd.Connection = Conn
        Dim Trans As SqlTransaction = Conn.BeginTransaction(IsolationLevel.ReadCommitted)
        Cmd.Transaction = Trans
        Try
            Cmd.ExecuteNonQuery()
            Trans.Commit()
            boolReturn = True
        Catch ex As SqlException
            MessageBox.Show(ex.ToString())
            Trans.Rollback()
            boolReturn = False
        Finally
            Cmd.Dispose()
            Conn.Close()
        End Try
        Return boolReturn
    End Function

    ''' <summary>
    ''' 單執行SQL Command(Update、Insert)無回傳值
    ''' </summary>
    ''' <param name="SQL"></param>
    ''' <returns></returns>
    Public Function ExecuteTransSQL(ByVal SQL As String) As Boolean
        Dim boolReturn As Boolean = False
        Dim Cmd As New SqlCommand()
        Cmd.CommandText = SQL
        Cmd.Connection = Conn
        'Dim Trans As SqlTransaction = Conn.BeginTransaction(IsolationLevel.ReadCommitted)
        Cmd.Transaction = TransCMD
        Try
            Cmd.ExecuteNonQuery()
            TransCMD.Commit()
            boolReturn = True
        Catch ex As SqlException
            MessageBox.Show(ex.ToString())
            TransCMD.Rollback()
            boolReturn = False
        Finally
            Cmd.Dispose()
            Conn.Close()
        End Try
        Return boolReturn
    End Function
    Public Sub TransStart()
        Conn.ConnectionString = ConnStr
        Conn.Open()
        TransCMD = Conn.BeginTransaction(IsolationLevel.ReadCommitted)
    End Sub

    Public Sub TransCommit()
        TransCMD.Commit()
        Conn.Close()
    End Sub

    ''' <summary>
    ''' 使用Hash Table執行Insert SQL Command(Update、Insert)無回傳值
    ''' </summary>
    ''' <param name="TableName"></param>
    ''' <param name="value"></param>
    ''' <returns></returns>
    Public Function ExcuteInsertSQL(ByVal TableName As String, ByVal value As Hashtable) As Boolean

        'Conn.ConnectionString = ConnStr
        'Conn.Open()
        'Dim Cmd As New SqlCommand()
        '''建立交易transaction
        'Dim tran As SqlTransaction = Conn.BeginTransaction()
        'Dim strSql As String = "INSERT INTO " + TableName + "("
        'For i As Integer = 0 To value.Count - 1
        '    strSql = strSql + value.Keys
        'Next
        'strSql += ") VALUES ("
        'For i As Integer = 0 To value.Count - 1
        '    strSql = strSql + value.Values
        'Next
        'strSql += ")"
        'Cmd.CommandText = strSql
        'Cmd.Connection = Conn
        'Cmd.Transaction = tran
        'Try
        '    Cmd.ExecuteNonQuery()
        '    Cmd.Dispose()
        '    Conn.Close()
        '    Return True
        'Catch
        '    Cmd.Dispose()
        '    Conn.Close()
        '    Return False
        'End Try
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If

            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下面的 Finalize()。
            ' TODO: 將大型欄位設定為 null。
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: 只有當上面的 Dispose(ByVal disposing As Boolean) 有可釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
    'Protected Overrides Sub Finalize()
    '    ' 請勿變更此程式碼。在上面的 Dispose(ByVal disposing As Boolean) 中輸入清除程式碼。
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' 由 Visual Basic 新增此程式碼以正確實作可處置的模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更此程式碼。在以上的 Dispose 置入清除程式碼 (ByVal 視為布林值處置)。
        If (Conn.State = System.Data.ConnectionState.Open) Then
            Conn.Close()
        End If
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class

''' <summary>
''' 簽核流程紀錄
''' </summary>
''' <remarks></remarks>
Public Class flowctl
    Inherits SQLDBControl
    Implements IDisposable

    Private _flowsn As Int32
    Private _eformid As String
    Private _eformrole As Int32
    Private _eformsn As String
    Private _stepsid As Int32
    Private _steps As Int32
    Private _empuid As String
    Private _emp_chinese_name As String
    Private _group_name As String
    Private _hddate As DateTime
    Private _gonogo As String
    Private _comment As String
    Private _nextstep As Int32
    Private _goback As Int32
    Private _nextuser As String
    Private _orgname As String
    Private _important As String
    Private _backto_steps As Int32
    Private _recdate As DateTime
    Private _is_testmode As Boolean
    Private _duty_for As String
    Private _appdate As DateTime
    Private _filler As String
    Private _subtype As String
    Private _deptcode As String
    Private _createdate As DateTime
    Private _presentSignFlowsn As Int32

    ''' <summary>
    ''' 流程編號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property flowsn() As Int32
        Get
            Return _flowsn
        End Get
        Set(ByVal value As Int32)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 表單種類編號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property eformid() As String
        Get
            Return _eformid
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 關卡類別
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property eformrole() As Int32
        Get
            Return _eformrole
        End Get
        Set(ByVal value As Int32)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 表單編號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property eformsn() As String
        Get
            Return _eformsn
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 流程編號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property stepsid() As Int32
        Get
            Return _stepsid
        End Get
        Set(ByVal value As Int32)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 步驟編號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property steps() As Int32
        Get
            Return _steps
        End Get
        Set(ByVal value As Int32)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 簽核人ID編號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property empuid() As String
        Get
            Return _empuid
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 建立日期
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property createdate() As DateTime
        Get
            Return _createdate
        End Get
        Set(ByVal value As DateTime)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 簽核人中文名稱
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property emp_chinese_name As String
        Get
            Return _emp_chinese_name
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 簽核人群組
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property group_name As String
        Get
            Return _group_name
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 簽核時間
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property hddate As DateTime
        Get
            Return _hddate
        End Get
        Set(ByVal value As DateTime)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 簽核狀態
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property gonogo As String
        Get
            Return _gonogo
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 簽核意見
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property comment As String
        Get
            Return _comment
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 流程下一步驟編號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property nextstep As Int32
        Get
            Return _nextstep
        End Get
        Set(ByVal value As Int32)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 跳回步驟編號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property goback As Int32
        Get
            Return _goback
        End Get
        Set(ByVal value As Int32)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 流程下一步驟簽核人ID
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property nextuser As String
        Get
            Return _nextuser
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property orgname As String
        Get
            Return _orgname
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property important As String
        Get
            Return _important
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 跳回步驟關數
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property backto_steps As Int32
        Get
            Return _backto_steps
        End Get
        Set(ByVal value As Int32)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 流程紀錄時間
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property recdate As DateTime
        Get
            Return _recdate
        End Get
        Set(ByVal value As DateTime)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 是否為測試流程關卡
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property is_testmode As Boolean
        Get
            Return _is_testmode
        End Get
        Set(ByVal value As Boolean)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property duty_for As String
        Get
            Return _duty_for
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 流程到達時間
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property appdate As DateTime
        Get
            Return _appdate
        End Get
        Set(ByVal value As DateTime)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property filler As String
        Get
            Return _filler
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 次類別
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property subtype As String
        Get
            Return _subtype
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property
    ''' <summary>
    ''' 簽核人組織編號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property deptcode As String
        Get
            Return _deptcode
        End Get
        Set(ByVal value As String)
            Return
        End Set
    End Property

    'Public Sub New(ByVal Server As String)
    '    MyBase.New(Server)
    'End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sEFormid">表單種類編號</param>
    ''' <param name="sEFormsn">表單編號</param>
    ''' <param name="sEmpID">簽核人ID</param>
    ''' <param name="sGonogo">簽核狀態</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal sEFormid As String, ByVal sEFormsn As String, ByVal sEmpID As String, ByVal sGonogo As String)
        MyBase.New()
        Load(sEFormid, sEFormsn, sEmpID, sGonogo)
    End Sub

    Public Sub New(ByVal sEFormid As String, ByVal sEFormsn As String, ByVal sGonogo As String)
        MyBase.New()
        Load(sEFormid, sEFormsn, sGonogo)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sFlowsn">流程編號</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal sFlowsn As Int32)
        MyBase.New()
        Load(sFlowsn)
    End Sub

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub Load(ByVal sEFormid As String, ByVal sEFormsn As String, ByVal sGonogo As String)
        Dim strSql As String
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader

        strSql = "SELECT * FROM FLOWCTL WHERE EFORMID='" + sEFormid + "' AND EFORMSN='" + sEFormsn + "' AND GONOGO='" + sGonogo + "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read() Then
                _flowsn = CType(IIf(IsDBNull(DR("flowsn")), 0, DR("flowsn")), Int32)
                _eformid = CType(IIf(IsDBNull(DR("eformid")), "", DR("eformid")), String)
                _eformrole = CType(IIf(IsDBNull(DR("eformrole")), 0, DR("eformrole")), Int32)
                _eformsn = CType(IIf(IsDBNull(DR("eformsn")), "", DR("eformsn")), String)
                _stepsid = CType(DR("stepsid").ToString(), Int32)
                _steps = CType(IIf(IsDBNull(DR("steps")), 0, DR("steps")), Int32)
                _empuid = CType(IIf(IsDBNull(DR("empuid")), "", DR("empuid")), String)
                _emp_chinese_name = CType(IIf(IsDBNull(DR("emp_chinese_name")), "", DR("emp_chinese_name")), String)
                _group_name = CType(IIf(IsDBNull(DR("group_name")), "", DR("group_name")), String)
                If Not IsDBNull(DR("hddate")) Then _hddate = CType(DR("hddate").ToString(), DateTime)
                _gonogo = CType(IIf(IsDBNull(DR("gonogo")), "", DR("gonogo")), String)
                _comment = CType(IIf(IsDBNull(DR("comment")), "", DR("comment")), String)
                _nextstep = CType(DR("nextstep").ToString(), Int32)
                _goback = CType(IIf(IsDBNull(DR("goback")), 0, DR("goback")), Int32)
                _nextuser = CType(IIf(IsDBNull(DR("nextuser")), "", DR("nextuser")), String)
                _orgname = CType(IIf(IsDBNull(DR("orgname")), "", DR("orgname")), String)
                _important = CType(IIf(IsDBNull(DR("important")), "", DR("important")), String)
                _backto_steps = CType(IIf(IsDBNull(DR("backto_steps")), 0, DR("backto_steps")), Int32)
                If Not IsDBNull(DR("recdate")) Then _recdate = CType(DR("recdate").ToString, DateTime)
                _is_testmode = CType(IIf(IsDBNull(DR("is_testmode")), True, DR("is_testmode")), Boolean)
                _duty_for = CType(IIf(IsDBNull(DR("duty_for")), "", DR("duty_for")), String)
                If Not IsDBNull(DR("appdate")) Then _appdate = CType(DR("appdate").ToString(), DateTime)
                _filler = DR("filler").ToString()
                _subtype = DR("subtype").ToString()
                _deptcode = DR("deptcode").ToString()
                If Not IsDBNull(DR("createdate")) Then _createdate = CType(DR("createdate"), DateTime)
            End If
        End If
        DC.Dispose()
    End Sub

    Public Sub Load(ByVal sEFormid As String, ByVal sEFormsn As String, ByVal sEmpID As String, ByVal sGonogo As String)
        Dim strSql As String
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader

        strSql = "SELECT * FROM FLOWCTL WHERE EFORMID='" + sEFormid + "' AND EFORMSN='" + sEFormsn + "' AND EMPUID='" + sEmpID + "' AND GONOGO='" + sGonogo + "'"
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read() Then
                _flowsn = CType(IIf(IsDBNull(DR("flowsn")), 0, DR("flowsn")), Int32)
                _eformid = CType(IIf(IsDBNull(DR("eformid")), "", DR("eformid")), String)
                _eformrole = CType(IIf(IsDBNull(DR("eformrole")), 0, DR("eformrole")), Int32)
                _eformsn = CType(IIf(IsDBNull(DR("eformsn")), "", DR("eformsn")), String)
                _stepsid = CType(DR("stepsid").ToString(), Int32)
                _steps = CType(IIf(IsDBNull(DR("steps")), 0, DR("steps")), Int32)
                _empuid = CType(IIf(IsDBNull(DR("empuid")), "", DR("empuid")), String)
                _emp_chinese_name = CType(IIf(IsDBNull(DR("emp_chinese_name")), "", DR("emp_chinese_name")), String)
                _group_name = CType(IIf(IsDBNull(DR("group_name")), "", DR("group_name")), String)
                If Not IsDBNull(DR("hddate")) Then _hddate = CType(DR("hddate").ToString(), DateTime)
                _gonogo = CType(IIf(IsDBNull(DR("gonogo")), "", DR("gonogo")), String)
                _comment = CType(IIf(IsDBNull(DR("comment")), "", DR("comment")), String)
                _nextstep = CType(DR("nextstep").ToString(), Int32)
                _goback = CType(IIf(IsDBNull(DR("goback")), 0, DR("goback")), Int32)
                _nextuser = CType(IIf(IsDBNull(DR("nextuser")), "", DR("nextuser")), String)
                _orgname = CType(IIf(IsDBNull(DR("orgname")), "", DR("orgname")), String)
                _important = CType(IIf(IsDBNull(DR("important")), "", DR("important")), String)
                _backto_steps = CType(IIf(IsDBNull(DR("backto_steps")), 0, DR("backto_steps")), Int32)
                If Not IsDBNull(DR("recdate")) Then _recdate = CType(DR("recdate").ToString, DateTime)
                _is_testmode = CType(IIf(IsDBNull(DR("is_testmode")), True, DR("is_testmode")), Boolean)
                _duty_for = CType(IIf(IsDBNull(DR("duty_for")), "", DR("duty_for")), String)
                If Not IsDBNull(DR("appdate")) Then _appdate = CType(DR("appdate").ToString(), DateTime)
                _filler = DR("filler").ToString()
                _subtype = DR("subtype").ToString()
                _deptcode = DR("deptcode").ToString()
                If Not IsDBNull(DR("createdate")) Then _createdate = CType(DR("createdate"), DateTime)
            End If
        End If
        DC.Dispose()
    End Sub

    Public Sub Load(ByVal sFlowsn As Int32)
        Dim strSql As String
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader

        strSql = "SELECT * FROM FLOWCTL WHERE FLOWSN=" + sFlowsn.ToString()
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read() Then
                _flowsn = CType(IIf(IsDBNull(DR("flowsn")), 0, DR("flowsn")), Int32)
                _eformid = CType(IIf(IsDBNull(DR("eformid")), "", DR("eformid")), String)
                _eformrole = CType(IIf(IsDBNull(DR("eformrole")), 0, DR("eformrole")), Int32)
                _eformsn = CType(IIf(IsDBNull(DR("eformsn")), "", DR("eformsn")), String)
                _stepsid = CType(DR("stepsid").ToString(), Int32)
                _steps = CType(IIf(IsDBNull(DR("steps")), 0, DR("steps")), Int32)
                _empuid = CType(IIf(IsDBNull(DR("empuid")), "", DR("empuid")), String)
                _emp_chinese_name = CType(IIf(IsDBNull(DR("emp_chinese_name")), "", DR("emp_chinese_name")), String)
                _group_name = CType(IIf(IsDBNull(DR("group_name")), "", DR("group_name")), String)
                If Not IsDBNull(DR("hddate")) Then _hddate = CType(DR("hddate").ToString(), DateTime)
                _gonogo = CType(IIf(IsDBNull(DR("gonogo")), "", DR("gonogo")), String)
                _comment = CType(IIf(IsDBNull(DR("comment")), "", DR("comment")), String)
                _nextstep = CType(DR("nextstep").ToString(), Int32)
                _goback = CType(IIf(IsDBNull(DR("goback")), 0, DR("goback")), Int32)
                _nextuser = CType(IIf(IsDBNull(DR("nextuser")), "", DR("nextuser")), String)
                _orgname = CType(IIf(IsDBNull(DR("orgname")), "", DR("orgname")), String)
                _important = CType(IIf(IsDBNull(DR("important")), "", DR("important")), String)
                _backto_steps = CType(IIf(IsDBNull(DR("backto_steps")), 0, DR("backto_steps")), Int32)
                If Not IsDBNull(DR("recdate")) Then _recdate = CType(DR("recdate").ToString, DateTime)
                _is_testmode = CType(IIf(IsDBNull(DR("is_testmode")), True, DR("is_testmode")), Boolean)
                _duty_for = CType(IIf(IsDBNull(DR("duty_for")), "", DR("duty_for")), String)
                If Not IsDBNull(DR("appdate")) Then _appdate = CType(DR("appdate").ToString(), DateTime)
                _filler = DR("filler").ToString()
                _subtype = DR("subtype").ToString()
                _deptcode = DR("deptcode").ToString()
                If Not IsDBNull(DR("createdate")) Then _createdate = CType(DR("createdate"), DateTime)
            End If
        End If
        DC.Dispose()
    End Sub

    Public Function PreStep(ByVal sFlowsn As Int32) As flowctl
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String

        strSql = "SELECT * FROM FLOWCTL WHERE FLOWSN=" + sFlowsn.ToString()
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read() Then
                strSql = "SELECT * FROM FLOWCTL WHERE EFORMSN='" + DR("EFORMSN").ToString() + "' AND EFORMID='" + DR("EFORMID").ToString() + "' AND NEXTSTEP ='" + _stepsid.ToString() + "' ORDER BY FLOWSN DESC"
                Dim DC2 As New SQLDBControl
                Dim DR2 As SqlDataReader
                DR2 = DC2.CreateReader(strSql)
                If DR2.HasRows Then
                    If DR2.Read() Then
                        Return New flowctl(CType(DR2("FLOWSN").ToString(), Int32))
                    End If
                End If
                DC2.Dispose()
            End If
        End If
        DC.Dispose()
        Return New flowctl()
    End Function

    Public Function PreStep() As flowctl
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String

        strSql = "SELECT * FROM FLOWCTL WHERE FLOWSN=" + _flowsn.ToString()
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read() Then
                strSql = "SELECT * FROM FLOWCTL WHERE EFORMSN='" + DR("EFORMSN").ToString() + "' AND EFORMID='" + DR("EFORMID").ToString() + "' AND NEXTSTEP ='" + _stepsid.ToString() + "' ORDER BY FLOWSN DESC"
                Dim DC2 As New SQLDBControl
                Dim DR2 As SqlDataReader
                DR2 = DC2.CreateReader(strSql)
                If DR2.HasRows Then
                    If DR2.Read() Then
                        Return New flowctl(CType(DR2("FLOWSN").ToString(), Int32))
                    End If
                End If
                DC2.Dispose()
            End If
        End If
        DC.Dispose()
        Return New flowctl()
    End Function

    ''' <summary>
    ''' 回上一關
    ''' </summary>
    ''' <remarks></remarks>
    Public Function JunpBack() As String
        Dim flowctl As New CFlowSend
        JunpBack = flowctl.F_BackM(_eformid.ToString() + "," + _empuid + "," + _eformsn + "," + _eformrole.ToString(), ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString, 1)
    End Function

    ''' <summary>
    ''' 回上n關
    ''' </summary>
    ''' <param name="iSteps">關卡數</param>
    ''' <remarks></remarks>
    Public Sub JumpBack(ByVal iSteps As Int32)
        Dim flowctl As New CFlowSend
        flowctl.F_BackM(_eformid.ToString() + "," + _empuid + "," + _eformsn + "," + _eformrole.ToString(), ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString, iSteps)
    End Sub

    ''' <summary>
    ''' 簽核
    ''' </summary>
    ''' <param name="sStatus">簽核的狀態</param>
    ''' <remarks></remarks>
    Public Sub Sign(ByVal sStatus As String)
        Dim sResult As Boolean
        Dim DC As New SQLDBControl
        Dim strSql As String = "UPDATE FLOWCTL SET HDDATE=GETDATE(),GONOGO='" + sStatus + "' WHERE EFORMSN='" + _eformsn + "' AND EMPUID='" + _empuid + "'AND HDDATE IS NULL"
        sResult = DC.ExecuteSQL(strSql)
        DC.Dispose()
        If Not sResult Then
            MessageBox.Show("簽核失敗，請洽管理員")
        End If
    End Sub
    ''' <summary>
    ''' 回到特定關卡名稱關卡
    ''' </summary>
    ''' <param name="sGroupName"></param>
    ''' <remarks></remarks>
    Public Function JumpByGroupName(ByVal sGroupName As String) As String
        Try
            JumpByGroupName = ""            
            '變更成表單狀態
            Sign("2")

            '找出特定關卡名稱關卡
            Dim DC As New SQLDBControl
            Dim DR As SqlDataReader
            Dim strSql As String

            ''如果流程中有相同的關卡群組名稱，取離簽核最近紀錄的特定關卡群組名稱
            strSql = "SELECT * FROM FLOWCTL WHERE EFORMID='" + _eformid + "' AND EFORMSN='" + _eformsn + "' AND EFORMROLE='1' AND GROUP_NAME='" + sGroupName + "' ORDER BY FLOWSN DESC"
            DR = DC.CreateReader(strSql)
            If DR.HasRows Then
                If DR.Read() Then
                    ''資訊報修管制單位流程
                    Dim flowStep As New flowctl(CType(DR("flowsn"), Int32))
                    Dim DC2 As New SQLDBControl
                    '新增FlowCTL資訊報修管制單位流程                    
                    strSql = "INSERT INTO FLOWCTL(EFORMID,EFORMROLE,EFORMSN,STEPS,EMPUID,EMP_CHINESE_NAME,GROUP_NAME,GONOGO,NEXTSTEP,GOBACK,NEXTUSER,STEPSID,ORGNAME,IMPORTANT,RECDATE,IS_TESTMODE,APPDATE,DEPTCODE,CREATEDATE) VALUES ("
                    strSql += "'" + flowStep.eformid + "'"
                    strSql += "," + flowStep.eformrole.ToString()
                    strSql += ",'" + flowStep.eformsn + "'"
                    strSql += "," + flowStep.steps.ToString()
                    strSql += ",'" + flowStep.empuid + "'"
                    strSql += ",'" + flowStep.emp_chinese_name + "'"
                    strSql += ",'" + flowStep.group_name + "'"
                    strSql += ",'?'"
                    strSql += "," + flowStep.nextstep.ToString()
                    strSql += "," + flowStep.goback.ToString()
                    strSql += ",'" + flowStep.nextuser + "'"
                    strSql += "," + flowStep.stepsid.ToString()
                    strSql += ",'" + flowStep.orgname + "'"
                    strSql += ",'" + flowStep.important + "'"
                    strSql += ",'" + flowStep.recdate.ToString("yyyy/MM/dd HH:MM:ss") + "'"
                    strSql += "," + IIf(flowStep.is_testmode, 1, 0).ToString()
                    strSql += ",'" + flowStep.appdate.ToString("yyyy/MM/dd HH:MM:ss") + "'"
                    strSql += ",'" + flowStep.deptcode + "'"
                    strSql += ",GETDATE()"
                    strSql += ")"
                    DC2.ExecuteSQL(strSql)
                    DC2.Dispose()

                    '回傳值給如影隨行
                    'eformid,eformsn,gonogo,stepsid,goto_steps,sShowMsg,eformrole
                    JumpByGroupName = eformid & "," & eformsn & "," & "1," & stepsid & "," & flowStep.stepsid & "," & "表單成功送給-" & flowStep.emp_chinese_name & "," & eformrole

                End If
            End If
            DC.Dispose()            
        Catch ex As Exception
            JumpByGroupName = ""
        End Try
    End Function

    ''' <summary>
    ''' 由關卡群組取得流程中曾簽核的人員資料
    ''' </summary>
    ''' <param name="StepGroupName">關卡名稱</param>
    ''' <param name="StepCount">第幾關</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SignerWithGroupInFlow(ByVal StepGroupName As String, ByVal StepCount As Integer) As BasicEmployee          
        Dim empReturn As New BasicEmployee
        Dim strSql As String = "SELECT * FROM FLOWCTL WHERE EFORMSN='" + eformsn + "' ORDER BY FLOWSN ASC"
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            While DR.Read()
                Dim flow As New flow(DR("stepsid").ToString(), eformid)

                Dim Flows As New List(Of KeyValuePair(Of Integer, String))
                flow.EFormFlows(eformid, 1, Flows)
                If Flows.Count > 0 Then
                    For Each Pair As KeyValuePair(Of Integer, String) In Flows
                        Dim key As String = Pair.Key
                        Dim value As Integer = Pair.Value
                        If DR("stepsid") = value And StepCount = key Then
                            empReturn = New BasicEmployee(DR("empuid"))
                            Exit While
                        End If
                    Next
                End If
                flow.Dispose()
            End While
        End If
        DC.Dispose()
        Return empReturn
    End Function

End Class

''' <summary>
''' 表單流程
''' </summary>
''' <remarks></remarks>
Public Class flow
    Inherits SQLDBControl
    Implements IDisposable    

    Private _eformid As String
    Private _eformrole As Int32
    Private _stepsid As Int32
    Private _steps As Int32
    Private _nextstep As Int32
    Private _major_step As String
    Private _pors As String
    Private _x As Int32
    Private _y As Int32
    Private _group_id As String
    Private _group_type As String
    Private _allhandle As String
    Private _canjump As String
    Private _canback As String
    Private _canadd As String
    Private _candist As String
    Private _candistcc As String
    Private _canaddatt As String
    Private _candelatt As String
    Private _canconti As String
    Private _bypass As String
    Private _opinonly As String
    Private _block As Int32
    Private _layer As Int32
    Private _tree As Int32
    Private _backto_steps As Int32
    Private _assign_column_id As String
    Private _assigncc_column_id As String
    Private _after_submit_proc As String
    Private _when_receive_proc As String
    Private _alias_group_name As String
    Private _condi As String
    Private _backcanassign As String
    Private _backassign_column As String
    Private _allowtempsave As String
    Private _allowgotoend As String
    Private _cansendmail As String
    Private _overduenotice As String
    Private _attachlimit As Int32
    Private _allowdutyforgotoend As String
    Private _gotoendcc As String
    Private _canskip As String

    Property eformid() As String
        Get
            Return _eformid
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property eformrole() As Int32
        Get
            Return _eformrole
        End Get
        Set(value As Int32)
            Return
        End Set
    End Property
    Property stepsid() As Int32
        Get
            Return _stepsid
        End Get
        Set(value As Int32)
            Return
        End Set
    End Property
    Property steps() As Int32
        Get
            Return _steps
        End Get
        Set(value As Int32)
            Return
        End Set
    End Property
    Property nextstep() As Int32
        Get
            Return _nextstep
        End Get
        Set(value As Int32)
            Return
        End Set
    End Property
    Property major_step() As String
        Get
            Return _major_step
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property pors() As String
        Get
            Return _pors
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property x() As Int32
        Get
            Return _x
        End Get
        Set(value As Int32)
            Return
        End Set
    End Property
    Property y() As Int32
        Get
            Return _y
        End Get
        Set(value As Int32)
            Return
        End Set
    End Property
    Property group_id() As String
        Get
            Return _group_id
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property group_type() As String
        Get
            Return _group_type
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property allhandle() As String
        Get
            Return _allhandle
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property canjump() As String
        Get
            Return _canjump
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property canback() As String
        Get
            Return _canback
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property canadd() As String
        Get
            Return _canadd
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property candist() As String
        Get
            Return _candist
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property candistcc() As String
        Get
            Return _candistcc
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property canaddatt() As String
        Get
            Return _canaddatt
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property candelatt() As String
        Get
            Return _candelatt
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property canconti() As String
        Get
            Return _canconti
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property bypass() As String
        Get
            Return _bypass
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property opinonly() As String
        Get
            Return _opinonly
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property block() As Int32
        Get
            Return _block
        End Get
        Set(value As Int32)
            Return
        End Set
    End Property
    Property layer() As Int32
        Get
            Return _layer
        End Get
        Set(value As Int32)
            Return
        End Set
    End Property
    Property tree() As Int32
        Get
            Return _tree
        End Get
        Set(value As Int32)
            Return
        End Set
    End Property
    Property backto_steps() As Int32
        Get
            Return _backto_steps
        End Get
        Set(value As Int32)
            Return
        End Set
    End Property
    Property assign_column_id() As String
        Get
            Return _assign_column_id
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property assigncc_column_id() As String
        Get
            Return _assigncc_column_id
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property after_submit_proc() As String
        Get
            Return _after_submit_proc
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property when_receive_proc() As String
        Get
            Return _when_receive_proc
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property alias_group_name() As String
        Get
            Return _alias_group_name
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property condi() As String
        Get
            Return _condi
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property backcanassign() As String
        Get
            Return _backcanassign
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property backassign_column() As String
        Get
            Return _backassign_column
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property allowtempsave() As String
        Get
            Return _allowtempsave
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property allowgotoend() As String
        Get
            Return _allowgotoend
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property cansendmail() As String
        Get
            Return _cansendmail
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property overduenotice() As String
        Get
            Return _overduenotice
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property attachlimit() As Int32
        Get
            Return _attachlimit
        End Get
        Set(value As Int32)
            Return
        End Set
    End Property
    Property allowdutyforgotoend() As String
        Get
            Return _allowdutyforgotoend
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property gotoendcc() As String
        Get
            Return _gotoendcc
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    Property canskip() As String
        Get
            Return _canskip
        End Get
        Set(value As String)
            Return
        End Set
    End Property
    
    'Public Sub New(ByVal Server As String)
    '    MyBase.New(Server)
    'End Sub
    ' ''' <summary>
    ' ''' 
    ' ''' </summary>
    ' ''' <param name="sEFormid">表單種類編號</param>
    ' ''' <param name="sEFormsn">表單編號</param>
    ' ''' <param name="sEmpID">簽核人ID</param>
    ' ''' <param name="sGonogo">簽核狀態</param>
    ' ''' <remarks></remarks>
    'Public Sub New(ByVal sEFormid As String, ByVal sEFormsn As String, ByVal sEmpID As String, ByVal sGonogo As String)
    '    MyBase.New()
    '    Load(sEFormid, sEFormsn, sEmpID, sGonogo)
    'End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sFlowsn">流程編號</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal SID As Int32)
        MyBase.New()
        Load(SID)
    End Sub

    Public Sub New(ByVal SID As Int32, ByVal eformid As String)
        MyBase.New()
        Load(SID, eformid)
    End Sub

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub Load(ByVal sEFormid As String, ByVal sEFormsn As String, ByVal sEmpID As String, ByVal sGonogo As String)
        'Dim strSql As String
        'Dim DC As New SQLDBControl
        'Dim DR As SqlDataReader

        'strSql = "SELECT * FROM FLOWCTL WHERE EFORMID='" + sEFormid + "' AND EFORMSN='" + sEFormsn + "' AND EMPUID='" + sEmpID + "' AND GONOGO='" + sGonogo + "'"
        'DR = DC.CreateReader(strSql)
        'If DR.HasRows Then
        '    If DR.Read() Then
        '        _flowsn = CType(IIf(IsDBNull(DR("flowsn")), 0, DR("flowsn")), Int32)
        '        _eformid = CType(IIf(IsDBNull(DR("eformid")), "", DR("eformid")), String)
        '        _eformrole = CType(IIf(IsDBNull(DR("eformrole")), 0, DR("eformrole")), Int32)
        '        _eformsn = CType(IIf(IsDBNull(DR("eformsn")), "", DR("eformsn")), String)
        '        _stepsid = CType(DR("stepsid").ToString(), Int32)
        '        _steps = CType(IIf(IsDBNull(DR("steps")), 0, DR("steps")), Int32)
        '        _empuid = CType(IIf(IsDBNull(DR("empuid")), "", DR("empuid")), String)
        '        _emp_chinese_name = CType(IIf(IsDBNull(DR("emp_chinese_name")), "", DR("emp_chinese_name")), String)
        '        _group_name = CType(IIf(IsDBNull(DR("group_name")), "", DR("group_name")), String)
        '        If Not IsDBNull(DR("hddate")) Then _hddate = CType(DR("hddate").ToString(), DateTime)
        '        _gonogo = CType(IIf(IsDBNull(DR("gonogo")), "", DR("gonogo")), String)
        '        _comment = CType(IIf(IsDBNull(DR("comment")), "", DR("comment")), String)
        '        _nextstep = CType(DR("nextstep").ToString(), Int32)
        '        _goback = CType(IIf(IsDBNull(DR("goback")), 0, DR("goback")), Int32)
        '        _nextuser = CType(IIf(IsDBNull(DR("nextuser")), "", DR("nextuser")), String)
        '        _orgname = CType(IIf(IsDBNull(DR("orgname")), "", DR("orgname")), String)
        '        _important = CType(IIf(IsDBNull(DR("important")), "", DR("important")), String)
        '        _backto_steps = CType(IIf(IsDBNull(DR("backto_steps")), 0, DR("backto_steps")), Int32)
        '        If Not IsDBNull(DR("recdate")) Then _recdate = CType(DR("recdate").ToString, DateTime)
        '        _is_testmode = CType(IIf(IsDBNull(DR("is_testmode")), True, DR("is_testmode")), Boolean)
        '        _duty_for = CType(IIf(IsDBNull(DR("duty_for")), "", DR("duty_for")), String)
        '        If Not IsDBNull(DR("appdate")) Then _appdate = CType(DR("appdate").ToString(), DateTime)
        '        _filler = DR("filler").ToString()
        '        _subtype = DR("subtype").ToString()
        '        _deptcode = DR("deptcode").ToString()
        '        If Not IsDBNull(DR("createdate")) Then _createdate = CType(DR("createdate"), DateTime)
        '    End If
        'End If
        'DC.Dispose()
    End Sub

    Public Sub Load(ByVal sid As Int32)
        Dim strSql As String
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader

        strSql = "SELECT * FROM FLOW WHERE STEPSID=" + sid.ToString()
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read() Then
                _eformid = CType(IIf(IsDBNull(DR("eformid")), "", DR("eformid")), String)
                _eformrole = CType(IIf(IsDBNull(DR("eformrole")), 0, DR("eformrole")), Int32)
                _stepsid = sid
                _steps = CType(IIf(IsDBNull(DR("steps")), 0, DR("steps")), Int32)
                _nextstep = CType(IIf(IsDBNull(DR("nextstep")), 0, DR("nextstep")), Int32)
                _major_step = CType(IIf(IsDBNull(DR("major_step")), "", DR("major_step")), String)
                _pors = CType(IIf(IsDBNull(DR("pors")), "", DR("pors")), String)
                _x = CType(IIf(IsDBNull(DR("x")), 0, DR("x")), Int32)
                _y = CType(IIf(IsDBNull(DR("y")), 0, DR("y")), Int32)
                _group_id = CType(IIf(IsDBNull(DR("group_id")), "", DR("group_id")), String)
                _group_type = CType(IIf(IsDBNull(DR("group_type")), "", DR("group_type")), String)
                _allhandle = CType(IIf(IsDBNull(DR("allhandle")), "", DR("allhandle")), String)
                _canjump = CType(IIf(IsDBNull(DR("canjump")), "", DR("canjump")), String)
                _canback = CType(IIf(IsDBNull(DR("canback")), "", DR("canback")), String)
                _canadd = CType(IIf(IsDBNull(DR("canadd")), "", DR("canadd")), String)
                _candist = CType(IIf(IsDBNull(DR("candist")), "", DR("candist")), String)
                _candistcc = CType(IIf(IsDBNull(DR("candistcc")), "", DR("candistcc")), String)
                _canaddatt = CType(IIf(IsDBNull(DR("canaddatt")), "", DR("canaddatt")), String)
                _candelatt = CType(IIf(IsDBNull(DR("candelatt")), "", DR("candelatt")), String)
                _canconti = CType(IIf(IsDBNull(DR("canconti")), "", DR("canconti")), String)
                _bypass = CType(IIf(IsDBNull(DR("bypass")), "", DR("bypass")), String)
                _opinonly = CType(IIf(IsDBNull(DR("opinonly")), "", DR("opinonly")), String)
                _block = CType(IIf(IsDBNull(DR("block")), 0, DR("block")), Int32)
                _layer = CType(IIf(IsDBNull(DR("layer")), 0, DR("layer")), Int32)
                _tree = CType(IIf(IsDBNull(DR("tree")), 0, DR("tree")), Int32)
                _backto_steps = CType(IIf(IsDBNull(DR("backto_steps")), 0, DR("backto_steps")), Int32)
                _assign_column_id = CType(IIf(IsDBNull(DR("assign_column_id")), "", DR("assign_column_id")), String)
                _assigncc_column_id = CType(IIf(IsDBNull(DR("assigncc_column_id")), "", DR("assigncc_column_id")), String)
                _after_submit_proc = CType(IIf(IsDBNull(DR("after_submit_proc")), "", DR("after_submit_proc")), String)
                _when_receive_proc = CType(IIf(IsDBNull(DR("when_receive_proc")), "", DR("when_receive_proc")), String)
                _alias_group_name = CType(IIf(IsDBNull(DR("alias_group_name")), "", DR("alias_group_name")), String)
                _condi = CType(IIf(IsDBNull(DR("condi")), "", DR("condi")), String)
                _backcanassign = CType(IIf(IsDBNull(DR("backcanassign")), "", DR("backcanassign")), String)
                _backassign_column = CType(IIf(IsDBNull(DR("backassign_column")), "", DR("backassign_column")), String)
                _allowtempsave = CType(IIf(IsDBNull(DR("allowtempsave")), "", DR("allowtempsave")), String)
                _allowgotoend = CType(IIf(IsDBNull(DR("allowgotoend")), "", DR("allowgotoend")), String)
                _cansendmail = CType(IIf(IsDBNull(DR("cansendmail")), "", DR("cansendmail")), String)
                _overduenotice = CType(IIf(IsDBNull(DR("overduenotice")), "", DR("overduenotice")), String)
                _attachlimit = CType(IIf(IsDBNull(DR("attachlimit")), 0, DR("attachlimit")), Int32)
                _allowdutyforgotoend = CType(IIf(IsDBNull(DR("allowdutyforgotoend")), "", DR("allowdutyforgotoend")), String)
                _gotoendcc = CType(IIf(IsDBNull(DR("gotoendcc")), "", DR("gotoendcc")), String)
                _canskip = CType(IIf(IsDBNull(DR("canskip")), "", DR("canskip")), String)
            End If
        End If
        DC.Dispose()
    End Sub

    Public Sub Load(ByVal sid As Int32, ByVal eformid As String)
        Dim strSql As String
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader

        strSql = "SELECT * FROM FLOW WHERE EFORMID='" + eformid + "' AND STEPSID=" + sid.ToString()
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read() Then
                _eformid = CType(IIf(IsDBNull(DR("eformid")), "", DR("eformid")), String)
                _eformrole = CType(IIf(IsDBNull(DR("eformrole")), 0, DR("eformrole")), Int32)
                _stepsid = sid
                _steps = CType(IIf(IsDBNull(DR("steps")), 0, DR("steps")), Int32)
                _nextstep = CType(IIf(IsDBNull(DR("nextstep")), 0, DR("nextstep")), Int32)
                _major_step = CType(IIf(IsDBNull(DR("major_step")), "", DR("major_step")), String)
                _pors = CType(IIf(IsDBNull(DR("pors")), "", DR("pors")), String)
                _x = CType(IIf(IsDBNull(DR("x")), 0, DR("x")), Int32)
                _y = CType(IIf(IsDBNull(DR("y")), 0, DR("y")), Int32)
                _group_id = CType(IIf(IsDBNull(DR("group_id")), "", DR("group_id")), String)
                _group_type = CType(IIf(IsDBNull(DR("group_type")), "", DR("group_type")), String)
                _allhandle = CType(IIf(IsDBNull(DR("allhandle")), "", DR("allhandle")), String)
                _canjump = CType(IIf(IsDBNull(DR("canjump")), "", DR("canjump")), String)
                _canback = CType(IIf(IsDBNull(DR("canback")), "", DR("canback")), String)
                _canadd = CType(IIf(IsDBNull(DR("canadd")), "", DR("canadd")), String)
                _candist = CType(IIf(IsDBNull(DR("candist")), "", DR("candist")), String)
                _candistcc = CType(IIf(IsDBNull(DR("candistcc")), "", DR("candistcc")), String)
                _canaddatt = CType(IIf(IsDBNull(DR("canaddatt")), "", DR("canaddatt")), String)
                _candelatt = CType(IIf(IsDBNull(DR("candelatt")), "", DR("candelatt")), String)
                _canconti = CType(IIf(IsDBNull(DR("canconti")), "", DR("canconti")), String)
                _bypass = CType(IIf(IsDBNull(DR("bypass")), "", DR("bypass")), String)
                _opinonly = CType(IIf(IsDBNull(DR("opinonly")), "", DR("opinonly")), String)
                _block = CType(IIf(IsDBNull(DR("block")), 0, DR("block")), Int32)
                _layer = CType(IIf(IsDBNull(DR("layer")), 0, DR("layer")), Int32)
                _tree = CType(IIf(IsDBNull(DR("tree")), 0, DR("tree")), Int32)
                _backto_steps = CType(IIf(IsDBNull(DR("backto_steps")), 0, DR("backto_steps")), Int32)
                _assign_column_id = CType(IIf(IsDBNull(DR("assign_column_id")), "", DR("assign_column_id")), String)
                _assigncc_column_id = CType(IIf(IsDBNull(DR("assigncc_column_id")), "", DR("assigncc_column_id")), String)
                _after_submit_proc = CType(IIf(IsDBNull(DR("after_submit_proc")), "", DR("after_submit_proc")), String)
                _when_receive_proc = CType(IIf(IsDBNull(DR("when_receive_proc")), "", DR("when_receive_proc")), String)
                _alias_group_name = CType(IIf(IsDBNull(DR("alias_group_name")), "", DR("alias_group_name")), String)
                _condi = CType(IIf(IsDBNull(DR("condi")), "", DR("condi")), String)
                _backcanassign = CType(IIf(IsDBNull(DR("backcanassign")), "", DR("backcanassign")), String)
                _backassign_column = CType(IIf(IsDBNull(DR("backassign_column")), "", DR("backassign_column")), String)
                _allowtempsave = CType(IIf(IsDBNull(DR("allowtempsave")), "", DR("allowtempsave")), String)
                _allowgotoend = CType(IIf(IsDBNull(DR("allowgotoend")), "", DR("allowgotoend")), String)
                _cansendmail = CType(IIf(IsDBNull(DR("cansendmail")), "", DR("cansendmail")), String)
                _overduenotice = CType(IIf(IsDBNull(DR("overduenotice")), "", DR("overduenotice")), String)
                _attachlimit = CType(IIf(IsDBNull(DR("attachlimit")), 0, DR("attachlimit")), Int32)
                _allowdutyforgotoend = CType(IIf(IsDBNull(DR("allowdutyforgotoend")), "", DR("allowdutyforgotoend")), String)
                _gotoendcc = CType(IIf(IsDBNull(DR("gotoendcc")), "", DR("gotoendcc")), String)
                _canskip = CType(IIf(IsDBNull(DR("canskip")), "", DR("canskip")), String)
            End If
        End If
        DC.Dispose()
    End Sub

    Public Function PreFlow() As flow
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String

        strSql = "SELECT * FROM FLOW WHERE NEXTSTEP=" + _stepsid.ToString()
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read() Then                
                Return New flow(CType(DR("STEPSID").ToString(), Int32))                
            End If
        End If
        DC.Dispose()
        Return New flow()
    End Function

    Public Function NextFlow() As flow
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String

        strSql = "SELECT * FROM FLOW WHERE STEPSID=" + _nextstep.ToString()
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read() Then
                Return New flow(CType(DR("STEPSID").ToString(), Int32))
            End If
        End If
        DC.Dispose()
        Return New flow()
    End Function

    Public Sub EFormFlows(ByVal eformid As String, ByVal stepid As String, ByRef FlowList As List(Of KeyValuePair(Of Integer, String)))
        Dim strSql As String = "SELECT * FROM FLOW WHERE EFORMID='" + eformid + "'" + IIf(stepid.Length > 0, " AND STEPSID=" + stepid, " AND STEPDIS='1'")
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader

        Dim FlowsCount As Integer = FlowList.Count
        DR = DC.CreateReader(strSql)
        If DR.HasRows Then
            If DR.Read Then
                FlowList.Add(New KeyValuePair(Of Integer, String)((FlowsCount + 1), DR("stepsid")))
                If DR("nextstep") <> "-1" Then EFormFlows(eformid, DR("nextstep"), FlowList)
            End If
        End If
        DC.Dispose()
    End Sub


End Class