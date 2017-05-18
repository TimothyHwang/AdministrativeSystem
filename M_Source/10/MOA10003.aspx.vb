Imports System.Data.SqlClient

Partial Class M_Source_10_MOA10003
    Inherits Page

    ReadOnly sql_function As New C_SQLFUN
    ReadOnly connstr As String = sql_function.G_conn_string
    Dim db As New SqlConnection(connstr)

#Region "Custom Function"
    ''**********************  以下為Custom Function  **************************

    ''' <summary>
    ''' 產生組織樹狀
    ''' </summary>
    ''' <param name="tv"></param>
    ''' <remarks></remarks>
    Protected Sub LoadDEPTree(ByRef tv As TreeView)
        ''建立XML節點
        Dim xmlRootNode As New TreeNode

        xmlRootNode.Text = "組織總覽"
        xmlRootNode.Value = "0"
        xmlRootNode.SelectAction = TreeNodeSelectAction.Select
        'xmlRootNode.NavigateUrl = "javascript:void(0)"

        CreateTree(xmlRootNode, "0")

        tv.Nodes.Clear()
        tv.Nodes.Add(xmlRootNode)
        tv.Attributes.Add("onclick", "CheckEvent()")
    End Sub

    ''' <summary>
    ''' 產生部門AdminGroup的
    ''' TreeView所需要的TreeNodes
    ''' </summary>
    ''' <param name="ParentNode"></param>
    ''' <param name="U_ID"></param>
    ''' <remarks></remarks>
    Protected Sub CreateTree(ByRef ParentNode As TreeNode, ByVal U_ID As String)
        ''建立樹狀資料
        Dim dbCT As New SqlConnection(connstr)
        dbCT.Open()
        Dim strSql As String = "SELECT * FROM ADMINGROUP WHERE PARENT_ORG_UID='" & U_ID & "'"
        Dim DR As SqlDataReader = New SqlCommand(strSql, dbCT).ExecuteReader

        While (DR.Read)
            Dim xmlTreeNode As New TreeNode
            xmlTreeNode.Text = DR("ORG_NAME").ToString()
            xmlTreeNode.Value = DR("ORG_UID").ToString()
            xmlTreeNode.SelectAction = TreeNodeSelectAction.Select
            ''xmlTreeNode.NavigateUrl = "javascript:void(0)"

            ParentNode.ChildNodes.Add(xmlTreeNode)
            'GetMembersFromDeptToTree(xmlTreeNode, DR("ORG_UID").ToString())
            If HasChildNodes(DR("ORG_UID").ToString()) Then
                CreateTree(xmlTreeNode, DR("ORG_UID").ToString())
            End If
        End While
        dbCT.Close()
        dbCT.Dispose()
    End Sub

    ''' <summary>   
    ''' 將部門下員工資料加入
    ''' TreeView。  
    ''' 無回傳值，但需傳址。
    ''' </summary>   
    ''' <param name="ParentNode">要放入資料的TreeView。</param>   
    ''' <param name="D_ID">部門ID。</param>   
    Public Sub GetMembersFromDeptToTree(ByRef ParentNode As TreeNode, ByVal D_ID As String)
        Dim dbGMFDTT As New SqlConnection(connstr)
        dbGMFDTT.Open()

        Dim DR As SqlDataReader

        Dim strSql As String = "SELECT EMPUID,EMPLOYEE_ID,ORG_UID,EMP_CHINESE_NAME FROM EMPLOYEE WHERE LEAVE='Y' AND ORG_UID='" & D_ID & "' ORDER BY EMP_CHINESE_NAME"

        DR = New SqlCommand(strSql, dbGMFDTT).ExecuteReader

        While (DR.Read)
            Dim xmlTreeNode As New TreeNode
            xmlTreeNode.Text = DR("EMP_CHINESE_NAME").ToString()
            xmlTreeNode.Value = DR("EMPUID").ToString()
            xmlTreeNode.SelectAction = TreeNodeSelectAction.Select
            ''xmlTreeNode.NavigateUrl = "javascript:void(0)"

            ParentNode.ChildNodes.Add(xmlTreeNode)
        End While
        dbGMFDTT.Close()
        dbGMFDTT.Dispose()
    End Sub

    ''' <summary>
    ''' 判斷是否還有
    ''' 子節點
    ''' </summary>
    ''' <param name="pID">欲判斷的節點</param>
    ''' <returns>是：有子節點；否：無子節點</returns>
    ''' <remarks></remarks>
    Protected Function HasChildNodes(ByVal pID As String) As Boolean
        Dim dbICN As New SqlConnection(connstr)
        dbICN.Open()
        Dim DR As SqlDataReader
        Dim IsChildNode As Boolean

        Dim strSql As String = "SELECT * FROM ADMINGROUP WHERE PARENT_ORG_UID='" & pID & "'"

        DR = New SqlCommand(strSql, dbICN).ExecuteReader

        If (DR.Read) Then
            IsChildNode = True
        Else
            IsChildNode = False
        End If
        dbICN.Close()
        dbICN.Dispose()
        Return IsChildNode
    End Function

    ''' <summary>
    ''' 判斷主官在營
    ''' P_1001是否有資料
    ''' </summary>
    ''' <param name="sManagerId"></param>
    ''' <param name="sEmployeeId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function IsExistP_1001(ByVal sManagerId As String, ByVal sEmployeeId As String) As Boolean
        Dim boolResult As Boolean

        Dim conn As New SqlConnection()
        conn.ConnectionString = ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString
        conn.Open()

        Dim cmd As New SqlCommand("SELECT P_Num FROM P_1001 WHERE MANAGER_ID='" & sManagerId & "' AND EMPLOYEE_ID='" & sEmployeeId & "'", conn)
        Dim DR As SqlDataReader = cmd.ExecuteReader
        boolResult = DR.HasRows

        cmd.Dispose()
        DR.Close()
        conn.Close()
        conn.Dispose()

        'Using connA As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
        '    connA.Open()
        '    'Dim trans As SqlTransaction
        '    'trans = connA.BeginTransaction

        '    Dim strSql = "SELECT P_Num FROM P_1001 WHERE MANAGER_ID='" & sManagerId & "' AND EMPLOYEE_ID='" & sEmployeeId & "'"
        '    'Dim DR As SqlDataReader = New SqlCommand(strSql, connA, trans).ExecuteReader
        '    Dim DR As SqlDataReader = New SqlCommand(strSql, connA).ExecuteReader
        '    boolResult = DR.HasRows
        '    DR.Close()
        '    'trans.Commit()
        '    'trans.Dispose()
        'End Using


        Return boolResult
    End Function
#End Region

#Region "Form Function"
    ''**********************  以下為Form Function  ****************************

    Protected Sub btnEmployeeModify_Click(sender As Object, e As EventArgs) Handles btnEmployeeModify.Click

        Try


            Dim gv As GridView = gvMembers
            Dim Mid As String = ViewState("Mid").ToString()

            db.Open()
            Dim trans As SqlTransaction = db.BeginTransaction

            Dim strSql As String

            For Each row As GridViewRow In gv.Rows

                Dim chk As CheckBox = CType(row.Cells(0).FindControl("chkSelect"), CheckBox)
                Dim rowIndex As Integer = row.RowIndex
                Dim strEmployeeId As String = CInt(gv.DataKeys(rowIndex).Value).ToString()
                strSql = "INSERT INTO P_1001 (MANAGER_ID,EMPLOYEE_ID) VALUES ('" & Mid & "','" & strEmployeeId & "')"
                Call New SqlCommand(strSql, db, trans).ExecuteNonQuery()

                If chk IsNot Nothing Then

                    If chk.Checked Then
                        strSql = "DELETE FROM P_1001 WHERE MANAGER_ID='" & Mid & "' AND EMPLOYEE_ID='" & strEmployeeId & "'"
                        Call New SqlCommand(strSql, db, trans).ExecuteNonQuery()

                        strSql = "INSERT INTO P_1001 (MANAGER_ID,EMPLOYEE_ID) VALUES ('" & Mid & "','" & strEmployeeId & "')"
                        Call New SqlCommand(strSql, db, trans).ExecuteNonQuery()

                    Else
                        strSql = "DELETE FROM P_1001 WHERE MANAGER_ID='" & Mid & "' AND EMPLOYEE_ID='" & strEmployeeId & "'"
                        Call New SqlCommand(strSql, db, trans).ExecuteNonQuery()

                    End If
                End If
            Next
            trans.Commit()
        Catch ex As Exception

        Finally
            db.Close()
        End Try
    End Sub

    Protected Sub gvMembers_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gvMembers.RowDataBound
        Dim Mid As String = ViewState("Mid").ToString()
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                Dim chk As CheckBox = CType(e.Row.FindControl("chkSelect"), CheckBox)
                If chk Is Nothing Then Return
                Dim uid As String = e.Row.Cells(1).Text
                chk.Checked = IsExistP_1001(Mid, uid)
        End Select
    End Sub

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not IsPostBack Then
            ViewState("Mid") = Request.QueryString("Mid")
            ViewState("ManagerName") = Request.QueryString("MName")
            If (Not ViewState("DEPTreeNodes") = Nothing) AndAlso (ViewState("DEPTreeNodes").ToString.Length > 0) Then
                tvEmployee.Nodes.Clear()
                tvEmployee.Nodes.Add(CType(ViewState("DEPTreeNodes"), TreeNode))
                tvEmployee.Attributes.Add("onclick", "CheckEvent()")

            Else
                LoadDEPTree(tvEmployee)
            End If
            lblManager.Text = ViewState("ManagerName").ToString

            tvEmployee.ImageSet = TreeViewImageSet.Arrows
            'Dim trigger = New AsyncPostBackTrigger
            'trigger.ControlID = tvEmployee.UniqueID
            'trigger.EventName = "SelectedNodeChanged"
            'UPMembers.Triggers.Add(trigger)
        End If
        Page.Header.DataBind()
    End Sub

    Protected Sub tvEmployee_SelectedNodeChanged(sender As Object, e As EventArgs) Handles tvEmployee.SelectedNodeChanged
        gvMembers.DataSourceID = Nothing
        gvMembers.DataSourceID = SqlDataSourceMembers.ID
        gvMembers.DataBind()
    End Sub

    Protected Sub SqlDataSourceMembers_Selected(sender As Object, e As SqlDataSourceStatusEventArgs) Handles SqlDataSourceMembers.Selected
        btnEmployeeModify.Visible = e.AffectedRows > 0
    End Sub
#End Region
End Class
