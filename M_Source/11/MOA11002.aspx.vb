Imports System.Data.SqlClient
Imports System.Linq

Partial Class M_Source_11_MOA11002
    Inherits Page

    Public flowAdmin As String
    Public user_id As String
    Public org_uid As String

    Protected Function SelectWholeTreeORG_UID(ByVal sUser_id As String) As String
        Dim sReturn As String = ""
        Dim tool As New C_Public

        sReturn = tool.GetWholeOrgIDs(sUser_id, ",", "'")
        SelectWholeTreeORG_UID = sReturn
    End Function

    Protected Sub CreateMemberORGTree(ByRef ParentNode As TreeNode, ByVal U_ID As String)
        ''database connection for SQL
        Dim tool As New C_Public
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim strSql As String = "SELECT * FROM ADMINGROUP WHERE PARENT_ORG_UID='" & U_ID & "'"
        ''系統管理者可看全部單位
        If Session("Role") <> "1" Then strSql += " AND ORG_TREE_LEVEL>3"
        DR = DC.CreateReader(strSql)
        While DR.Read
            ''combine tree
            Dim xmlTreeNode As New TreeNode
            xmlTreeNode.Text = DR("ORG_NAME").ToString()
            xmlTreeNode.Value = "ORG_" + DR("ORG_UID").ToString()
            xmlTreeNode.ToolTip = DR("ORG_NAME").ToString()
            xmlTreeNode.SelectAction = TreeNodeSelectAction.Select
            xmlTreeNode.NavigateUrl = "javascript:void(0)"

            ParentNode.ChildNodes.Add(xmlTreeNode)
            tool.GetMembersFromORGToTreeNoPB(xmlTreeNode, DR("ORG_UID").ToString())

            ''child loop
            If IsMemberChildNodes(DR("ORG_UID").ToString()) Then
                CreateMemberORGTree(xmlTreeNode, DR("ORG_UID").ToString())
            End If
        End While
        DC.Dispose()
    End Sub

    Protected Function IsMemberChildNodes(ByVal pID As String) As Boolean
        Dim DC As New SQLDBControl
        Dim DR As SqlDataReader
        Dim IsChildNode As Boolean

        Dim strSql As String = "SELECT * FROM ADMINGROUP WHERE PARENT_ORG_UID='" & pID & "'"
        ''系統管理者可看全部單位
        If Session("Role") <> "1" Then strSql += " AND ORG_TREE_LEVEL>3"

        DR = DC.CreateReader(strSql)

        If (DR.Read) Then
            IsChildNode = True
        Else
            IsChildNode = False
        End If
        DC.Dispose()

        Return IsChildNode
    End Function

    Protected Function Show(ByVal sFlowadmin As String) As String
        'Dim sReturn As String = "" ''class="hide"
        'If sFlowadmin <> "1" Then
        '    sReturn = " class=""hide"""
        'End If
        'If Session("Role") = "1" Then
        '    sReturn = ""
        'End If
        'Show = sReturn
        Return ""
    End Function

    Protected Function SQLALL() As String
        Dim sReturn As String = "SELECT * FROM P_11 WHERE 1=1 "

        Dim tool As New C_Public
        'Dim arr() As String = Split(lblOrgSelID.Text, "_")
        'Dim OrgNames As String = ""
        'Select Case arr(0)
        '    Case "EMP"
        '        sReturn += " AND PAIDNO='" + arr(1) + "'"
        '    Case "ORG"
        '        tool.GetOrgChildNamesByID(OrgNames, arr(1), "'")
        '        sReturn += "AND PAUNIT IN(" + OrgNames + ")"
        'End Select
        '部門名稱
        If ddlOrgSel.Items.Count > 0 AndAlso ddlOrgSel.SelectedValue.ToString() <> "0" Then
            sReturn += "AND PAUNIT='" & ddlOrgSel.SelectedItem.Text & "'"
        End If
        If ddlUserSel.Items.Count > 0 AndAlso ddlUserSel.SelectedValue.ToString() <> "0" Then
            sReturn += " AND PAIDNO='" & ddlUserSel.SelectedValue & "'"
        End If
        ''申請時間
        If txtAPPSDate.Text.Length > 0 AndAlso txtAPPEDate.Text.Length > "0" Then
            sReturn += " AND (APPTIME>='" & txtAPPSDate.Text & " 0:0:0' AND APPTIME<='" & txtAPPEDate.Text & " 23:59:59')"
        End If
        ''種類
        If (ddlRepairMainKind.Items.Count > 0 AndAlso ddlProblemKind.Items.Count > 0) AndAlso (ddlRepairMainKind.SelectedValue.ToString() <> "0" AndAlso ddlProblemKind.SelectedValue.ToString() <> "0") Then
            sReturn += " AND BROKENTYPE='" & ddlRepairMainKind.SelectedValue & ddlProblemKind.SelectedValue & "'"
        ElseIf (ddlRepairMainKind.Items.Count > 0 AndAlso ddlProblemKind.Items.Count > 0) AndAlso (ddlRepairMainKind.SelectedValue.ToString() <> "0" AndAlso ddlProblemKind.SelectedValue.ToString() = "0") Then
            sReturn += " AND BROKENTYPE LIKE '" & ddlRepairMainKind.SelectedValue & "%'"
        End If
        ''狀態
        If ddlStatusSel.Items.Count > 0 AndAlso ddlStatusSel.SelectedValue.ToString() <> "99" Then
            If ddlStatusSel.SelectedValue = "E" Then
                sReturn += " AND (PENDFLAG='" & ddlStatusSel.SelectedValue & "'  OR PENDFLAG='0')"
            Else
                sReturn += " AND (PENDFLAG='" & ddlStatusSel.SelectedValue & "'  OR PENDFLAG IS NULL)"
            End If

        End If
        ''完修時間
        If txtFixedSDate.Text.Length > 0 AndAlso txtFixedEDate.Text.Length > 0 Then
            sReturn += " AND (FINALDATE>='" & txtFixedSDate.Text & " 0:0:0' AND FINALDATE<='" & txtFixedEDate.Text & " 23:59:59')"
        End If
        ''報修時間
        If txtCallSDate.Text.Length > 0 AndAlso txtCallEDate.Text.Length > 0 Then
            sReturn += " AND (CALLTIME>='" & txtCallSDate.Text & " 0:0:0' AND CALLTIME<='" & txtCallEDate.Text & " 23:59:59')"
        End If
        ''到修時間
        If txtARRSDate.Text.Length > 0 AndAlso txtARREDate.Text.Length > 0 Then
            sReturn += " AND (ARRIVETIME>='" & txtARRSDate.Text & " 0:0:0' AND ARRIVETIME<='" & txtARRSDate.Text & " 23:59:59')"
        End If

        SQLALL = sReturn
    End Function

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim tool As New C_Public

        '取得登入者帳號
        If Page.User.Identity.Name.ToString.IndexOf("\", StringComparison.Ordinal) > 0 Then

            Dim LoginAll As String = Page.User.Identity.Name.ToString

            Dim LoginID() As String = Split(LoginAll, "\")

            user_id = LoginID(1)
        Else
            user_id = Page.User.Identity.Name.ToString
        End If

        org_uid = tool.GetOrgIDByIDNo(user_id)
        If tool.CheckStepGroupEMPByName("單位資訊官", user_id, tool.GetObjectTypeFromStep("單位資訊官")) Then
            flowAdmin = "1"
            btnSelOrg.Enabled = True
            ddlOrgSel.Enabled = True
            ddlUserSel.Enabled = True
        ElseIf tool.CheckStepGroupEMPByName("資訊報修管制單位", user_id, tool.GetObjectTypeFromStep("資訊報修管制單位")) Then
            flowAdmin = "2"
            btnSelOrg.Enabled = True
            ddlOrgSel.Enabled = True
            ddlUserSel.Enabled = True
        ElseIf tool.CheckStepGroupEMPByName("資訊維修單位", user_id, tool.GetObjectTypeFromStep("資訊維修單位")) Then
            flowAdmin = "3"
            btnSelOrg.Enabled = True
            ddlOrgSel.Enabled = True
            ddlUserSel.Enabled = True
        ElseIf Session("Role") = "1" Then
            flowAdmin = "0"
            btnSelOrg.Enabled = True
            ddlOrgSel.Enabled = True
            ddlUserSel.Enabled = True
        Else            
            flowAdmin = "0"
            If tool.IsFixmanSupervisor(user_id) Then
                btnSelOrg.Enabled = True
                ddlOrgSel.Enabled = True
                ddlUserSel.Enabled = True
            Else
                btnSelOrg.Enabled = False
                ddlOrgSel.Enabled = False
                ddlUserSel.Enabled = False
            End If
        End If

        ''下拉單位讀取
        'SqlDataSource1.SelectCommand = "SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] WHERE ORG_UID IN (" + SelectWholeTreeORG_UID(org_uid) + ") ORDER BY [ORG_NAME]"
        'ddlOrgSel.DataBind()

        If Not IsPostBack Then

            '判斷登入者權限
            If Session("Role") = "1" Or flowAdmin = "2" Or flowAdmin = "3" Then
                SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
            ElseIf flowAdmin = "1" Then
                SqlDataSource1.SelectCommand = "SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] WHERE ORG_UID IN (" + SelectWholeTreeORG_UID(user_id) + ") ORDER BY [ORG_NAME]"
            Else
                ''維修人員的上一級主管需可查全連參
                SqlDataSource1.SelectCommand = "SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME"
                If tool.IsFixmanSupervisor(user_id) Then
                    ddlOrgSel.Enabled = True
                Else
                    ddlOrgSel.Enabled = False
                End If
            End If
            ddlOrgSel.DataBind()

            ' ''組織部門樹狀圖
            ' ''暫時取消使用樹狀選取部門及人員
            'lblOrgSel.Text = tool.GetOrgNameByIDNo(user_id) + "-" + tool.GetUserNameByID(user_id)
            'lblOrgSelID.Text = "EMP_" + user_id

            'Dim strEMPRootORGID = tool.EMPRootORGID(user_id, 3)
            ' ''系統管理者可看全部單位
            'If Session("Role") = "1" Then strEMPRootORGID = "EDKOr12jWP"
            'Dim DC As New SQLDBControl
            'Dim DR As SqlDataReader
            'Dim strSql As String = "SELECT * FROM ADMINGROUP WHERE ORG_UID='" + strEMPRootORGID + "'"
            'DR = DC.CreateReader(strSql)
            'If DR.HasRows Then
            '    If DR.Read Then

            '        Dim TreeViewORG As TreeView = tvOrg
            '        Dim xmlTreeViewORGNode As New TreeNode

            '        xmlTreeViewORGNode.Text = DR("ORG_NAME").ToString()
            '        xmlTreeViewORGNode.Value = "ORG_" + DR("ORG_UID").ToString()
            '        xmlTreeViewORGNode.SelectAction = TreeNodeSelectAction.None

            '        CreateMemberORGTree(xmlTreeViewORGNode, DR("ORG_UID").ToString())
            '        TreeViewORG.Nodes.Clear()
            '        TreeViewORG.Nodes.Add(xmlTreeViewORGNode)
            '    End If
            'End If
        End If

        'SqlDataSource2.SelectCommand = SQLALL()
        'GridView1.DataBind()

        'SqlDataSource4.SelectCommand = "SELECT DISTINCT KIND_NUM,KIND_NAME FROM [SYSKIND] WHERE ([Kind_Num] IN ('7','8','9')) "
        'ddlRepairMainKind.DataBind()

    End Sub

    Protected Sub GridView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gvList.SelectedIndexChanged
        Dim streformsn As String
        Dim strPath As String

        '顯示選取的表單資料        
        streformsn = gvList.Rows(gvList.SelectedIndex).Cells(0).Text

        '表單資料夾
        strPath = "../00/MOA00020.aspx?x=MOA11001&y=BL7U2QP3IG&Read_Only=1&EFORMSN=" & streformsn

        Response.Write(" <script language='javascript'>")
        Response.Write(" sPath = '" & strPath & "';")
        Response.Write(" strFeatures = 'dialogWidth=900px;dialogHeight=750px;help=no;status=no;resizable=yes;scroll=no;dialogTop=100;dialogLeft=100';")
        Response.Write(" showModalDialog(sPath,self,strFeatures);")
        Response.Write(" </script>")
    End Sub

    Protected Sub GridView1_RowCreated(sender As Object, e As GridViewRowEventArgs) Handles gvList.RowCreated
        If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then
            '隱藏eformsn
            e.Row.Cells(0).Visible = False
            'e.Row.Cells(8).Visible = False
        End If
    End Sub

    Protected Sub GridView1_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gvList.RowDataBound
        If e.Row.RowType.Equals(DataControlRowType.DataRow) Then
            ''報修種類 Rows(4)
            Dim sMainKind As String = ""
            Dim sSubKind As String = ""
            Dim sKind As String = ""
            sKind = e.Row.Cells(4).Text
            sMainKind = sKind.Substring(0, 1)
            sSubKind = sKind.Substring(1, 2)
            Dim DC As New SQLDBControl
            Dim strSql As String = ""
            Dim DR As SqlDataReader
            strSql = "SELECT * FROM SYSKIND WHERE KIND_SYSID='" & sSubKind & "'"
            DR = DC.CreateReader(strSql)
            If DR.Read() Then
                e.Row.Cells(4).Text = DR("KIND_NAME").ToString() & "-" & DR("STATE_NAME").ToString()
            End If
            DC.Dispose()

            ''維修進度 Rows(5)
            Dim lblFixStatus As Label = CType(e.Row.FindControl("lblFixStatus"), Label)
            If Not IsNothing(lblFixStatus) Then
                strSql = "SELECT TOP 1 * FROM FLOWCTL WHERE EFORMID='BL7U2QP3IG' AND EFORMSN='" + e.Row.Cells(0).Text + "' ORDER BY FLOWSN DESC"
                DC = New SQLDBControl()
                DR = DC.CreateReader(strSql)
                If DR.HasRows Then
                    If DR.Read() Then
                        Select Case DR("gonogo").ToString()
                            Case "-"
                                lblFixStatus.Text = "申請"
                            Case "F"
                                lblFixStatus.Text = "修繕"
                            Case "N"
                                lblFixStatus.Text = "不修繕"
                            Case "C"
                                lblFixStatus.Text = "完工"
                            Case "0"
                                lblFixStatus.Text = "駁回"
                            Case "1"
                                lblFixStatus.Text = "送件"
                            Case "?"
                                lblFixStatus.Text = "審核中"
                            Case "E"
                                lblFixStatus.Text = "完成"
                            Case "G"
                                lblFixStatus.Text = "補登"
                            Case "B", "X"
                                lblFixStatus.Text = "申請者撤銷"
                            Case "R"
                                lblFixStatus.Text = "重新分派"
                            Case "T"
                                lblFixStatus.Text = "呈轉"
                            Case "2"
                                lblFixStatus.Text = "退件"
                            Case Else
                                lblFixStatus.Text = "未知"
                        End Select
                    End If
                End If
            End If
            'Dim sStatus As String = e.Row.Cells(5).Text
            'Select Case sStatus
            '    Case "0"
            '        e.Row.Cells(5).Text = "未處理"
            '    Case "1"
            '        e.Row.Cells(5).Text = "處理中"
            '    Case "2"
            '        e.Row.Cells(5).Text = "處理完成"
            'End Select
        ElseIf e.Row.RowType = DataControlRowType.Pager Then
            ''設定GridView頁碼列
            Dim ddlPageNo As DropDownList = CType(e.Row.FindControl("ddlPageNo"), DropDownList)
            Dim lblPageCount As Label = CType(e.Row.FindControl("lblPageCount"), Label)
            Dim ddlPageSize As DropDownList = CType(e.Row.FindControl("ddlPageSize"), DropDownList)

            ddlPageNo.Items.Clear()
            For i = 1 To gvList.PageCount
                ddlPageNo.Items.Add(i.ToString())
            Next
            ddlPageNo.Text = CType((gvList.PageIndex + 1), String)
            lblPageCount.Text = lblPageCount.Text & gvList.PageCount.ToString()
            ddlPageSize.Text = CType(gvList.PageSize, String)
        End If
    End Sub

    'Protected Sub ddlOrgSel_DataBound(sender As Object, e As EventArgs) Handles ddlOrgSel.DataBound
    '    Dim ddl As DropDownList = CType(sender, DropDownList)
    '    ''ddl.Items.Insert(0, New ListItem("全部", "0"))
    '    If Not IsPostBack AndAlso ddl.Items.Count > 0 Then
    '        ddl.SelectedValue = org_uid
    '    Else
    '        If Not Request.Form("ddlOrgSel") = Nothing Then ddl.SelectedValue = Request.Form("ddlOrgSel")
    '    End If

    'End Sub

    'Protected Sub ddlUserSel_DataBound(sender As Object, e As EventArgs) Handles ddlUserSel.DataBound
    '    Dim ddl As DropDownList = CType(sender, DropDownList)
    '    ddl.Items.Insert(0, New ListItem("全部", "0"))
    '    If ddlOrgSel.SelectedValue = Session("org_uid") Then ddl.SelectedValue = UCase(user_id)
    'End Sub

    Protected Sub ddlRepairMainKind_DataBound(sender As Object, e As EventArgs) Handles ddlRepairMainKind.DataBound
        Dim ddl As DropDownList = CType(sender, DropDownList)
        ddl.Items.Insert(0, New ListItem("全部", "0"))
        ''If ViewState("RepairMainKind") IsNot Nothing AndAlso ViewState("RepairMainKind").ToString.Length > 0 Then ddl.SelectedValue = ViewState("RepairMainKind")
    End Sub

    Protected Sub ddlProblemKind_DataBound(sender As Object, e As EventArgs) Handles ddlProblemKind.DataBound
        Dim ddl As DropDownList = CType(sender, DropDownList)
        ddl.Items.Insert(0, New ListItem("全部", "0"))
    End Sub

    Protected Sub Searchbtn_Click(sender As Object, e As ImageClickEventArgs) Handles Searchbtn.Click
        SqlDataSource2.SelectCommand = SQLALL()
        gvList.DataBind()
    End Sub

    Protected Sub GridView1_PageIndexChanged(sender As Object, e As EventArgs) Handles gvList.PageIndexChanged
        SqlDataSource2.SelectCommand = SQLALL()
        gvList.DataBind()
    End Sub

    Protected Sub GridView1_Sorted(sender As Object, e As EventArgs) Handles gvList.Sorted
        SqlDataSource2.SelectCommand = SQLALL()
        gvList.DataBind()
    End Sub

    Protected Sub ddlPageNo_SelectedIndexChanged(sender As Object, e As EventArgs)
        Session("P11002PageNo") = CType(sender, DropDownList).Text
        gvList.PageIndex = CInt(CType(sender, DropDownList).Text) - 1
        SqlDataSource2.SelectCommand = SQLALL()
        gvList.DataBind()
    End Sub

    Protected Sub ddlPageSize_SelectedIndexChanged(sender As Object, e As EventArgs)
        Session("P11002PageSize") = CType(sender, DropDownList).Text
        gvList.PageSize = CInt(CType(sender, DropDownList).Text)
        SqlDataSource2.SelectCommand = SQLALL()
        gvList.DataBind()
    End Sub

    Protected Sub gvList_PreRender(sender As Object, e As EventArgs) Handles gvList.PreRender
        If Session("P11002PageNo") IsNot Nothing Then gvList.PageIndex = CType(Session("P11002PageNo"), Integer) - 1
        If Session("P11002PageSize") IsNot Nothing Then gvList.PageSize = CType(Session("P11002PageSize"), Integer)
    End Sub

    'Protected Sub Button1_Click(sender As Object, e As System.EventArgs) Handles Button1.Click
    '    Response.Write(" <script language='javascript'>")
    '    Response.Write(" sPath = '../inc/OrgTree.aspx';")
    '    Response.Write(" strFeatures = 'dialogWidth=900px;dialogHeight=750px;help=no;status=no;resizable=yes;scroll=no;dialogTop=100;dialogLeft=100';")
    '    Response.Write(" showModalDialog(sPath,self,strFeatures);")
    '    Response.Write(" </script>")
    'End Sub
    Protected Sub btnORGSelectOK_Click(ByVal sender As Object, ByVal e As EventArgs)

        Dim strSelectedValue As String = ""
        Dim tvSignManager As TreeView = CType(FindControl("tvOrg"), TreeView)
        Dim lbSignManager As Label = CType(FindControl("lblOrgSel"), Label)
        Dim lbSignManagerID As Label = CType(FindControl("lblOrgSelID"), Label)
        Dim tool As New C_Public

        strSelectedValue = tvSignManager.SelectedValue

        Dim arr() As String = Split(strSelectedValue, "_")
        Select Case arr(0)
            Case "EMP"
                lbSignManager.Text = tool.GetOrgNameByIDNo(arr(1)) + "-" + tool.GetUserNameByID(arr(1))
            Case "ORG"
                lbSignManager.Text = tool.GetOrgNameByID(arr(1))
        End Select
        'lbSignManager.Text = tvSignManager.SelectedNode.Text
        lbSignManagerID.Text = strSelectedValue
    End Sub
    Protected Sub btnORGSelectCancel_Click(ByVal sender As Object, ByVal e As EventArgs)

    End Sub

    Protected Sub ddlOrgSel_DataBound(sender As Object, e As EventArgs) Handles ddlOrgSel.DataBound
        CType(sender, DropDownList).Items.Insert(0, New ListItem("全部", "0"))
        CType(sender, DropDownList).SelectedValue = org_uid
    End Sub

    Protected Sub ddlOrgSel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlOrgSel.SelectedIndexChanged

        '清空User重新讀取
        ddlUserSel.Items.Clear()

        If ddlOrgSel.SelectedValue = "" Then
            SqlDataSource3.SelectCommand = "SELECT employee_id, emp_chinese_name, ORG_UID FROM EMPLOYEE WHERE 1=2"
        Else
            SqlDataSource3.SelectCommand = "SELECT employee_id, emp_chinese_name, ORG_UID FROM EMPLOYEE WHERE ORG_UID ='" & ddlOrgSel.SelectedValue & "' ORDER BY emp_chinese_name"                    
        End If
        ddlUserSel.Enabled = ddlOrgSel.Enabled
    End Sub

    Protected Sub ddlUserSel_DataBound(sender As Object, e As EventArgs) Handles ddlUserSel.DataBound
        CType(sender, DropDownList).Items.Insert(0, New ListItem("全部", "0"))
        If ddlOrgSel.SelectedValue = org_uid Then
            'ddlUserSel.SelectedValue = user_id
            For Each o As ListItem In ddlUserSel.Items
                If o.Value = UCase(user_id) Then
                    ddlUserSel.SelectedValue = UCase(user_id)
                    Exit For
                ElseIf o.Value = LCase(user_id) Then
                    ddlUserSel.SelectedValue = LCase(user_id)
                    Exit For
                End If
            Next
        Else
            ddlUserSel.SelectedValue = "0"
        End If

    End Sub
End Class
