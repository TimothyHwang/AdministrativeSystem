
Namespace M_Source._11
    Partial Class M_Source_11_MOA11004
        Inherits Page

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
                sReturn += " AND (FIXSTATUS='" & ddlStatusSel.SelectedValue & "'  OR FIXSTATUS IS NULL)"
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
    End Class
End Namespace