
Partial Class Source_90_FlowRoute
    Inherits System.Web.UI.UserControl

    Protected Function FunStatus(ByVal str As String) As String
        Try
            '轉換表單狀態代號
            Dim tmpStr = Eval(str)

            If tmpStr = "-" Then
                tmpStr = "申請"
            ElseIf tmpStr = "F" Then
                tmpStr = "同意修繕"
            ElseIf tmpStr = "N" Then
                tmpStr = "不修繕"
            ElseIf tmpStr = "C" Then
                tmpStr = "完工"
            ElseIf tmpStr = "0" Then
                tmpStr = "駁回"
            ElseIf tmpStr = "1" Then
                tmpStr = "送件"
            ElseIf tmpStr = "?" Then
                tmpStr = "審核中"
            ElseIf tmpStr = "E" Then
                tmpStr = "完成"
            ElseIf tmpStr = "G" Then
                tmpStr = "補登"
            ElseIf tmpStr = "B" Or tmpStr = "X" Then
                tmpStr = "申請者撤銷"
            ElseIf tmpStr = "R" Then
                tmpStr = "重新分派"
            ElseIf tmpStr = "T" Then
                tmpStr = "呈轉"
            ElseIf tmpStr = "2" Then ''資訊設備報修專用狀態，用在指定給管制單位重新分派維修單位
                tmpStr = "重派"
            Else
                tmpStr = "未知"
            End If

            FunStatus = tmpStr

        Catch ex As Exception
            FunStatus = ""
        End Try
    End Function

    Protected Sub GridView1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim lb_eformid As Label = CType(e.Row.FindControl("lb_eformid"), Label)
            If lb_eformid.Text = "74BN58683M" Then '影印使用申請單沒有批核文字
                GridView1.Columns(3).Visible = False
            End If
        End If
    End Sub
End Class
