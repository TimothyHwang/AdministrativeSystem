
Imports System.Data.SqlClient

Partial Class M_Source_10_MOA10007
    Inherits Page

    ReadOnly sql_function As New C_SQLFUN
    ReadOnly connstr As String = sql_function.G_conn_string
    Dim db As New SqlConnection(connstr)

#Region "Custom Function"
    ''**********************  以下為Custom Function  **************************
#End Region

#Region "Form Function"
    ''**********************  以下為  Form Function  **************************
    Protected Sub dvP_10Mqrquee_DataBound(sender As Object, e As EventArgs) Handles dvP_10Mqrquee.DataBound
        Dim dv As DetailsView = CType(sender, DetailsView)
        If dv IsNot Nothing Then
            If dv.CurrentMode = DetailsViewMode.ReadOnly Then
                Dim lblMessage As Label = CType(dv.FindControl("lblMarquee"), Label)
                'lblMessage.Text = lblMessage.Text.Replace("\r\n", "<br/>").Replace(vbCrLf, "<br/>").Replace(" ", "&nbsp;")
            ElseIf dv.CurrentMode = DetailsViewMode.Edit Then
                Dim txtMessage As TextBox = CType(dv.FindControl("txtMarquee"), TextBox)
                'txtMessage.Text = txtMessage.Text.Replace("\r\n", vbNewLine).Replace(vbCrLf, vbNewLine)
                'txtMessage.Text = txtMessage.Text.Replace("\r\n", vbNewLine)

            End If
        End If
    End Sub

    Protected Sub dvP_10Mqrquee_ItemCommand(sender As Object, e As DetailsViewCommandEventArgs) Handles dvP_10Mqrquee.ItemCommand
        Try
            Dim sUpdateString As String = ""
            Dim i As Integer
            Dim tool As New C_Public
            Select Case e.CommandName
                Case "Update"
                    Dim pk_index As Integer = CInt(e.CommandArgument)

                    Dim txtMarquee As TextBox = CType(dvP_10Mqrquee.FindControl("txtMarquee"), TextBox)
                    'Dim arrMarquee() As String = txtMarquee.Text.Split(vbNewLine)

                    'For Each txtLine As String In arrMarquee
                    '    'If sUpdateString.Length > 0 Then sUpdateString += vbNewLine
                    '    If txtLine.Length > 20 Then
                    '        Dim oDivRemainder As Integer = 0
                    '        Dim oDivResult As Integer = Math.DivRem(txtLine.Length, 20, oDivRemainder)
                    '        For i = 0 To oDivResult - 1
                    '            sUpdateString += txtLine.Substring(i * 20, 20) & vbCrLf
                    '        Next
                    '        If oDivRemainder > 0 Then
                    '            sUpdateString += txtLine.Substring(i * 20, oDivRemainder)
                    '        End If
                    '        ''Dim arrMarquee As String() = txtMarquee.Text.Split(CType(vbCrLf, Char))
                    '        ''Dim strNewMarquee As String = ""
                    '        ''For Each s In arrMarquee
                    '        ''    If s.Length > 30 Then
                    '        ''        strNewMarquee += s.ToString.Insert(30, vbCrLf)
                    '        ''    Else
                    '        ''        strNewMarquee += s
                    '        ''    End If
                    '        ''Next
                    '        ''txtMarquee.Text = strNewMarquee
                    '    Else
                    '        'sUpdateString = txtMarquee.Text
                    '        sUpdateString += txtLine
                    '    End If
                    'Next





                    If txtMarquee.Text.Length > 20 Then
                        Dim oDivRemainder As Integer = 0
                        Dim oDivResult As Integer = Math.DivRem(txtMarquee.Text.Length, 20, oDivRemainder)
                        For i = 0 To oDivResult - 1
                            sUpdateString += txtMarquee.Text.Substring(i * 20, 20) & vbCrLf
                        Next
                        If oDivRemainder > 0 Then
                            sUpdateString += txtMarquee.Text.Substring(i * 20, oDivRemainder)
                        End If
                        ''Dim arrMarquee As String() = txtMarquee.Text.Split(CType(vbCrLf, Char))
                        ''Dim strNewMarquee As String = ""
                        ''For Each s In arrMarquee
                        ''    If s.Length > 30 Then
                        ''        strNewMarquee += s.ToString.Insert(30, vbCrLf)
                        ''    Else
                        ''        strNewMarquee += s
                        ''    End If
                        ''Next
                        ''txtMarquee.Text = strNewMarquee
                    Else
                        sUpdateString = txtMarquee.Text
                    End If


                    ''判斷公告字數是否超過資料庫欄位大小
                    If tool.ChineseStringLenth(sUpdateString) > 100 Then
                        'MessageBox.Show("公告字數超過資料庫大小，請修正內容！")
                        'Exit Sub
                        Response.Write("<script type=""text/javascript"" language=""javascript"">alert('公告字數超過資料庫大小，請修正內容！');</script>")
                        Server.Transfer("MOA10007.aspx")
                    End If

                    db.Open()
                    'Dim strSql As String = "UPDATE SYSCONFIG SET CONFIG_VALUE='" & txtMarquee.Text & "' WHERE CONFIG_NUM='" & pk_index & "'"
                    Dim strSql As String = "UPDATE SYSCONFIG SET CONFIG_VALUE='" & sUpdateString & "' WHERE CONFIG_NUM='" & pk_index & "'"
                    Call New SqlCommand(strSql, db).ExecuteNonQuery()
                    db.Close()
            End Select
        Catch ex As Exception
            MessageBox.Show("公告字數超過資料庫大小，請修正內容！")
        End Try

    End Sub
#End Region

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        Page.Header.DataBind()
    End Sub
End Class
