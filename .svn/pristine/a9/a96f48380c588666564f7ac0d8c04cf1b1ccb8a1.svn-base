Imports System.Data.SqlClient

Partial Class M_Source_10_MOA10002
    Inherits Page

    ReadOnly sql_function As New C_SQLFUN
    ReadOnly connstr As String = sql_function.G_conn_string
    Dim db As New SqlConnection(connstr)
#Region "Custom Function"
    ''**********************  以下為Custom Function  **************************
    Protected Function GetNewRANKID(ByVal fieldname As String,ByVal db As SqlConnection) As Integer
        Dim iReturn As Integer
        db.Open()
        Dim strSql = "SELECT MAX(RANKID)+1 AS RANKID FROM " + fieldname
        Dim DR As SqlDataReader
        DR = New SqlCommand(strSql, db).ExecuteReader
        If DR.HasRows Then
            If DR.Read() Then
                If DR("RANKID").ToString() = "" Then
                    iReturn = 1
                Else
                    iReturn = CInt(DR("RANKID").ToString())
                End If
            End If
        Else
            iReturn = 1
        End If

        db.Close()
        GetNewRANKID = iReturn
    End Function
#End Region

#Region "Form Function"
    ''**********************  以下為  Form Function  **************************
    Protected Sub ibnAddOK_Click(ByVal sender As Object, ByVal e As ImageClickEventArgs)
	
        ''新增資料表資料
        If dvManager.CurrentMode = DetailsViewMode.Insert Then

            Dim txtTitle As TextBox = CType(dvManager.FindControl("txtTitle"), TextBox)
            Dim txtName As TextBox = CType(dvManager.FindControl("txtName"), TextBox)
            Dim txtRankID As TextBox = CType(dvManager.FindControl("txtRankId"), TextBox)

            Dim strSQL As String = "INSERT INTO P_10 (Name,RankId,Status,CreateDate,Creator) VALUES ("
            'strSQL += "'" & IIf(txtTitle Is Nothing, "", txtTitle.Text).ToString() & "'"
            strSQL += "'" & IIf(txtName Is Nothing, "", txtName.Text).ToString() & "'"
            'strSQL += ",'" & IIf(txtRankID Is Nothing, "", txtRankID.Text).ToString() & "'"
            strSQL += "," + CType(GetNewRANKID("P_10", db), String)		
            strSQL += ",0"            
            strSQL += ",GetDate(),'" & Session("user_id").ToString() & "'"
            strSQL += ")"

            db.Open()
            Call New SqlCommand(strSQL, db).ExecuteNonQuery()
            MessageBox.Show("新增完成！！")
            db.Close()
            Server.Transfer("MOA10001.aspx")
        End If
    End Sub

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not IsPostBack Then
            dvManager.ChangeMode(DetailsViewMode.Insert)
        End If
        ''繫結Header
        Page.Header.DataBind()
    End Sub

    Protected Sub ibnCancel_Click(sender As Object, e As ImageClickEventArgs)
        Server.Transfer("MOA10001.aspx")
    End Sub
#End Region

End Class
