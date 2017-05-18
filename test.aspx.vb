
Imports System.Data.SqlClient

Partial Class test
    Inherits Page

    Protected Sub Button1_Click(sender As Object, e As System.EventArgs) Handles Button1.Click
        Dim tool As New C_Public
        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString2").ConnectionString)
            conn.Open()
            Dim strSql As String = ""
            strSql = "SELECT TOP 1 * FROM UPLOAD ORDER BY UPLOAD_ID DESC"
            Dim DR As SqlDataReader
            Dim dt As DateTime
            Dim cmd As New SqlCommand(strSql, conn)
            DR = cmd.ExecuteReader
            If DR.HasRows Then
                If DR.Read() Then
                    If CType(DR("data"), Byte()).Length > 0 Then
                        tool.ConvertByteArrayToFile("/MOA/M_Source/99/temp/", Split((DR("filename").ToString()), "-")(1), CType(DR("data"), Byte()))
                    End If
                End If
            End If
        End Using
    End Sub
End Class
