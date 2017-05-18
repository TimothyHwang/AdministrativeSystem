Imports System.IO
Imports System.Threading

Partial Class M_Source_DownloadFile
    Inherits Page

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim filename As String = Request("file").ToString()
        fileDownload(filename, Server.MapPath("~/M_Source/11/Upload/" + filename))
    End Sub

    Private Sub fileDownload(ByVal fileName As String, ByVal fileUrl As String)

        Page.Response.Clear()
        Dim success As Boolean = ResponseFile(Page.Request, Page.Response, fileName, fileUrl, 1024000)
        If Not success Then
            Response.Write("Downloading Error!")
            Page.Response.End()
        End If
    End Sub

    Public Shared Function ResponseFile(ByVal _Request As HttpRequest, ByVal _Response As HttpResponse, ByVal _fileName As String, ByVal _fullPath As String, ByVal _speed As Long) As Boolean

        Try
            Dim myFile As New FileStream(_fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Dim br As New BinaryReader(myFile)
            Try
                _Response.AddHeader("Accept-Ranges", "bytes")
                _Response.Buffer = False
                Dim fileLength As Long = myFile.Length
                Dim startBytes As Long = 0

                Dim pack As Integer = 10240 ''10K bytes
                Dim sleep As Integer = CType(Math.Floor((1000 * pack / _speed) + 1), Integer)
                If Not IsNothing(_Request.Headers("Range")) Then
                    _Response.StatusCode = 206                    
                    Dim range As String() = Split(_Request.Headers("Range"), New [Char]() {"=", "-"})
                    startBytes = Convert.ToInt64(range(1))
                End If
                _Response.AddHeader("Content-Length", (fileLength - startBytes).ToString())
                If startBytes <> 0 Then
                    _Response.AddHeader("Content-Range", String.Format(" bytes {0}-{1}/{2}", startBytes, fileLength - 1, fileLength))
                End If
                _Response.AddHeader("Connection", "Keep-Alive")
                _Response.ContentType = "application/octet-stream"
                _Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(_fileName, System.Text.Encoding.UTF8))

                br.BaseStream.Seek(startBytes, SeekOrigin.Begin)
                Dim maxCount As Integer = CType(Math.Floor((fileLength - startBytes) / pack) + 1, Integer)
                For i As Integer = 0 To maxCount - 1
                    If (_Response.IsClientConnected) Then
                        _Response.BinaryWrite(br.ReadBytes(pack))
                        Thread.Sleep(sleep)
                    Else
                        i = maxCount
                    End If
                Next
            Catch
                Return False
            Finally
                br.Close()
                myFile.Close()
            End Try
        Catch
            Return False
        End Try
        Return True
    End Function
End Class
