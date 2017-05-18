Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Hosting
Imports System
Imports System.IO

Public Class C_Xprint
    Inherits System.Web.UI.Page
#Region "列印"
    Private head As String
    Private body As String
    Private rpt_filename As String
    Public ErrMsg As String = ""
    Private line_separator As String = Chr(13) & Chr(10)

    Public Sub C_Xprint(ByVal format_filename As String, ByVal out_filename As String)
        head = ""
        body = ""
        Me.setRptFilename(out_filename)
        NewPage(format_filename, 1)
    End Sub

    Public Sub setRptFilename(ByVal rptfilename As String)
        rpt_filename = rptfilename
    End Sub

    Public Function getRptFilename() As String
        getRptFilename = rpt_filename
    End Function

    Public Function getHead() As String
        getHead = head
    End Function

    Public Sub setHead(ByVal head As String)
        Me.head = head
    End Sub

    Public Function getBody() As String
        getBody = body
    End Function

    Public Sub setBody(ByVal body As String)
        Me.body = body
    End Sub

    Public Sub EndFile()
        Dim F_file As String

        Try
            'F_file = Server.MapPath("~/Drs/" + Me.getRptFilename())
            F_file = Server.MapPath("~/Drs/" + Me.getRptFilename().Split("/")(Me.getRptFilename().Split("/").Length - 1))
            Dim fop As Integer = FreeFile()
            FileOpen(fop, F_file, OpenMode.Output, OpenAccess.Write)

            PrintLine(fop, Me.getHead())
            PrintLine(fop, Me.getBody())
            FileClose(fop)

        Catch e As Exception
            ErrMsg = "C_Xprint.EndFile0:" & e.Message()
        Finally
            head = ""
            body = ""
        End Try
    End Sub

    Public Sub EndFile(ByVal Filename As String)
        Dim F_file As String

        Try
            F_file = Server.MapPath("~/Drs/" + Filename)
            Dim fop As Integer = FreeFile()
            FileOpen(fop, F_file, OpenMode.Output, OpenAccess.Write)

            PrintLine(fop, Me.getHead())
            PrintLine(fop, Me.getBody())
            FileClose(fop)

        Catch e As Exception
            ErrMsg = "C_Xprint.EndFile1" & e.Message()
        Finally
            head = ""
            body = ""
        End Try
    End Sub


    Private Sub NewPage(ByVal formatfilename As String, ByVal preview As Int32)
        Dim tmp_head As String
        Dim tmp_body As String

        Dim F_file As String
        Try
            tmp_head = Me.getHead()
            tmp_head += "/dtd " + formatfilename + line_separator

            F_file = Server.MapPath("~/Rpt/" + formatfilename)
            If File.Exists(F_file) = False Then
                Throw New Exception(F_file + "檔案不存在")
                Exit Sub
            End If

            Dim fop As Integer = FreeFile()
            Dim line As String
            FileOpen(fop, F_file, OpenMode.Input, OpenAccess.Read)
            Do While EOF(fop) = False
                line = LineInput(fop)
                tmp_head = tmp_head & CStr(line) & line_separator
            Loop
            FileClose(1)

            tmp_head = tmp_head & "/dtd " & formatfilename & line_separator
            Me.setHead(tmp_head)
            tmp_body = Me.getBody()
            tmp_body = tmp_body & "/init " & formatfilename & " " & preview & line_separator
            Me.setBody(tmp_body)

        Catch e As Exception
            ErrMsg = "C_Xprint.NewPage2:" & e.Message()
        End Try
    End Sub

    Public Sub NewPage(ByVal formatfilename As String)
        Dim tmp_head As String
        Dim tmp_body As String
        Dim F_file As String

        Try
            tmp_head = Me.getHead()
            tmp_head += "/dtd " + formatfilename + line_separator

            F_file = Server.MapPath("~/Rpt/" + formatfilename)
            If File.Exists(F_file) = False Then
                Throw New Exception(F_file + "檔案不存在")
                Exit Sub
            End If

            Dim fop As Integer = FreeFile()
            Dim line As String
            FileOpen(fop, F_file, OpenMode.Input, OpenAccess.Read)
            Do While EOF(fop) = False
                line = LineInput(fop)
                tmp_head = tmp_head & CStr(line) & line_separator
            Loop
            FileClose(1)

            tmp_head = tmp_head & "/dtd " & formatfilename & line_separator
            Me.setHead(tmp_head)
            tmp_body = Me.getBody()
            tmp_body = tmp_body & "/newpage " & formatfilename & line_separator
            Me.setBody(tmp_body)

        Catch e As Exception
            ErrMsg = "C_Xprint.NewPage1:" & e.Message()
        End Try
    End Sub

    Public Sub NewPage()
        Dim tmp_body As String = Me.getBody()

        tmp_body += "/newpage null" + line_separator
        Me.setBody(tmp_body)
    End Sub

    Public Sub Add(ByVal label As String, ByVal data As String, ByVal xdisp As Double, ByVal ydisp As Double)
        Dim tmp_body As String = Me.getBody()

        tmp_body = tmp_body & "/add " & label & " " & xdisp & " " & ydisp & line_separator
        tmp_body = tmp_body & data & line_separator
        tmp_body = tmp_body & "/add " & label & " " & xdisp & " " & ydisp & line_separator
        Me.setBody(tmp_body)
    End Sub

    ''' <summary>
    ''' 寫資料
    ''' </summary>
    ''' <param name="block_name">block_name 套印變數名</param>
    ''' <param name="x1">x1 座標x</param>
    ''' <param name="y1">y1 座標y</param>
    ''' <param name="block_value">欲印之值</param>
    ''' <param name="x2">x2 離原點x座標</param>
    ''' <param name="y2">y2 離原點y座標</param>
    ''' <param name="x_count">印幾次</param>
    ''' <remarks></remarks>
    Sub Add(ByVal block_name As String, ByVal x1 As String, ByVal y1 As String, ByVal block_value As String, ByVal x2 As String, ByVal y2 As String, ByVal x_count As Integer)
        Dim di As Integer
        Dim x_long As Integer = CInt(x1)
        Dim y_long As Integer = CInt(y1)
        Dim tmp_Body As String = Me.getBody

        For di = 1 To x_count
            tmp_Body = tmp_Body + "/add " + block_name + " " + x_long.ToString + " " + y_long.ToString + line_separator
            tmp_Body = tmp_Body + block_value + line_separator
            tmp_Body = tmp_Body + "/add " + block_name + " " + x_long.ToString + " " + y_long.ToString + line_separator
            x_long += CInt(x2)
            y_long += CInt(y2)
        Next
        Me.setBody(tmp_Body)
    End Sub
#End Region

End Class
