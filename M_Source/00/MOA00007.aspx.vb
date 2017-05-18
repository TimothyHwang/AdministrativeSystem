Imports System.Data.SqlClient

Partial Class Source_00_MOA00007
    Inherits System.Web.UI.Page
    Public val As String
    Dim GetValue
    Public eformid As String
    Public eformsn As String
    Public gonogo As String
    Public stepsid As String
    Public goto_steps As String
    Public sShowMsg As String
    Public bWaitHandle As String
    Public bFlowDesignMode As Boolean = False
    Public bFlowSnNotFound As Boolean = False
    Public do_sql As New C_SQLFUN
    Public newx As Integer
    Public newy As Integer
    Public backto As String
    Public backtomsg As String
    Public bDesignColumnMode As Boolean = False
    Public P_05RoomManager As String
    Public P_09RoomManager As String
    '---
    Public sFlowTable As String
    Public sCCTable As String
    Public sEformTable As String
    Public NodesEndPosX As Integer
    Public eformrole As String
    Public NextStep As String
    Public minx As String
    Public maxx As String
    Public startX As Integer
    Public endX As Integer
    Public minnextstep As String
    Public maxnextstep As String
    Public orgnextstep As String
    '---
    Dim dr As System.Data.DataRow
    Dim n_table As New System.Data.DataTable
    Dim orgx As Integer
    Dim orgy As Integer
    Dim user_id, org_uid As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            user_id = Session("user_id")
            org_uid = Session("ORG_UID")

            'stop
            '
            'dim backto,backtomsg
            'backto="uListToDoTask.asp"
            'backtomsg="回未處理表單"
            'eformid="w3Rr68Mi1d"
            'eformsn="59KK8MCGPRSIM594"
            'gonogo="1"
            'stepsid="27218"
            'sShowMsg="您所填寫的「註銷申請單」已經傳送給下一關的處理者單位資產管理員：王崑仰"
            'bWaitHandle="0"
            'goto_steps="6534" 

            'str_EFORMSN = d_pub.randstr(16)
            'val = "w3Rr68Mi1d,59KK8MCGPRSIM594,1,27218,6534,您所填寫的「註銷申請單」已經傳送給下一關的處理者單位資產管理員：王崑仰,1"

            'val = "U28r13D6EA,8I8HK7LA15Y9662O,1,1,29683,您所填寫的「註銷申請單」已經傳送給下一關的處理者單位資產管理員：王崑仰,1"
            'val = "U28r13D6EA,8I8HK7LA15Y9662O,1,19792,-1,您所填寫的「註銷申請單」已經傳送給下一關的處理者單位資產管理員：王崑仰,1"
            'val = "YAqBTxRP8P,18Z63725MJ5H618C,1,28093,-1,成功,1"
            'val =  eformid     eformsn    gonogo stepsid goto_steps      sShowMsg

            val = Request("val")

            If val = "" Then

                '判斷登入者權限
                Dim LoginCheck As New C_Public

                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00007.aspx")
                Response.End()

            Else

                GetValue = Split(val, ",")
                eformid = GetValue(0)
                eformsn = GetValue(1)
                gonogo = GetValue(2)
                stepsid = GetValue(3)
                goto_steps = GetValue(4)

                sShowMsg = GetValue(5)

                '去除身分證字號
                If InStr(1, sShowMsg, "(") > 0 Then
                    sShowMsg = Left(sShowMsg, InStr(1, sShowMsg, "(") - 1)
                Else
                    sShowMsg = sShowMsg
                End If

                eformrole = GetValue(6)
                bWaitHandle = "0"

                'bFlowDesignMode = False
                'Dim sString1, sString2, bFlowSnNotFound
                'bFlowSnNotFound = False
                backto = "uListToDoTask.asp"  '錯誤
                backtomsg = "回未處理表單"

            End If

        Catch ex As Exception

        End Try

    End Sub
    Sub first_step()
        Dim sqlstrQT As String

        If eformid = "U28r13D6EA" Then
            sqlstrQT = "select b.* from p_05 a,syskind b where eformsn='" & eformsn & "' and a.nrecroom=b.state_name "
            If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                Exit Sub
            End If
            If do_sql.G_table.Rows.Count > 0 Then
                P_05RoomManager = do_sql.G_table.Rows(0).Item("State_Desc")
            End If
        End If



        sqlstrQT = "select x,y,nextstep,allhandle from flow where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & stepsid
        If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
            Exit Sub
        End If
        Dim allhandle As String

        If do_sql.G_table.Rows.Count > 0 Then
            orgx = CInt(do_sql.G_table.Rows(0).Item("x").ToString)
            orgy = CInt(do_sql.G_table.Rows(0).Item("y").ToString)
            orgnextstep = do_sql.G_table.Rows(0).Item("nextstep").ToString
            allhandle = do_sql.G_table.Rows(0).Item("allhandle").ToString
        End If


        If goto_steps > "" Then
            orgnextstep = goto_steps
        End If
        If bWaitHandle = "1" Then
            Response.Write("<img ID=m" & stepsid & " STYLE='position:absolute;top:" & orgy - 30 & "; left:" & orgx + 20 & "; z-index:6' SRC='../../image/mail3.gif'>")
            Response.Write("<span id=disp style='position:absolute;top:" & orgy - 10 & "px;left:" & orgx + 100 & "px;width:250px;border:solid 1px silver;padding:10px;z-index:200;background-color:#FEF1E2;color:#000080'>")
            Response.Write(sShowMsg)
            Response.Write("</span>")
        ElseIf gonogo = "-" Or gonogo = "1" Or gonogo = "0" Or gonogo = "W" Then ' 申請或送件或駁回或回收

            sqlstrQT = "select x,y,stepsid from flow where eformid='" & eformid & "' and steps=" & orgnextstep
            If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                Exit Sub
            End If

            newx = 0
            newy = 0

            n_table = do_sql.G_table

            For Each dr In n_table.Rows
                Response.Write("<img ID=m" & Trim(dr("stepsid").ToString) & " STYLE='position:absolute;top:" & (orgy - 12) & "; left:" & orgx + 20 & "; z-index:6;visibility:hidden;' SRC='../../image/mail3.gif'>")
                If CInt(Trim(dr("x").ToString)) > newx Then
                    newx = CInt(Trim(dr("x").ToString))
                End If

                If CInt(Trim(dr("y").ToString)) > newy Then
                    newy = CInt(Trim(dr("y").ToString))
                End If
                'Exit For
            Next
            '設定多關簽核時每一張圖片的移動起始位置
            Response.Write("<script language=""javascript"" type=""text/javascript"">")
            SetDefaultPicPosition(orgnextstep, orgx, orgy)
            Response.Write("</script>")

            Response.Write("<span id=disp style='position:absolute;top:" & newy - 10 & "px;left:" & newx + 100 & "px;width:250px;border:solid 1px silver;padding:10px;z-index:200;background-color:#FEF1E2;color:#000080'>")
            Response.Write(sShowMsg)
            Response.Write("</span>")

            Dim PageUp As String = Request("PageUp")

            If PageUp = "New" Then
            Else
                '重新整理頁面
                Response.Write(" <script language='javascript'>")
                Response.Write(" window.dialogArguments.location='../00/MOA00010.aspx';")
                'Response.Write(" window.open('../00/MOA00010.aspx');")
                Response.Write(" </script>")
            End If

        End If
    End Sub
    Sub second_step()
        Dim sqlstrQT As String

        If bWaitHandle = "1" Then '等候其他人處理
            ' 原地不動
            Response.Write("window.scrollTo(100," & orgy - 220 & ");" & Chr(13))
        ElseIf orgy = newy Then
            ' 原地不動
            Response.Write("window.scrollTo(100," & orgy - 220 & ");" & Chr(13))
        Else
            Dim timerGap As Integer
            Dim sString1 As String
            Dim sString2 As String
            timerGap = Math.Abs(12000 / (orgy - newy))

            If eformid = "U28r13D6EA" Then
                sqlstrQT = "select count(*) as reccount from flow where eformid='" & eformid & "' and eformrole=" & eformrole & " and steps=" & orgnextstep
                If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                    Exit Sub
                End If
                If do_sql.G_table.Rows(0)("reccount") > 1 Then
                    sqlstrQT = "select x,y,stepsid from flow where eformid='" & eformid & "' and eformrole=" & eformrole & " and steps=" & orgnextstep & " and group_id='" & P_05RoomManager & "'"
                Else
                    sqlstrQT = "select x,y,stepsid from flow where eformid='" & eformid & "' and eformrole=" & eformrole & " and steps=" & orgnextstep
                End If
                If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                    Exit Sub
                End If
            Else
                sqlstrQT = "select x,y,stepsid from flow where eformid='" & eformid & "' and eformrole=" & eformrole & " and steps=" & orgnextstep
                If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                    Exit Sub
                End If
            End If

            If do_sql.G_table.Rows.Count = 0 Then
                'Response.write  ("location.href=""errormsg.asp?reason=" & "找不到所要傳送的流程關卡" & "&source=無&desc=" & sqlcommand & "&path=" & server.urlencode(Request.ServerVariables("url")) & """;" & chr(13))
            End If

            sString1 = ""
            sString2 = ""
            n_table = do_sql.G_table

            ''next_x
            ''指定原始起始坐標
            'Response.Write("m" + Trim(dr("stepsid").ToString) + ".style.pixelLeft=" + CStr(orgx) + ";")
            'Response.Write("m" + Trim(dr("stepsid").ToString) + ".style.pixelTop=" + CStr(orgy) + ";")

            Response.Write("function Next1(){")
            Dim vstepsid As String
            Dim mstepsid As String
            For Each dr In n_table.Rows

                vstepsid = "v" & Trim(dr("stepsid").ToString)
                mstepsid = "m" & Trim(dr("stepsid").ToString)
                Response.Write(mstepsid & ".style.visibility='visible'" & Chr(13))
                Response.Write("var " & vstepsid & " = 0;" & Chr(13))
                Response.Write("if (" & mstepsid & ".style.pixelLeft < " & CInt(Trim(dr("x").ToString)) + 20 & ") {" & Chr(13))
                Response.Write(mstepsid & ".style.pixelLeft += 3;" & Chr(13))
                Response.Write("    }" & Chr(13))
                Response.Write("if (" & mstepsid & ".style.pixelLeft > " & CInt(Trim(dr("x").ToString)) + 20 & ") {" & Chr(13))
                Response.Write(mstepsid & ".style.pixelLeft -= 3;" & Chr(13))
                Response.Write("    }" & Chr(13))
                Response.Write("if (" & mstepsid & ".style.pixelTop < " & CInt(Trim(dr("y").ToString)) - 30 & ") {" & Chr(13))
                Response.Write(mstepsid & ".style.pixelTop += 3;" & Chr(13))
                Response.Write("    }" & Chr(13))
                Response.Write("if (" & mstepsid & ".style.pixelTop > " & CInt(Trim(dr("y").ToString)) - 30 & ") {" & Chr(13))
                Response.Write(mstepsid & ".style.pixelTop -= 3;" & Chr(13))
                Response.Write("    }" & Chr(13))
                Response.Write("if (Math.abs(" & mstepsid & ".style.pixelTop - " & CInt(Trim(dr("y").ToString)) - 30 & ") <= 3 && Math.abs(" & mstepsid & ".style.pixelLeft - " & CInt(Trim(dr("x").ToString)) + 19 & ") <=3 ){" & Chr(13))
                Response.Write(vstepsid & "= 1;" & Chr(13))
                Response.Write("    }" & Chr(13))
                If sString1 > "" Then
                    sString1 = sString1 & ","
                End If
                sString1 = sString1 & mstepsid & ".style.pixelTop"
                If sString2 > "" Then
                    sString2 = sString2 & "||"
                End If
                sString2 = sString2 & vstepsid & " != 1"

                Exit For

            Next

            If sString1 > "" Then
                Response.Write("var maxY = Math.max(0 , " & sString1 & ");" & Chr(13))
                Response.Write("window.scrollTo(100,maxY-220);" & Chr(13))
                Response.Write("if ( " & sString2 & ") {" & Chr(13))
                Response.Write("   setTimeout('Next1()'," & timerGap & ");" & Chr(13))
                Response.Write("}" & Chr(13))
                Response.Write("else {" & Chr(13))
                Response.Write("   KBStatic.style.display=''" & Chr(13))
                Response.Write("}" & Chr(13))
            End If
            Response.Write("}" & Chr(13))
            Response.Write("   setTimeout('Next1()'," & timerGap & ");")
        End If

    End Sub

    Public Function GetGroupid(ByVal eformsn As String) As String
        Dim sReturn As String = ""
        Dim db As New SqlConnection
        db.ConnectionString = do_sql.G_conn_string
        db.Open()
        Dim DR As SqlDataReader = New SqlCommand("select c.* from p_05 a,syskind b,flow c where eformsn='" & eformsn & "' and a.nrecroom=b.state_name and c.group_id=b.state_desc", db).ExecuteReader
        If DR.HasRows Then
            If DR.Read Then
                sReturn = DR("group_id")
            End If
        End If
        db.Close()
        GetGroupid = sReturn
    End Function

    Public Sub SetDefaultPicPosition(ByVal steps As String, ByVal pX As String, ByVal pY As String)
        Dim db As New SqlConnection
        db.ConnectionString = do_sql.G_conn_string
        db.Open()
        Dim DR As SqlDataReader = New SqlCommand("select stepsid from flow where eformid='" & eformid & "' and eformrole=" & eformrole & " and steps=" & steps, db).ExecuteReader
        If DR.HasRows Then
            While DR.Read
                Response.Write("m" + Trim(DR("stepsid").ToString) + ".style.pixelLeft=" + pX + ";" + Chr(13))
                Response.Write("m" + Trim(DR("stepsid").ToString) + ".style.pixelTop=" + pY + ";" + Chr(13))
            End While
        End If
        db.Close()

    End Sub
End Class
