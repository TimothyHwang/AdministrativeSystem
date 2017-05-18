
Partial Class Source_00_MOA00006
    Inherits System.Web.UI.Page
    Public do_sql As New C_SQLFUN
    Public redirect_flag As Boolean = False
    Dim user_id, org_uid As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            user_id = Session("user_id")
            org_uid = Session("ORG_UID")

            'MsgBox("xxx", 0, "")
            'On Error Resume Next
            ' program start here
            Dim scrollLeftPixel, scrollTopPixel
            'Dim objUniversal_r, objUniversal_w
            Dim eformid As String
            Dim stepsid As String
            Dim frm_chinese_name As String
            Dim group_id As String
            Dim organization_id As String
            Dim group_name As String
            Dim sqlcommand As String
            Dim eformrole As String

            eformid = Request("eformid")
            stepsid = Request("STEPSID")
            frm_chinese_name = Request("frm_chinese_name")
            group_id = Request("group_id")
            organization_id = Request("organization_id")
            scrollLeftPixel = Request("scrollLeftPixel")
            scrollTopPixel = Request("scrollTopPixel")
            group_name = Request("group_name") '更名時新的 group name
            eformrole = Request("eformrole")

            If eformid = "" Then

                '判斷登入者權限
                Dim LoginCheck As New C_Public

                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00006.aspx")
                Response.End()

            Else
                Select Case Request("MODE")
                    Case "NEW" '重新規劃
                        If do_sql.begin_tran(do_sql.G_conn_string) = False Then
                            GoTo go_back
                            Exit Sub
                        End If
                        sqlcommand = "delete from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " and steps not in(0,-1)"
                        If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                            Call do_sql.rollback_tran()
                            GoTo go_back
                            Exit Sub
                        End If
                        sqlcommand = "update tempflow set nextstep=-1,y=200,x=180 where eformid='" & eformid & "' and eformrole=" & eformrole & " and steps =0"
                        If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                            Call do_sql.rollback_tran()
                            GoTo go_back
                            Exit Sub
                        End If
                        sqlcommand = "update tempflow set y=100,x=180 where eformid='" & eformid & "' and eformrole=" & eformrole & " and steps =-1" '結束
                        If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                            Call do_sql.rollback_tran()
                            GoTo go_back
                            Exit Sub
                        End If
                        sqlcommand = "delete from tempcc where eformid='" & eformid & "' and eformrole=" & eformrole
                        If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                            Call do_sql.rollback_tran()
                            GoTo go_back
                            Exit Sub
                        End If
                        Call do_sql.commit_tran()
                    Case "PROPERTY" '設定step屬性   
                        sqlcommand = "update tempflow set opinonly='" & Request("opinonly") & "'"
                        sqlcommand += ",allhandle='" & Request("allhandle") & "'"
                        sqlcommand += ",cansendmail='" & Request("cansendmail") & "'"
                        If Request("overduenotice") > "" And Request("overduenotice") < "2" Then
                            sqlcommand = sqlcommand & ",overduenotice='" & Request("overduenotice") & "'"
                        End If
                        If Request("canjump") > "" Then
                            sqlcommand = sqlcommand & ",canjump='" & Request("canjump") & "'"
                        Else
                            sqlcommand = sqlcommand & ",canjump='0'"
                        End If
                        If Request("canback") > "" Then
                            sqlcommand = sqlcommand & ",canback='" & Request("canback") & "'"
                            If Request("canback") = "3" Then
                                sqlcommand = sqlcommand & ",backto_steps=" & Request("backto_steps")
                            Else
                                sqlcommand = sqlcommand & ",backto_steps=null"
                            End If
                        Else
                            sqlcommand = sqlcommand & ",canback='0'"
                        End If
                        If Request("canadd") > "" Then
                            sqlcommand = sqlcommand & ",canadd='" & Request("canadd") & "'"
                        Else
                            sqlcommand = sqlcommand & ",canadd='0'"
                        End If
                        sqlcommand = sqlcommand & ",candist='" & Request("candist") & "'"
                        If Request("candist") = "4" Then
                            sqlcommand = sqlcommand & ",assign_column_id='" & Request("assign_column_id") & "'"
                        Else
                            sqlcommand = sqlcommand & ",assign_column_id=null"
                        End If
                        sqlcommand = sqlcommand & ",candistcc='" & Request("candistcc") & "'"
                        If Request("candistcc") = "4" Then
                            sqlcommand = sqlcommand & ",assigncc_column_id='" & Request("assigncc_column_id") & "'"
                        Else
                            sqlcommand = sqlcommand & ",assigncc_column_id=null"
                        End If
                        If Request("canaddatt") > "" Then
                            sqlcommand = sqlcommand & ",canaddatt='" & Request("canaddatt") & "'"
                        Else
                            sqlcommand = sqlcommand & ",canaddatt='0'"
                        End If
                        If Request("candelatt") > "" Then
                            sqlcommand = sqlcommand & ",candelatt='" & Request("candelatt") & "'"
                        Else
                            sqlcommand = sqlcommand & ",candelatt='0'"
                        End If
                        If Request("attachlimit") > "" Then
                            sqlcommand = sqlcommand & ",attachlimit=" & Request("attachlimit")
                        End If
                        If Request("canconti") > "" Then
                            sqlcommand = sqlcommand & ",canconti='" & Request("canconti") & "'"
                        Else
                            sqlcommand = sqlcommand & ",canconti='0'"
                        End If
                        If Request("bypass") > "" Then
                            sqlcommand = sqlcommand & ",bypass='" & Request("bypass") & "'"
                        Else
                            sqlcommand = sqlcommand & ",bypass='0'"
                        End If
                        sqlcommand = sqlcommand & ",backcanassign='" & Request("backcanassign") & "'"
                        If Request("backcanassign") = "4" Then
                            sqlcommand = sqlcommand & ",backassign_column='" & Request("backassign_column") & "'"
                        Else
                            sqlcommand = sqlcommand & ",backassign_column=null"
                        End If
                        If Request("allowtempsave") > "" Then
                            sqlcommand = sqlcommand & ",allowtempsave='" & Request("allowtempsave") & "'"
                        Else
                            sqlcommand = sqlcommand & ",allowtempsave='0'"
                        End If
                        If Request("allowgotoend") > "" Then
                            sqlcommand = sqlcommand & ",allowgotoend='" & Request("allowgotoend") & "'"
                        Else
                            sqlcommand = sqlcommand & ",allowgotoend='0'"
                        End If
                        If Request("allowdutyforgotoend") > "" Then
                            sqlcommand = sqlcommand & ",allowdutyforgotoend='" & Request("allowdutyforgotoend") & "'"
                        Else
                            sqlcommand = sqlcommand & ",allowdutyforgotoend='0'"
                        End If
                        If Request("gotoendcc") > "" Then
                            sqlcommand = sqlcommand & ",gotoendcc='" & Request("gotoendcc") & "'"
                        Else
                            sqlcommand = sqlcommand & ",gotoendcc='0'"
                        End If
                        If Request("canskip") > "" Then
                            sqlcommand = sqlcommand & ",canskip='" & Request("canskip") & "'"
                        Else
                            sqlcommand = sqlcommand & ",canskip='0'"
                        End If

                        sqlcommand = sqlcommand & " where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & stepsid
                        If do_sql.db_exec(sqlcommand, do_sql.G_conn_string) = False Then
                            GoTo go_back
                            Exit Sub
                        End If
                        'Case "SAVE" '儲存設定
                        '    If do_sql.begin_tran(do_sql.G_conn_string) = False Then
                        '        Exit Sub
                        '    End If
                        '    sqlcommand = "delete from flow where eformid='" & eformid & "' and eformrole=" & eformrole
                        '    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                        '        Call do_sql.rollback_tran()
                        '        Response.Redirect("errormsg.asp?reason=" & "無法儲存流程屬性設定" & "&source=" & sqlcommand & "&desc=" & do_sql.G_errmsg)
                        '        Exit Sub
                        '    End If
                        '    sqlcommand = "delete from cc where eformid='" & eformid & "' and eformrole=" & eformrole
                        '    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                        '        Call do_sql.rollback_tran()
                        '        Response.Redirect("errormsg.asp?reason=" & "無法儲存流程屬性設定" & "&source=" & sqlcommand & "&desc=" & do_sql.G_errmsg)
                        '        Exit Sub
                        '    End If
                        '    sqlcommand = "insert into flow select * from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole
                        '    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                        '        Call do_sql.rollback_tran()
                        '        Response.Redirect("errormsg.asp?reason=" & "無法儲存流程屬性設定" & "&source=" & sqlcommand & "&desc=" & do_sql.G_errmsg)
                        '        Exit Sub
                        '    End If
                        '    sqlcommand = "insert into cc select * from tempcc where eformid='" & eformid & "' and eformrole=" & eformrole
                        '    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                        '        Call do_sql.rollback_tran()
                        '        Response.Redirect("errormsg.asp?reason=" & "無法儲存流程屬性設定" & "&source=" & sqlcommand & "&desc=" & do_sql.G_errmsg)
                        '        Exit Sub
                        '    End If
                        '    sqlcommand = "delete from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole
                        '    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                        '        Call do_sql.rollback_tran()
                        '        Response.Redirect("errormsg.asp?reason=" & "無法儲存流程屬性設定" & "&source=" & sqlcommand & "&desc=" & do_sql.G_errmsg)
                        '        Exit Sub
                        '    End If
                        '    sqlcommand = "delete from tempcc where eformid='" & eformid & "' and eformrole=" & eformrole
                        '    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                        '        Call do_sql.rollback_tran()
                        '        Response.Redirect("errormsg.asp?reason=" & "無法儲存流程屬性設定" & "&source=" & sqlcommand & "&desc=" & do_sql.G_errmsg)
                        '        Exit Sub
                        '    End If
                        '    sqlcommand = "delete from tempeforms where eformid='" & eformid & "'"
                        '    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                        '        Call do_sql.rollback_tran()
                        '        Response.Redirect("errormsg.asp?reason=" & "無法儲存流程屬性設定" & "&source=" & sqlcommand & "&desc=" & do_sql.G_errmsg)
                        '        Exit Sub
                        '    End If
                        '    Call do_sql.commit_tran()
                        'Case "CANCEL" ' 取消變更
                        '    If do_sql.begin_tran(do_sql.G_conn_string) = False Then
                        '        Exit Sub
                        '    End If
                        '    sqlcommand = "delete from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole
                        '    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                        '        Call do_sql.rollback_tran()
                        '        Response.Redirect("errormsg.asp?reason=" & "無法取消流程變更" & "&source=" & sqlcommand & "&desc=" & do_sql.G_errmsg)
                        '        Exit Sub
                        '    End If
                        '    sqlcommand = "delete from tempcc where eformid='" & eformid & "' and eformrole=" & eformrole
                        '    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                        '        Call do_sql.rollback_tran()
                        '        Response.Redirect("errormsg.asp?reason=" & "無法取消流程變更" & "&source=" & sqlcommand & "&desc=" & do_sql.G_errmsg)
                        '        Exit Sub
                        '    End If
                        '    sqlcommand = "delete from tempeforms where eformid='" & eformid & "'"
                        '    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                        '        Call do_sql.rollback_tran()
                        '        Response.Redirect("errormsg.asp?reason=" & "無法取消流程變更" & "&source=" & sqlcommand & "&desc=" & do_sql.G_errmsg)
                        '        Exit Sub
                        '    End If
                        '    Call do_sql.commit_tran()
                    Case "CC" '通知
                        sqlcommand = "insert into tempcc(eformid, eformrole,stepsid, group_id,group_type) values ('"
                        sqlcommand += eformid & "'," & eformrole & "," & stepsid & ",'" & Left(group_id, InStr(1, group_id, "_") - 1) & "','" & Mid(group_id, InStr(1, group_id, "_") + 1) & "')"
                        If do_sql.db_exec(sqlcommand, do_sql.G_conn_string) = False Then
                            GoTo go_back
                            Exit Sub
                        End If

                    Case "DELCC" '刪除通知
                        Dim cc As String
                        cc = Request("cc")
                        sqlcommand = "delete tempcc where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & stepsid & " and group_id='" & cc & "'"
                        If do_sql.db_exec(sqlcommand, do_sql.G_conn_string) = False Then
                            GoTo go_back
                            Exit Sub
                        End If
                    Case "CHANGENAME" '更名
                        sqlcommand = "update tempflow set alias_group_name='" & group_name & "' where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & stepsid
                        If do_sql.db_exec(sqlcommand, do_sql.G_conn_string) = False Then
                            GoTo go_back
                            Exit Sub
                        End If
                    Case "MAJOR" '設定為主關卡                
                        Dim sqlstrQT As String
                        Dim oldSteps As String = ""
                        sqlstrQT = "select * from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & stepsid
                        If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                            GoTo go_back
                            Exit Sub
                        End If
                        If do_sql.G_table.Rows.Count > 0 Then
                            oldSteps = do_sql.G_table.Rows(0).Item("steps").ToString.Trim
                        Else
                            GoTo go_back
                            Exit Sub
                        End If
                        If do_sql.begin_tran(do_sql.G_conn_string) = False Then
                            Exit Sub
                        End If
                        sqlcommand = "update tempflow set major_step='0' where eformid='" & eformid & "' and eformrole=" & eformrole & " and steps=" & oldSteps & ""
                        If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                            Call do_sql.rollback_tran()
                            GoTo go_back
                            Exit Sub
                        End If

                        sqlcommand = "update tempflow set major_step='1' where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & stepsid & ""
                        If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                            Call do_sql.rollback_tran()
                            GoTo go_back
                            Exit Sub
                        End If
                        Call do_sql.commit_tran()
                    Case "ASER" '串簽
                        Dim sqlstrQT As String
                        Dim oldNextStep As String = ""
                        Dim oldX As String = ""
                        Dim oldY As String = ""
                        Dim oldblock As String = ""
                        Dim oldlayer As String = ""
                        Dim oldtree As String = ""
                        sqlstrQT = "select * from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & stepsid
                        If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                            GoTo go_back
                            Exit Sub
                        End If
                        If do_sql.G_table.Rows.Count = 0 Then
                            GoTo go_back
                            Exit Sub
                        End If
                        oldNextStep = do_sql.G_table.Rows(0).Item("nextstep").ToString.Trim
                        oldX = do_sql.G_table.Rows(0).Item("x").ToString.Trim
                        oldY = do_sql.G_table.Rows(0).Item("y").ToString.Trim
                        oldblock = do_sql.G_table.Rows(0).Item("block").ToString.Trim
                        oldlayer = do_sql.G_table.Rows(0).Item("layer").ToString.Trim
                        oldtree = do_sql.G_table.Rows(0).Item("tree").ToString.Trim

                        If oldblock > 0 Then ' 在會簽block 內
                            Dim yrangemin As Integer
                            Dim yrangeMax As Integer
                            Dim treemin As Integer
                            Dim treemax As Integer

                            yrangemin = 999
                            yrangeMax = -1
                            treemin = 0
                            treemax = 0

                            sqlstrQT = "select distinct tree as tree from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " and block=" & oldblock
                            If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                                GoTo go_back
                                Exit Sub
                            End If
                            Dim dr As System.Data.DataRow
                            Dim n_table As New System.Data.DataTable
                            Dim NewStepsId As String
                            n_table = do_sql.G_table
                            For Each dr In n_table.Rows

                                sqlstrQT = "select max(isnull(layer,0)) as maxlayer from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " and block=" & oldblock & " and tree=" & Trim(dr("tree").ToString)
                                If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                                    GoTo go_back
                                    Exit Sub
                                End If

                                If CInt(do_sql.G_table.Rows(0).Item("maxlayer").ToString.Trim) < yrangemin Then
                                    yrangemin = CInt(do_sql.G_table.Rows(0).Item("maxlayer").ToString.Trim)
                                    treemin = CInt(Trim(dr("tree").ToString))
                                End If
                                If CInt(do_sql.G_table.Rows(0).Item("maxlayer").ToString.Trim) > yrangeMax Then
                                    yrangeMax = CInt(do_sql.G_table.Rows(0).Item("maxlayer").ToString.Trim)
                                    treemax = CInt(Trim(dr("tree").ToString))
                                End If

                            Next

                            NewStepsId = AddStep(eformid, "NA", oldNextStep, "S", oldX, oldY, Left(group_id, InStr(1, group_id, "_") - 1), Mid(group_id, InStr(1, group_id, "_") + 1), oldblock, CStr(CInt(oldlayer) + 1), oldtree, "1", eformrole)

                            If do_sql.begin_tran(do_sql.G_conn_string) = False Then
                                Exit Sub
                            End If

                            sqlcommand = "update tempflow set nextstep=" & NewStepsId & " where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & stepsid
                            If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                Call do_sql.rollback_tran()
                                GoTo go_back
                                Exit Sub
                            End If

                            If yrangemin = yrangeMax Then 'the same layers for all trees
                                sqlcommand = "update tempflow set y=y+100 where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid <>" & NewStepsId & " and y >= " & oldY & " and tree=" & oldtree & " and block=" & oldblock
                                If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                    Call do_sql.rollback_tran()
                                    GoTo go_back
                                    Exit Sub
                                End If
                                sqlcommand = "update tempflow set y=y+100 where eformid='" & eformid & "' and eformrole=" & eformrole & " and y > " & oldY & " and block <>" & oldblock
                                If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                    Call do_sql.rollback_tran()
                                    GoTo go_back
                                    Exit Sub
                                End If
                            ElseIf oldtree = treemax Then 'selected tree is the max tree
                                sqlcommand = "update tempflow set y=y+100 where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid <>" & NewStepsId & " and y >= " & oldY & " and tree=" & oldtree & " and block=" & oldblock
                                If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                    Call do_sql.rollback_tran()
                                    GoTo go_back
                                    Exit Sub
                                End If
                                sqlcommand = "update tempflow set y=y+100 where eformid='" & eformid & "' and eformrole=" & eformrole & " and y > " & oldY & " and block <>" & oldblock
                                If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                    Call do_sql.rollback_tran()
                                    GoTo go_back
                                    Exit Sub
                                End If
                            Else 'oldx not in the max x 'selected tree is not the max tree
                                sqlcommand = "update tempflow set y=y+100 where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid <>" & NewStepsId & " and y >= " & oldY & " and tree=" & oldtree & " and block=" & oldblock
                                If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                    Call do_sql.rollback_tran()
                                    GoTo go_back
                                    Exit Sub
                                End If
                            End If
                            Call do_sql.commit_tran()
                        Else ' 不在會簽block 內
                            Dim NewStepsId As String
                            NewStepsId = AddStep(eformid, "NA", oldNextStep, "S", oldX, oldY, Left(group_id, InStr(1, group_id, "_") - 1), Mid(group_id, InStr(1, group_id, "_") + 1), oldblock, 0, oldtree, "1", eformrole)
                            If do_sql.begin_tran(do_sql.G_conn_string) = False Then
                                Exit Sub
                            End If
                            sqlcommand = "update tempflow set nextstep=" & NewStepsId & " where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & stepsid
                            If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                Call do_sql.rollback_tran()
                                GoTo go_back
                                Exit Sub
                            End If
                            sqlcommand = "update tempflow set y=y+100 where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid <>" & NewStepsId & " and y >= " & oldY
                            If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                Call do_sql.rollback_tran()
                                GoTo go_back
                                Exit Sub
                            End If
                            Call do_sql.commit_tran()
                        End If
                    Case "PAR" '會簽
                        Dim sqlstrQT As String
                        Dim numblocks As Integer
                        sqlstrQT = "select count(distinct block) as numblocks from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole
                        If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                            GoTo go_back
                            Exit Sub
                        End If

                        numblocks = CInt(do_sql.G_table.Rows(0).Item("numblocks").ToString.Trim)

                        sqlstrQT = "select * from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & stepsid
                        If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                            GoTo go_back
                            Exit Sub
                        End If
                        If do_sql.G_table.Rows.Count = 0 Then
                            GoTo go_back
                            Exit Sub
                        End If
                        Dim oldNextStep As String
                        Dim oldX As String
                        Dim oldY As String
                        Dim oldSteps As String
                        Dim oldblock As String
                        Dim oldlayer As String
                        Dim oldtree As String
                        Dim newblock As String
                        Dim newlayer As String
                        Dim NewStepsId As String
                        Dim newtree As String

                        oldNextStep = do_sql.G_table.Rows(0).Item("nextstep").ToString.Trim
                        oldX = do_sql.G_table.Rows(0).Item("x").ToString.Trim
                        oldY = do_sql.G_table.Rows(0).Item("y").ToString.Trim
                        oldSteps = do_sql.G_table.Rows(0).Item("steps").ToString.Trim
                        oldblock = do_sql.G_table.Rows(0).Item("block").ToString.Trim
                        oldlayer = do_sql.G_table.Rows(0).Item("layer").ToString.Trim
                        oldtree = do_sql.G_table.Rows(0).Item("tree").ToString.Trim

                        If oldblock = 0 Then
                            newblock = CStr(AddBlock(eformid, "1", eformrole))
                        Else
                            newblock = oldblock
                        End If
                        If oldlayer = "0" Then
                            newlayer = "1"
                        Else
                            newlayer = oldlayer
                        End If
                        If numblocks = 1 Then ' 目前沒有會簽存在
                            NewStepsId = AddStep(eformid, oldSteps, oldNextStep, "P", oldX, oldY, Left(group_id, InStr(1, group_id, "_") - 1), Mid(group_id, InStr(1, group_id, "_") + 1), newblock, newlayer, 2, "1", eformrole)
                            If do_sql.begin_tran(do_sql.G_conn_string) = False Then
                                Exit Sub
                            End If
                            sqlcommand = "update tempflow set pors='P',x=" & oldX + 160 & ", block=" & newblock & ", layer=" & newlayer & ", tree=1 where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & stepsid
                            If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                Call do_sql.rollback_tran()
                                GoTo go_back
                                Exit Sub
                            End If
                            sqlcommand = "update tempflow set x=x+80 where eformid='" & eformid & "' and eformrole=" & eformrole & " and steps <>" & oldSteps
                            If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                Call do_sql.rollback_tran()
                                GoTo go_back
                                Exit Sub
                            End If
                            Call do_sql.commit_tran()
                        ElseIf numblocks = 2 And oldblock = "1" Then ' 目前已有一個會簽存在且新增會簽在block 上
                            If CInt(oldlayer) > 1 Then
                                newtree = oldtree
                            ElseIf CInt(oldNextStep) = -1 Then
                                newtree = CStr(AddTree(eformid, oldblock, "1", eformrole))
                            Else
                                sqlstrQT = "select block from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & oldNextStep
                                If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                                    GoTo go_back
                                    Exit Sub
                                End If
                                If do_sql.G_table.Rows.Count = 0 Then
                                    GoTo go_back
                                    Exit Sub
                                End If
                                If do_sql.G_table.Rows(0).Item("block").ToString.Trim = oldblock Then
                                    newtree = oldtree
                                Else
                                    newtree = CStr(CInt(oldtree) + 1)
                                End If
                            End If
                            NewStepsId = AddStep(eformid, oldSteps, oldNextStep, "P", oldX, oldY, Left(group_id, InStr(1, group_id, "_") - 1), Mid(group_id, InStr(1, group_id, "_") + 1), newblock, newlayer, newtree, "1", eformrole)
                            If do_sql.begin_tran(do_sql.G_conn_string) = False Then
                                Exit Sub
                            End If
                            sqlcommand = "update tempflow set pors='P', block=" & newblock & " where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & stepsid
                            If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                Call do_sql.rollback_tran()
                                GoTo go_back
                                Exit Sub
                            End If
                            sqlcommand = "update tempflow set x=x+160 where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid <>" & NewStepsId & " and y=" & oldY & " and x>=" & oldX
                            If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                Call do_sql.rollback_tran()
                                GoTo go_back
                                Exit Sub
                            End If
                            sqlcommand = "update tempflow set x=x+160 where eformid='" & eformid & "' and eformrole=" & eformrole & " and block =" & oldblock & " and y <>" & oldY & " and x >=" & oldX & " and tree <> " & oldtree
                            If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                Call do_sql.rollback_tran()
                                GoTo go_back
                                Exit Sub
                            End If
                            sqlcommand = "update tempflow set x=x+80 where eformid='" & eformid & "' and eformrole=" & eformrole & " and block =" & oldblock & " and y <>" & oldY & " and x >=" & oldX & " and tree = " & oldtree
                            If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                Call do_sql.rollback_tran()
                                GoTo go_back
                                Exit Sub
                            End If
                            sqlcommand = "update tempflow set x=x+80 where eformid='" & eformid & "' and eformrole=" & eformrole & " and block <>" & oldblock
                            If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                Call do_sql.rollback_tran()
                                GoTo go_back
                                Exit Sub
                            End If
                            Call do_sql.commit_tran()
                        Else ' 目前已有多個會簽block存在
                            If CInt(oldblock) = 0 Then '新增會簽在基準線上
                                NewStepsId = AddStep(eformid, oldSteps, oldNextStep, "P", oldX - 80, oldY, Left(group_id, InStr(1, group_id, "_") - 1), Mid(group_id, InStr(1, group_id, "_") + 1), newblock, newlayer, 2, "1", eformrole)
                                sqlcommand = "update tempflow set pors='P',x=" & oldX + 80 & ", block=" & newblock & ", layer=" & newlayer & ",tree=1 where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & stepsid
                                If do_sql.db_exec(sqlcommand, do_sql.G_conn_string) = False Then
                                    GoTo go_back
                                    Exit Sub
                                End If
                            Else
                                Dim currBlocksize As Integer
                                Dim minBlocksize As Integer
                                Dim maxBlocksize As Integer

                                currBlocksize = 0
                                minBlocksize = 9999
                                maxBlocksize = -1

                                sqlstrQT = "select distinct block as blocks from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole
                                If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                                    GoTo go_back
                                    Exit Sub
                                End If

                                Dim dr As System.Data.DataRow
                                Dim n_table As New System.Data.DataTable
                                n_table = do_sql.G_table

                                For Each dr In n_table.Rows
                                    sqlstrQT = "select max(isnull(x,0))-min(isnull(x,0)) as blocksize from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " and block=" & Trim(dr("blocks").ToString)
                                    If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                                        GoTo go_back
                                        Exit Sub
                                    End If

                                    If CInt(do_sql.G_table.Rows(0).Item("blocksize").ToString.Trim) > maxBlocksize Then
                                        maxBlocksize = CInt(do_sql.G_table.Rows(0).Item("blocksize").ToString.Trim)
                                    End If

                                    If CInt(do_sql.G_table.Rows(0).Item("blocksize").ToString.Trim) < minBlocksize Then
                                        minBlocksize = CInt(do_sql.G_table.Rows(0).Item("blocksize").ToString.Trim)
                                    End If
                                    If Trim(dr("blocks").ToString) = oldblock Then
                                        currBlocksize = CInt(do_sql.G_table.Rows(0).Item("blocksize").ToString.Trim)
                                    End If

                                Next

                                'get newtree
                                If CInt(oldlayer) > 1 Then
                                    newtree = oldtree
                                ElseIf oldNextStep = -1 Then
                                    newtree = CStr(AddTree(eformid, oldblock, "1", eformrole))
                                Else

                                    sqlstrQT = "select block from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & oldNextStep
                                    If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                                        GoTo go_back
                                        Exit Sub
                                    End If

                                    If do_sql.G_table.Rows(0).Item("block").ToString.Trim = oldblock Then
                                        newtree = oldtree
                                    Else
                                        newtree = CStr(CInt(oldtree) + 1)
                                    End If
                                End If
                                If do_sql.begin_tran(do_sql.G_conn_string) = False Then
                                    Exit Sub
                                End If
                                If currBlocksize < maxBlocksize Then 'has larger block than this
                                    sqlcommand = "update tempflow set x=x-80 where eformid='" & eformid & "' and eformrole=" & eformrole & " and x <" & oldX & " and block=" & oldblock
                                    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                    NewStepsId = AddStep(eformid, oldSteps, oldNextStep, "P", oldX - 80, oldY, Left(group_id, InStr(1, group_id, "_") - 1), Mid(group_id, InStr(1, group_id, "_") + 1), newblock, newlayer, newtree, "2", eformrole)
                                    sqlcommand = "update tempflow set pors='P',x=" & CStr(CInt(oldX) + 80) & ", block=" & newblock & " where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & stepsid
                                    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                Else
                                    NewStepsId = AddStep(eformid, oldSteps, oldNextStep, "P", oldX, oldY, Left(group_id, InStr(1, group_id, "_") - 1), Mid(group_id, InStr(1, group_id, "_") + 1), newblock, newlayer, newtree, "2", eformrole)
                                    sqlcommand = "update tempflow set pors='P', block=" & newblock & " where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & stepsid
                                    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                    sqlcommand = "update tempflow set x=x+160 where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid <>" & NewStepsId & " and y=" & oldY & " and x>=" & oldX
                                    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                    sqlcommand = "update tempflow set x=x+160 where eformid='" & eformid & "' and eformrole=" & eformrole & " and block =" & oldblock & " and y <>" & oldY & " and x >=" & oldX & " and tree <> " & oldtree
                                    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                    sqlcommand = "update tempflow set x=x+80 where eformid='" & eformid & "' and eformrole=" & eformrole & " and block =" & oldblock & " and y <>" & oldY & " and x >=" & oldX & " and tree = " & oldtree
                                    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                    sqlcommand = "update tempflow set x=x+80 where eformid='" & eformid & "' and eformrole=" & eformrole & " and block <>" & oldblock
                                    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                End If
                                Call do_sql.commit_tran()
                            End If
                        End If
                        'If Err.Number <> 0 Then
                        '    Response.Redirect("errormsg.asp?reason=" & "無法增加流程設定" & "&source=" & Err.Source & "&desc=" & Err.Description)
                        'End If
                    Case "DEL" '刪除此關
                        Dim sqlstrQT As String
                        Dim oldNextStep As String
                        Dim oldSteps As String
                        Dim oldX As String
                        Dim oldY As String
                        Dim oldblock As String
                        Dim oldlayer As String
                        Dim oldtree As String
                        Dim oldpors As String

                        sqlstrQT = "select * from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & stepsid

                        If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                            GoTo go_back
                            Exit Sub
                        End If
                        If do_sql.G_table.Rows.Count = 0 Then
                            GoTo go_back
                            Exit Sub
                        End If

                        oldNextStep = do_sql.G_table.Rows(0).Item("nextstep").ToString.Trim
                        oldSteps = do_sql.G_table.Rows(0).Item("steps").ToString.Trim
                        oldX = do_sql.G_table.Rows(0).Item("x").ToString.Trim
                        oldY = do_sql.G_table.Rows(0).Item("y").ToString.Trim
                        oldblock = do_sql.G_table.Rows(0).Item("block").ToString.Trim
                        oldlayer = do_sql.G_table.Rows(0).Item("layer").ToString.Trim
                        oldtree = do_sql.G_table.Rows(0).Item("tree").ToString.Trim
                        oldpors = do_sql.G_table.Rows(0).Item("pors").ToString.Trim


                        If oldpors = "S" Then '串簽
                            If CInt(oldblock) > 0 Then ' 在會簽block 內
                                Dim yrangemin As Integer
                                Dim yrangeMax As Integer
                                Dim treemin As Integer
                                Dim treemax As Integer

                                yrangemin = 999
                                yrangeMax = -1
                                treemin = 0
                                treemax = 0

                                sqlstrQT = "select distinct tree as tree from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " and block=" & oldblock
                                If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                                    GoTo go_back
                                    Exit Sub
                                End If
                                Dim dr As System.Data.DataRow
                                Dim n_table As New System.Data.DataTable
                                n_table = do_sql.G_table

                                For Each dr In n_table.Rows

                                    sqlstrQT = "select max(isnull(layer,0)) as maxlayer from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " and block=" & oldblock & " and tree=" & Trim(dr("tree").ToString)
                                    If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                                        GoTo go_back
                                        Exit Sub
                                    End If

                                    If CInt(do_sql.G_table.Rows(0).Item("maxlayer").ToString.Trim) < yrangemin Then
                                        yrangemin = CInt(do_sql.G_table.Rows(0).Item("maxlayer").ToString.Trim)
                                        treemin = CInt(Trim(dr("tree").ToString))
                                    End If
                                    If CInt(do_sql.G_table.Rows(0).Item("maxlayer").ToString.Trim) > yrangeMax Then
                                        yrangeMax = CInt(do_sql.G_table.Rows(0).Item("maxlayer").ToString.Trim)
                                        treemax = CInt(Trim(dr("tree").ToString))
                                    End If

                                Next

                                If do_sql.begin_tran(do_sql.G_conn_string) = False Then
                                    Exit Sub
                                End If
                                If DelStep(eformid, stepsid, "2", eformrole) Then
                                    Call do_sql.rollback_tran()
                                    GoTo go_back
                                    Exit Sub
                                End If
                                sqlcommand = "update tempflow set nextstep=" & oldNextStep & " where eformid='" & eformid & "' and eformrole=" & eformrole & " and nextstep =" & stepsid
                                If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                    Call do_sql.rollback_tran()
                                    GoTo go_back
                                    Exit Sub
                                End If
                                sqlcommand = "update tempflow set layer=layer-1 where eformid='" & eformid & "' and eformrole=" & eformrole & " and block =" & oldblock & " and tree=" & oldtree & " and layer >" & oldlayer
                                If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                    Call do_sql.rollback_tran()
                                    GoTo go_back
                                    Exit Sub
                                End If
                                If yrangemin = yrangeMax Then 'the same layers for all trees
                                    sqlcommand = "update tempflow set y=y-100 where eformid='" & eformid & "' and eformrole=" & eformrole & " and y > " & oldY & " and tree=" & oldtree & " and block=" & oldblock
                                    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                ElseIf oldtree = treemax Then 'selected tree is the max tree
                                    sqlcommand = "update tempflow set y=y-100 where eformid='" & eformid & "' and eformrole=" & eformrole & " and y > " & oldY & " and tree=" & oldtree & " and block=" & oldblock
                                    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                    sqlcommand = "update tempflow set y=y-100 where eformid='" & eformid & "' and eformrole=" & eformrole & " and y > " & oldY & " and block <>" & oldblock
                                    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                Else 'oldx not in the max x 'selected tree is not the max tree
                                    sqlcommand = "update tempflow set y=y-100 where eformid='" & eformid & "' and eformrole=" & eformrole & " and y > " & oldY & " and tree=" & oldtree & " and block=" & oldblock
                                    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                End If
                                Call do_sql.commit_tran()
                            Else ' 不在會簽block 內
                                If do_sql.begin_tran(do_sql.G_conn_string) = False Then
                                    Exit Sub
                                End If
                                If DelStep(eformid, stepsid, "2", eformrole) = False Then
                                    Call do_sql.rollback_tran()
                                    GoTo go_back
                                    Exit Sub
                                End If
                                sqlcommand = "update tempflow set nextstep=" & oldNextStep & " where eformid='" & eformid & "' and eformrole=" & eformrole & " and nextstep =" & stepsid
                                If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                    Call do_sql.rollback_tran()
                                    GoTo go_back
                                    Exit Sub
                                End If
                                sqlcommand = "update tempflow set y=y-100 where eformid='" & eformid & "' and eformrole=" & eformrole & " and y >" & oldY
                                If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                    Call do_sql.rollback_tran()
                                    GoTo go_back
                                    Exit Sub
                                End If
                                Call do_sql.commit_tran()
                            End If
                        Else '會簽
                            Dim numblocks As Integer
                            Dim numSteps As Integer

                            sqlstrQT = "select count(distinct block) as numblocks from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole
                            If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                                GoTo go_back
                                Exit Sub
                            End If
                            numblocks = CInt(do_sql.G_table.Rows(0).Item("numblocks").ToString.Trim)

                            If numblocks = 2 Then ' 目前已有一個會簽存在且刪除會簽在block 上                       

                                sqlstrQT = "select count(*) as numsteps from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " and steps=" & oldSteps
                                If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                                    GoTo go_back
                                    Exit Sub
                                End If

                                numSteps = CInt(do_sql.G_table.Rows(0).Item("numsteps").ToString.Trim)
                                If do_sql.begin_tran(do_sql.G_conn_string) = False Then
                                    Exit Sub
                                End If
                                If DelStep(eformid, stepsid, "2", eformrole) = False Then
                                    Call do_sql.rollback_tran()
                                    GoTo go_back
                                    Exit Sub
                                End If
                                If numSteps = 2 Then
                                    sqlcommand = "update tempflow set pors='S' where eformid='" & eformid & "' and eformrole=" & eformrole & " and steps=" & oldSteps
                                    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                End If
                                sqlcommand = "update tempflow set x=x-160 where eformid='" & eformid & "' and eformrole=" & eformrole & " and y=" & oldY & " and x>=" & oldX
                                If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                    Call do_sql.rollback_tran()
                                    GoTo go_back
                                    Exit Sub
                                End If
                                sqlcommand = "update tempflow set x=x-160 where eformid='" & eformid & "' and eformrole=" & eformrole & " and block =" & oldblock & " and y <>" & oldY & " and x >=" & oldX & " and tree <> " & oldtree
                                If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                    Call do_sql.rollback_tran()
                                    GoTo go_back
                                    Exit Sub
                                End If
                                sqlcommand = "update tempflow set x=x-80 where eformid='" & eformid & "' and eformrole=" & eformrole & " and block =" & oldblock & " and y <>" & oldY & " and tree = " & oldtree
                                If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                    Call do_sql.rollback_tran()
                                    GoTo go_back
                                    Exit Sub
                                End If
                                sqlcommand = "update tempflow set x=x-80 where eformid='" & eformid & "' and eformrole=" & eformrole & " and block <>" & oldblock
                                If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                    Call do_sql.rollback_tran()
                                    GoTo go_back
                                    Exit Sub
                                End If


                                sqlstrQT = "select count(*) as numsteps from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " and block=" & oldblock & " and pors='P'"
                                If do_sql.db_sql(sqlstrQT, do_sql.G_Trans) = False Then
                                    Call do_sql.rollback_tran()
                                    GoTo go_back
                                    Exit Sub
                                End If

                                numSteps = CInt(do_sql.G_table.Rows(0).Item("numsteps").ToString.Trim)

                                If numSteps = 0 Then
                                    sqlcommand = "update tempflow set block=0 where eformid='" & eformid & "' and eformrole=" & eformrole & " and block=" & oldblock
                                    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                End If
                                Call do_sql.commit_tran()
                            Else ' 目前已有多個會簽block存在
                                Dim currBlocksize As Integer
                                Dim minBlocksize As Integer
                                Dim maxBlocksize As Integer

                                currBlocksize = 0
                                minBlocksize = 9999
                                maxBlocksize = -1


                                sqlstrQT = "select distinct block as blocks from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole
                                If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                                    Call do_sql.rollback_tran()
                                    GoTo go_back
                                    Exit Sub
                                End If

                                Dim dr As System.Data.DataRow
                                Dim n_table As New System.Data.DataTable
                                n_table = do_sql.G_table

                                For Each dr In n_table.Rows
                                    sqlstrQT = "select max(isnull(x,0))-min(isnull(x,0)) as blocksize from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " and block=" & Trim(dr("blocks").ToString)
                                    If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If

                                    If CInt(do_sql.G_table.Rows(0).Item("blocksize").ToString.Trim) > maxBlocksize Then
                                        maxBlocksize = CInt(do_sql.G_table.Rows(0).Item("blocksize").ToString.Trim)
                                    End If
                                    If CInt(do_sql.G_table.Rows(0).Item("blocksize").ToString.Trim) < minBlocksize Then
                                        minBlocksize = CInt(do_sql.G_table.Rows(0).Item("blocksize").ToString.Trim)
                                    End If
                                    If Trim(dr("blocks").ToString) = oldblock Then
                                        currBlocksize = CInt(do_sql.G_table.Rows(0).Item("blocksize").ToString.Trim)
                                    End If
                                Next

                                If do_sql.begin_tran(do_sql.G_conn_string) = False Then
                                    Exit Sub
                                End If
                                If currBlocksize < maxBlocksize Then 'has larger block than this
                                    sqlcommand = "update tempflow set x=x+80 where eformid='" & eformid & "' and eformrole=" & eformrole & " and x <" & oldX & " and block=" & oldblock
                                    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                    If DelStep(eformid, stepsid, "2", eformrole) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                Else
                                    If DelStep(eformid, stepsid, "2", eformrole) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If

                                    sqlcommand = "update tempflow set x=x-160 where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid <>" & stepsid & " and y=" & oldY & " and x>=" & oldX
                                    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                    sqlcommand = "update tempflow set x=x-160 where eformid='" & eformid & "' and eformrole=" & eformrole & " and block =" & oldblock & " and y <>" & oldY & " and x >=" & oldX & " and tree <> " & oldtree
                                    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                    sqlcommand = "update tempflow set x=x-80 where eformid='" & eformid & "' and eformrole=" & eformrole & " and block =" & oldblock & " and y <>" & oldY & " and x >=" & oldX & " and tree = " & oldtree
                                    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                        Call do_sql.rollback_tran()
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                    sqlcommand = "update tempflow set x=x-80 where eformid='" & eformid & "' and eformrole=" & eformrole & " and block <>" & oldblock
                                    If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                                        Call do_sql.rollback_tran()
                                        'Response.Redirect("errormsg.asp?reason=" & "無法刪除流程設定" & "&source=" & sqlcommand & "&desc=" & do_sql.G_errmsg)
                                        GoTo go_back
                                        Exit Sub
                                    End If
                                End If
                                Call do_sql.commit_tran()
                            End If

                        End If
                        'If Err.Number <> 0 Then
                        '    Response.Redirect("errormsg.asp?reason=" & "無法刪除流程設定" & "&source=" & Err.Source & "&desc=" & Err.Description)
                        'End If
                    Case "finish" '完成
                        If do_sql.begin_tran(do_sql.G_conn_string) = False Then
                            Exit Sub
                        End If

                        sqlcommand = "delete from flow where eformid='" & eformid & "' and eformrole=" & eformrole
                        If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                            Call do_sql.rollback_tran()
                            GoTo go_back
                            Exit Sub
                        End If
                        sqlcommand = "insert into flow select * from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole
                        If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                            Call do_sql.rollback_tran()
                            GoTo go_back
                            Exit Sub
                        End If
                        'sqlcommand = "delete from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole
                        'If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                        '    Call do_sql.rollback_tran()
                        '    GoTo go_back
                        '    Exit Sub
                        'End If
                        '--------
                        sqlcommand = "delete from cc where eformid='" & eformid & "' and eformrole=" & eformrole
                        If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                            Call do_sql.rollback_tran()
                            GoTo go_back
                            Exit Sub
                        End If
                        sqlcommand = "insert into cc select * from tempcc where eformid='" & eformid & "' and eformrole=" & eformrole
                        If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                            Call do_sql.rollback_tran()
                            GoTo go_back
                            Exit Sub
                        End If
                        'sqlcommand = "delete from tempcc where eformid='" & eformid & "' and eformrole=" & eformrole
                        'If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                        '    Call do_sql.rollback_tran()
                        '    GoTo go_back
                        '    Exit Sub
                        'End If
                        '-----------
                        sqlcommand = "delete from eforms where eformid='" & eformid & "'"
                        If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                            Call do_sql.rollback_tran()
                            GoTo go_back
                            Exit Sub
                        End If
                        sqlcommand = "insert into eforms select * from tempeforms where eformid='" & eformid & "'"
                        If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                            Call do_sql.rollback_tran()
                            GoTo go_back
                            Exit Sub
                        End If
                        'sqlcommand = "delete from tempeforms where eformid='" & eformid & "'"
                        'If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                        '    Call do_sql.rollback_tran()
                        '    GoTo go_back
                        '    Exit Sub
                        'End If
                        Call do_sql.commit_tran()
                        'Response.Write("<script language='javascript'>window.parent.location='MOA00001.aspx?" + "</" + "Script>")
                        redirect_flag = True '導向MOA00001.aspx
                        Exit Sub
                End Select
                'objUniversal_r = Nothing
                'objUniversal_w = Nothing

go_back:
                Response.Redirect("MOA00005.aspx?eformid=" & eformid & "&eformrole= " & eformrole & "&frm_chinese_name=" & frm_chinese_name & "&organization_id=" & organization_id & "&scrollLeftPixel=" & scrollLeftPixel & "&scrollTopPixel=" & scrollTopPixel & "&save=TRUE&err_msg=" & do_sql.G_errmsg)



            End If

        Catch ex As Exception

        End Try
    End Sub
    'conn_p "1"為connection 字串 否則為tranction
    Function DelStep(ByVal eformid As String, ByVal stepsid As String, ByVal conn_p As String, ByVal eformrole As String) As Boolean
        DelStep = False
        Dim sqlcommand As String
        sqlcommand = "delete from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & stepsid
        If conn_p = "1" Then
            If do_sql.db_exec(sqlcommand, do_sql.G_conn_string) = False Then
                Exit Function
            End If
        Else
            If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                Exit Function
            End If
        End If
        sqlcommand = "delete from tempcc where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & stepsid
        If conn_p = "1" Then
            If do_sql.db_exec(sqlcommand, do_sql.G_conn_string) = False Then
                Exit Function
            End If
        Else
            If do_sql.db_exec(sqlcommand, do_sql.G_Trans) = False Then
                Exit Function
            End If
        End If
        DelStep = True
    End Function
    Function AddBlock(ByVal eformid As String, ByVal conn_p As String, ByVal eformrole As String) As Integer
        AddBlock = -1
        Dim sqlstrQT As String
        sqlstrQT = "select max(isnull(block,0)) as block from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole
        If conn_p = "1" Then
            If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                Exit Function
            End If
        Else
            If do_sql.db_sql(sqlstrQT, do_sql.G_Trans) = False Then
                Exit Function
            End If
        End If
        AddBlock = CInt(do_sql.G_table.Rows(0).Item("block").ToString.Trim) + 1
    End Function
    Function AddTree(ByVal eformid As String, ByVal block As String, ByVal conn_p As String, ByVal eformrole As String) As Integer
        Dim sqlstrQT As String
        sqlstrQT = "select max(isnull(tree,0)) as tree from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " and block=" & block
        If conn_p = "1" Then
            If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                Exit Function
            End If
        Else
            If do_sql.db_sql(sqlstrQT, do_sql.G_Trans) = False Then
                Exit Function
            End If
        End If
        AddTree = CInt(do_sql.G_table.Rows(0).Item("tree").ToString.Trim) + 1
    End Function

    '以下-----from inc_AddStep.asp
    Function AddStep(ByVal eformid As String, ByVal steps As String, ByVal nextstep As String, ByVal pors As String, ByVal x As String, ByVal y As String, ByVal group_id As String, ByVal group_type As String, ByVal block As String, ByVal layer As String, ByVal tree As String, ByVal conn_p As String, ByVal eformrole As String) As String
        AddStep = ""
        Dim sSqlcommand, major_step
        Dim stepsid As Integer
        stepsid = CInt(getIntID())
        sSqlcommand = "insert into tempflow(eformid,eformrole, stepsid, steps, nextstep, major_step, pors, x, y, group_id, group_type, block,layer,tree,allhandle,canjump,canback,canadd,candist,candistcc,canaddatt,candelatt,canconti,bypass,opinonly,overduenotice) values ('" & _
           eformid & "'," & eformrole & "," & stepsid & ","
        If steps = "NA" Then
            sSqlcommand = sSqlcommand & stepsid.ToString
            major_step = "1"
        Else
            sSqlcommand = sSqlcommand & steps
            major_step = "0"
        End If
        sSqlcommand = sSqlcommand & "," & nextstep & ",'" & major_step & "','" & pors & "'," & x & "," & y & ",'" & group_id & "','" & group_type & "'," & block & "," & layer & "," & tree & ",'2','0','1','0','0','0','1','1','0','0','0','1')"
        If conn_p = "1" Then
            If do_sql.db_exec(sSqlcommand, do_sql.G_conn_string) = False Then
                Exit Function
            End If
        Else
            If do_sql.db_exec(sSqlcommand, do_sql.G_Trans) = False Then
                Exit Function
            End If
        End If
        AddStep = stepsid.ToString

    End Function
    Function getIntID()
        Dim sID, iMi, iSe
        Dim iHr As Integer

        Randomize()
        sID = ""
        sID = sID & CStr(Int(3 * Rnd))
        iHr = CInt(Now.ToString("HH"))
        If iHr > 9 Then
            sID = sID & CStr((CInt(left(CStr(iHr), 1)) + CInt(right(CStr(iHr), 1))) Mod 10)
        Else
            sID = sID & CStr(iHr)
        End If
        iMi = CInt(Now.ToString("mm"))
        If iMi > 9 Then
            sID = sID & CStr((CInt(left(CStr(iMi), 1)) + CInt(right(CStr(iMi), 1))) Mod 10)
        Else
            sID = sID & CStr(iMi)
        End If
        iSe = CInt(Now.ToString("ss"))
        If iSe > 9 Then
            sID = sID & CStr((CInt(left(CStr(iSe), 1)) + CInt(right(CStr(iSe), 1))) Mod 10)
        Else
            sID = sID & CStr(iSe)
        End If
        sID = sID & CStr(Int(10 * Rnd))
        getIntID = sID
    End Function
End Class
