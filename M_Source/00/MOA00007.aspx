<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00007.aspx.vb" Inherits="Source_00_MOA00007" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
<link rel="stylesheet" type="text/css" href="../../Styles.css" />
    <title>傳送流程</title>
    <script type="text/JavaScript" language="JavaScript">
        function switchdisplay(sid, type) {
            var obj = document.getElementById(sid);
            obj.style.display = type ? '' : 'none';
        }
</script>
</head>
<body>
    
    <span style="top:3000px;left:2000;position:absolute;">&nbsp;</span>
<div id="showbody" style="position:absolute; top:0 ;left:200;display:none; ">

<%
    Dim sNodeTitle As String = ""
    Dim sqlstrQT As String

    If bFlowDesignMode Or bDesignColumnMode Then
        sFlowTable = "tempflow"
        sCCTable = "tempcc"
        sEformTable = "tempeforms"
    Else
        sFlowTable = "flow"
        sCCTable = "cc"
        sEformTable = "eforms"
    End If
    
    Response.Write("<span id=flowbody style='position:absolute;top:0;left:0;'>")
    NodesEndPosX = 0 ' X position for Nodes 結束 -used for 基準線
    'sSqlcommand="select * from " & sFlowTable & " where eformid='" & eformid & "' order by y desc, x asc"
    ' 畫圖時由下而上畫起
    'set rsNodes=objRouting.ExecSQLReturnRS(sSqlcommand)


    sqlstrQT = "select * from " & sFlowTable & " where eformid='" & eformid & "'"
    sqlstrQT += " and eformrole=" & eformrole
    sqlstrQT += " order by y desc, x asc"
    If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
        Exit Sub
    End If

    Dim dr As System.Data.DataRow
    Dim dr2 As System.Data.DataRow
    Dim n_table As New System.Data.DataTable
    Dim n_table2 As New System.Data.DataTable
    n_table = do_sql.G_table
    
    Dim stepi As Integer = 0

    For Each dr In n_table.Rows
        Select Case Trim(dr("steps").ToString)
            Case "-1" '結束
                Response.Write("<span STYLE='padding:0;position:absolute;width=80;height:25px;top:" & CStr(CInt(Trim(dr("y").ToString)) - 12) & "px;left:" & CStr(CInt(Trim(dr("x").ToString))) & "px;text-align:center;'>")
                Response.Write("   <img ID='I" & Trim(dr("stepsid").ToString) & "' src='../../image/ok.gif' border=0 WIDTH=40 HEIGHT=30 title=結束")
                If bFlowDesignMode Then
                    Response.Write("  style='CURSOR: hand' onmousedown='showPopup_1()'")
                End If
                Response.Write("><br><nobr>")
                'If Request.Cookies("global")("language") = "EN" Or Request.Cookies("global")("language") = "JA" Then
                '    Response.Write("Finish")
                'Else
                Response.Write("結束")
                'End If
                Response.Write("</nobr>")
                Response.Write("</span>")
                Response.Write("<img STYLE='position:absolute;top:" & CStr(CInt(Trim(dr("y").ToString)) + 44) & "px;left:" & CStr(CInt(Trim(dr("x").ToString)) - 20) & "px;width:80px;z-index:1' src='../../image/tryA.gif' WIDTH=81 HEIGHT=4>")
                Response.Write("<img STYLE='position:absolute;top:" & CStr(CInt(Trim(dr("y").ToString)) + 45) & "px;left:" & CStr(CInt(Trim(dr("x").ToString)) - 20) & "px;width:81px;height:18px;' src='../../image/L1.gif' WIDTH=81 HEIGHT=38>")
                NodesEndPosX = Trim(dr("x").ToString) '以結束關當作X軸基準線
            Case "0" '申請人
                Response.Write("<img STYLE='position:absolute;top:" & CStr(CInt(Trim(dr("y").ToString)) - 37) & "px;left:" & CStr(CInt(Trim(dr("x").ToString)) - 20) & "px;' src='../../image/L1.gif' WIDTH=81 HEIGHT=38>")
                Response.Write("<span STYLE='padding:0;position:absolute;width=80;height:25px;top:" & CStr(CInt(Trim(dr("y").ToString)) - 12) & "px;left:" & CStr(CInt(Trim(dr("x").ToString))) & "px;text-align:center;'>")
                Response.Write("   <img ID='I" & Trim(dr("stepsid").ToString) & "' src='../../image/apy.gif' border=0 title=申請人  style='CURSOR: hand'")
                If bFlowDesignMode Then
                    Response.Write(" onmousedown='showPopup0()'")
                ElseIf bDesignColumnMode Then
                    Response.Write(" onmousedown= 'showFlowFields()'")
                End If
                Response.Write(" WIDTH=40 HEIGHT=30><br><nobr>")
                'If Request.Cookies("global")("language") = "EN" Or Request.Cookies("global")("language") = "JA" Then
                '    Response.Write("Applyer")
                'Else
                Response.Write("申請人")
                'End If
                Response.Write("</nobr>")
                Response.Write("</span>")
            Case Else
                'set obj=server.CreateObject("Routing.flow_r")
                'sNodeDisplayName=obj.GetFlowGroupTitle(g_SYSMIS_db,rsNodes("group_type"),rsNodes("group_id"),rsNodes("alias_group_name"))
                'sNodeOriginalName=obj.GetFlowGroupTitle(g_SYSMIS_db,rsNodes("group_type"),rsNodes("group_id"),"")
                Dim sNodeDisplayName As String = ""

                sqlstrQT = "select * from SYSTEMOBJ where object_uid='" & Trim(dr("group_id").ToString) & "'"
                If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                    Exit Sub
                End If
                If do_sql.G_table.Rows.Count > 0 Then
                    sNodeDisplayName = do_sql.G_table.Rows(0).Item("object_name").ToString.Trim
                End If

                If Trim(dr("major_step").ToString) & "" = "1" Then
                    sNodeDisplayName = sNodeDisplayName & "*"
                End If

                Response.Write("<img STYLE='position:absolute;top:" & CStr(CInt(Trim(dr("y").ToString)) - 37) & "px;left:" & CStr(CInt(Trim(dr("x").ToString)) - 20) & "px;' src='../../image/L1.gif' WIDTH=81 HEIGHT=38>")
                
                '各個圖檔DIV的ID
                Dim Perid As String = dr("group_id") & stepi
                
                If bFlowDesignMode = False Then
                    Response.Write("<span onMouseOver=""javascript:switchdisplay('" & Perid & "',true)"" onMouseOut=""javascript:switchdisplay('" & Perid & "',false)"" STYLE='padding:0;position:absolute;width=80;height:25px;top:" & CStr(CInt(Trim(dr("y").ToString)) - 12) & "px;left:" & CStr(CInt(Trim(dr("x").ToString)) - 30) & "px;text-align:center;'>")
                Else
                    Response.Write("<span STYLE='padding:0;position:absolute;width=80;height:25px;top:" & CStr(CInt(Trim(dr("y").ToString)) - 12) & "px;left:" & CStr(CInt(Trim(dr("x").ToString)) - 30) & "px;text-align:center;'>")
                End If
                
                Response.Write("   <img ID='I" & Trim(dr("stepsid").ToString) & "' style='CURSOR: hand'")
                
                If Trim(dr("group_type").ToString) = "0" Or Trim(dr("group_type").ToString) = "3" Or Trim(dr("group_type").ToString) = "6" Or Trim(dr("group_type").ToString) = "a" Then
                    Response.Write("src='../../image/head.gif'")
                ElseIf Trim(dr("group_type").ToString) = "1" Then '個人
                    Response.Write("src='../../image/one.gif'")
                Else
                    Response.Write("src='../../image/group.gif'")
                End If
                
                Response.Write(" border=0 ")
                
                If bFlowDesignMode Then
                    Response.Write("title=""" & sNodeTitle & """ onmousedown='show_div()'")
                ElseIf bDesignColumnMode Then
                    Response.Write("title=""" & sNodeTitle & """ onmousedown= 'showFlowFields()'")
                Else
                    Response.Write("title=""" & sNodeDisplayName & """")
                End If
                
                Response.Write(" WIDTH=40 HEIGHT=30><br><nobr>")
                
                '顯示出關卡人員名稱
                If bFlowDesignMode = False Then
                    Response.Write(sNodeDisplayName)
                    
                    '判斷顯示人員高度與寬度     
                    Dim RelTop As Integer = 0
                    Dim RelLeft As Integer = 0
                    
                    '寬度
                    If CInt(Trim(dr("x").ToString)) > 300 Then
                        RelLeft = 100
                    Else
                        RelLeft = 200
                    End If
                    
                    
                    '顯示上一級主管
                    If dr("group_id") = "Z860" Then
                        
                        
                        Dim UpStepID As Integer = Trim(dr("stepsid").ToString)
                        
                        '上一級是否為申請者
                        Dim sqlAppUser As String = "select group_id from " & sFlowTable & " where eformid='" & eformid & "' and nextstep =" & UpStepID
                        If do_sql.db_sql(sqlAppUser, do_sql.G_conn_string) = False Then
                            Exit Sub
                        End If
                        
                        If do_sql.G_table.Rows(0).Item("group_id").ToString() = "0" Then
                            
                            '顯示人員
                            Dim sqlstrUpStep As String = "select emp_chinese_name from EMPLOYEE where ORG_UID IN (select PARENT_ORG_UID from ADMINGROUP where ORG_UID IN (select ORG_UID from EMPLOYEE e where e.employee_id ='" & Session("user_id") & "'))"
                            If do_sql.db_sql(sqlstrUpStep, do_sql.G_conn_string) = False Then
                                Exit Sub
                            End If
                            
                        Else
                            
                            
                            '顯示人員
                            Dim sqlstrUpStep As String = "select emp_chinese_name from EMPLOYEE where ORG_UID IN (select PARENT_ORG_UID from ADMINGROUP where ORG_UID IN (select ORG_UID from EMPLOYEE e where e.employee_id IN (select employee_id from SYSTEMOBJUSE where object_uid IN (select group_id from " & sFlowTable & " where eformid='" & eformid & "' and nextstep =" & UpStepID & "))))"
                            If do_sql.db_sql(sqlstrUpStep, do_sql.G_conn_string) = False Then
                                Exit Sub
                            End If
                            
                        End If
                        
                    Else
                        
                        '顯示人員
                        Dim sqlstrQTSys As String = "select emp_chinese_name from SYSTEMOBJUSE s,EMPLOYEE e where e.employee_id = s.employee_id and s.object_uid='" & dr("group_id") & "' and leave='Y'"
                        If do_sql.db_sql(sqlstrQTSys, do_sql.G_conn_string) = False Then
                            Exit Sub
                        End If
                        
                    End If
                    
                    Dim drSys As System.Data.DataRow
                    Dim nSys_table As New System.Data.DataTable
                    nSys_table = do_sql.G_table
                                        
                    Response.Write("<div id='" & Perid & "' style='position:absolute; z-index:3; border:1 solid lightslategray; background-color:white; width:150pt; height:60pt; left:" & RelLeft & "px; top:" & RelTop & "px;'>")
                    
                    Response.Write("<table>")
                    Response.Write("<tr>")
                    
                    If nSys_table.Rows.Count >= 10 Then
                        Dim RDI As Integer = 0
                        For Each drSys In nSys_table.Rows
                            Response.Write("<td class='row_1'>")
                            Response.Write(drSys("emp_chinese_name") & "<br>")
                            Response.Write("</td>")
                            
                            RDI = RDI + 1
                            If RDI Mod 3 = 0 Then
                                Response.Write("</tr>")
                                Response.Write("<tr>")
                            End If
                        Next
                        
                    Else
                        
                        Response.Write("<td class='row_1'>")
                        For Each drSys In nSys_table.Rows
                            Response.Write(drSys("emp_chinese_name") & "<br>")
                        Next
                        Response.Write("</td>")
                        
                    End If
                    
                    Response.Write("</tr>")
                    Response.Write("</table>")
                    
                    Response.Write("</div>")
                    
                    '隱藏
                    Response.Write("<script language='javascript'>")
                    Response.Write("switchdisplay('" & Perid & "',false);")
                    Response.Write("</script>")
                    
                Else
                    Response.Write(sNodeDisplayName)
                End If
                
                Response.Write("</nobr>")
                'display cc nodes

                sqlstrQT = "select * from " & sCCTable & " where eformid='" & eformid & "' and eformrole=" & eformrole & " and stepsid=" & Trim(dr("stepsid").ToString)
                If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                    Exit Sub
                End If


                If do_sql.G_table.Rows.Count > 0 Then
                    Response.Write("<span ID=CC" & Trim(dr("stepsid").ToString) & " STYLE='padding:0;font-size:9pt;position:absolute;top:-25;left:-90;Z-INDEX:103;display:none' onMouseOver=""this.style.display=''""' onMouseOut=""this.style.display='none'"">" & Chr(13) & Chr(10))
                    Response.Write("   <table class='setbox' width='100%' cellpadding='0' cellspacing='1' >" & Chr(13) & Chr(10))
                    Response.Write("		<tr><td>" & Chr(13) & Chr(10))
                    Response.Write("			<table width='100%' border='0' bgcolor='cccccc' cellpadding='2' cellspacing='0'>" & Chr(13) & Chr(10))

                    Response.Write("         <tr><td class='c2' align='center' colspan='2'><nobr>")
                    'If Request.Cookies("global")("language") = "EN" Or Request.Cookies("global")("language") = "JA" Then
                    ' Response.Write("Notice")
                    'Else
                    Response.Write("副本通知")
                    'End If
                    Response.Write("</nobr></td></tr>" & Chr(13) & Chr(10))
                    Dim cclist As String
                    
                    cclist = ""
                    Dim groupname As String = ""
                    n_table2 = do_sql.G_table
                    For Each dr2 In n_table2.Rows
                        '未做
                        'groupname = obj.GetFlowGroupTitle(g_SYSMIS_db, Trim(dr2("group_type").ToString),  Trim(dr2("group_id").ToString), null)

                        Response.Write("         <tr align='center' class='nr' onmouseover=""this.className='cl'"" onmouseout=""this.className='nr'"" onmousedown=""delCC('" & Trim(dr("stepsid").ToString) & "','" & Trim(dr2("group_id").ToString) & "')"">" & Chr(13) & Chr(10))
                        Response.Write("            <td><img src='../../image/x.gif' border=0 alt=刪除該通知></td><td><nobr>" & groupname & "</nobr></td>" & Chr(13) & Chr(10))
                        cclist = cclist & groupname & ","

                    Next
                    cclist = Left(cclist, Len(cclist) - 1)
                    Response.Write("         </tr>")
                    Response.Write("			</table>")
                    Response.Write("		</td></tr>")
                    Response.Write("   </table>")
                    Response.Write("</span>")
                    Response.Write("<img id=img" & Trim(dr("stepsid").ToString) & " style='position:absolute;left:-62;top:-20;' src='../../image/CC.gif' title='通知' ")
                    If bFlowDesignMode Then
                        Response.Write("  style='CURSOR: hand'  onmousedown = 'CC" & Trim(dr("stepsid").ToString) & ".style.display ="""" ;'")
                    End If
                    Response.Write(" >")
                    ' extra code below for displaying cclist
                    Response.Write("<span align=center style='position:absolute;left:-62;top:2;width:72px;font-size:7pt;line-height:9pt;color:blue;'>" & cclist & "</span>")
                    ' ----
                End If
                
                ' end display cc node	
                Response.Write("</span>")

                ' display arrow and vertical line	
                Response.Write("<img STYLE='position:absolute;top:" & CStr(CInt(Trim(dr("y").ToString)) + 44) & "px;left:" & CStr(CInt(Trim(dr("x").ToString)) - 20) & "px;width:81px;height:4px;z-index:1' src='../../image/tryA.gif' >")

        End Select
        stepi = stepi + 1
    Next
    
    Dim nodecount As String

    'draw horizontal line
    For Each dr In n_table.Rows
        If Trim(dr("steps").ToString) = -1 Then ' 結束
            Exit For
        End If
        NextStep = Trim(dr("NextStep").ToString)

        sqlstrQT = "select count(*) as nodecount from " & sFlowTable & " where eformid='" & eformid & "' and eformrole=" & eformrole & " and steps=" & NextStep
        If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
            Exit Sub
        End If


        nodecount = do_sql.G_table.Rows(0).Item("nodecount").ToString.Trim

        If CInt(nodecount) > 1 Then ' has many nextstep : one to many
            'sSqlcommand="select count(*) as nodecount, min(x) as minx, max(x) as maxx from " & sFlowTable & " where eformid='" & eformid & "' and y=" &  rsNodes("y") 
            'set rsStep=objRouting.ExecSQLReturnRS(sSqlcommand)

            sqlstrQT = "select count(*) as nodecount, min(isnull(x,0)) as minx, max(isnull(x,0)) as maxx from " & sFlowTable & " where eformid='" & eformid & "' and eformrole=" & eformrole & " and y=" & Trim(dr("y").ToString)
            If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                Exit Sub
            End If

            nodecount = do_sql.G_table.Rows(0).Item("nodecount").ToString.Trim
            minx = do_sql.G_table.Rows(0).Item("minx").ToString.Trim
            maxx = do_sql.G_table.Rows(0).Item("maxx").ToString.Trim

            If CInt(nodecount) > 1 Then 'many to many
                startX = 0
                endX = 0
                'sSqlcommand="select x,y from " & sFlowTable & " where eformid='" & eformid & "' and steps=" &  NextStep & " order by x asc"
                'set rsStep=objRouting.ExecSQLReturnRS(sSqlcommand)



                sqlstrQT = "select x,y from " & sFlowTable & " where eformid='" & eformid & "' and eformrole=" & eformrole & " and steps=" & NextStep & " order by x asc"
                If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                    Exit Sub
                End If
                If do_sql.G_table.Rows.Count > 0 Then
                    startX = do_sql.G_table.Rows(0).Item("x").ToString.Trim
                End If

                n_table2 = do_sql.G_table
                For Each dr2 In n_table2.Rows
                    Response.Write("<img STYLE='position:absolute;top:" & CStr(CInt(Trim(dr2("y").ToString)) + 45) & "px;left:" & CStr(CInt(Trim(dr2("x").ToString)) - 20) & "px;width:81px;height:" & CStr(CInt(Trim(dr2("y").ToString)) - CInt(Trim(dr2("y").ToString)) - 93) & "px;' src='../../image/L1.gif'>")
                    endX = Trim(dr2("x").ToString)
                Next

                Response.Write("<img STYLE='position:absolute;top:" & CStr(CInt(Trim(dr("y").ToString)) - 48) & "px;left:" & NodesEndPosX & "px;width:81px;height:10px;' src='../../image/L1.gif'>")
                Response.Write("<img STYLE= 'position:absolute;top:" & CStr(CInt(Trim(dr("y").ToString)) - 38) & "px;left:" & CStr(CInt(minx) + 40) & "px;width:" & CStr(CInt(maxx) - CInt(minx)) & "px;height:1px;' src='../../image/H.gif'>")
                Response.Write("<img STYLE= 'position:absolute;top:" & CStr(CInt(Trim(dr("y").ToString)) - 48) & "px;left:" & CStr(CInt(startX) + 40) & "px;width:" & CStr(CInt(endX) - CInt(startX)) & "px;height:1px;' src='../../image/H.gif'>")
            Else
                startX = "0"
                endX = "0"
                'sSqlcommand="select x,y from " & sFlowTable & " where eformid='" & eformid & "' and steps=" &  NextStep & " order by x asc"
                'set rsStep=objRouting.ExecSQLReturnRS(sSqlcommand)

                sqlstrQT = "select x,y from " & sFlowTable & " where eformid='" & eformid & "' and eformrole=" & eformrole & " and steps=" & NextStep & " order by x asc"
                If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                    Exit Sub
                End If

                If do_sql.G_table.Rows.Count > 0 Then
                    startX = do_sql.G_table.Rows(0).Item("x").ToString.Trim
                End If
                n_table2 = do_sql.G_table
                For Each dr2 In n_table2.Rows
                    Response.Write("<img STYLE='position:absolute;top:" & CStr(CInt(Trim(dr2("y").ToString)) + 45) & "px;left:" & CStr(CInt(Trim(dr2("x").ToString)) - 20) & "px;width:81px;height:" & CStr(CInt(Trim(dr("y").ToString)) - CInt(Trim(dr2("y").ToString)) - 83) & "px;' src='../../image/L1.gif'>")
                    endX = Trim(dr2("x").ToString)
                Next

                Response.Write("<img STYLE= 'position:absolute;top:" & CStr(CInt(Trim(dr("y").ToString)) - 38) & "px;left:" & CStr(CInt(startX) + 40 - 20) & "px;width:" & CStr(CInt(endX) - CInt(startX)) & "px;height:1px;' src='../../image/H.gif'>")
            End If
            ' many to many   
        Else ' has one nextstep
            'sSqlcommand="select count(*) as nodecount,min(x) as minx, max(x) as maxx from " & sFlowTable & " where eformid='" & eformid & "' and y=" &  rsNodes("y") 
            'set rsStep=objRouting.ExecSQLReturnRS(sSqlcommand)      


            sqlstrQT = "select count(*) as nodecount,min(isnull(x,0)) as minx, max(isnull(x,0)) as maxx from " & sFlowTable & " where eformid='" & eformid & "' and eformrole=" & eformrole & " and y=" & Trim(dr("y").ToString)
            If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                Exit Sub
            End If

            nodecount = do_sql.G_table.Rows(0).Item("nodecount").ToString.Trim
            startX = do_sql.G_table.Rows(0).Item("minx").ToString.Trim
            endX = do_sql.G_table.Rows(0).Item("maxx").ToString.Trim

            sqlstrQT = "select min(isnull(nextstep,0)) as minnextstep, max(isnull(nextstep,0)) as maxnextstep from " & sFlowTable & " where eformid='" & eformid & "' and eformrole=" & eformrole & " and y=" & Trim(dr("y").ToString)
            If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                Exit Sub
            End If

            minnextstep = do_sql.G_table.Rows(0).Item("minnextstep").ToString.Trim
            maxnextstep = do_sql.G_table.Rows(0).Item("maxnextstep").ToString.Trim

            If CInt(nodecount) > 1 And minnextstep = maxnextstep Then ' many to 1;if minnextstep=maxnextstep 表示有相同的nextstep

                sqlstrQT = "select x,y from " & sFlowTable & " where eformid='" & eformid & "' and eformrole=" & eformrole & " and steps=" & NextStep 'not stepsid becase this step may be delete when processing
                If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = False Then
                    Exit Sub
                End If

                If do_sql.G_table.Rows.Count > 0 Then
                    Response.Write("<img STYLE='position:absolute;top:" & CStr(CInt(do_sql.G_table.Rows(0).Item("y").ToString) + 45) & "px;left:" & CStr(CInt(do_sql.G_table.Rows(0).Item("x").ToString) - 20) & "px;width:81px;height:18px;' src='../../image/L1.gif'>")
                End If
                'rsStep = Nothing
                'draw horizontal line
                Response.Write("<img STYLE= 'position:absolute;top:" & CStr(CInt(Trim(dr("y").ToString)) - 38) & "px;left:" & CStr(CInt(startX) + 40 - 20) & "px;width:" & CStr(CInt(endX) - CInt(startX)) & "px;height:1px;' src='../../image/H.gif'>")
            Else ' 1 to 1
                Response.Write("<img STYLE='position:absolute;top:" & CStr(CInt(Trim(dr("y").ToString)) - 55) & "px;left:" & CStr(CInt(Trim(dr("x").ToString)) - 20) & "px;width:81px;height:18px;' src='../../image/L1.gif'>")
            End If
        End If

    Next

    Response.Write("</span>")
    %>


<!-- inc_FlowChart -->
<%  

    'Dim sqlstrQT As String
    call first_step()

%>
</div>
 <script language="javascript">
showbody.style.pixelLeft =  -399 + document.body.offsetWidth / 2 - 40; 
showbody.style.display = '';
flowbody.style.display = '';
<% 
     call second_step()
%>
</script>
<!--如影隨形-->

<div id="KBStatic" style="position:absolute;display:none">
<!--
<dd>
<a HREF="<%=backto%>">
<img SRC="images/addBtnOk_up.gif" ALT="<%=backtomsg%>" NOSAVE BORDER="0" WIDTH="21" HEIGHT="21">
</a>
</dd>
-->
</div>
<script language="JavaScript">
function KB_keepIt(theName,theWantTop,theWantLeft) {
	theRealTop=parseInt(document.body.scrollTop)
	theTrueTop=theWantTop+theRealTop
	document.all[theName].style.top=theTrueTop
	theRealLeft=parseInt(document.body.scrollLeft)
	theTrueLeft=theWantLeft+theRealLeft
	document.all[theName].style.left=theTrueLeft
}
setInterval('KB_keepIt("KBStatic",10,10)',1)
</script>

<!--如影隨形-->   
</body>
</html>
