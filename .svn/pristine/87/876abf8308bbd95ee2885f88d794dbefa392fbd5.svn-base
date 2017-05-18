<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00005.aspx.vb" Inherits="Source_00_MOA00005"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>表單流程拖曳頁面</title>
    <link rel="stylesheet" type="text/css" href="../../Styles.css" />
    <style>
.cl  {font:9pt; color:#FFFFFF; background:#000099; cursor:hand}
.nr  {font:9pt; color:#000000}
.c2  {font:9pt; color:#FFFFFF; letter-spacing:0.1em; background-color:#707070}
span {color:#000000; cursor:default; text-decoration:none;}
.grouptxt
{
    cursor:hand;
}
.setbox
{
    BACKGROUND-COLOR: #cccccc;
    BORDER-BOTTOM: dimgray thin solid;
    BORDER-LEFT: silver thin solid;
    BORDER-RIGHT: dimgray thin solid;
    BORDER-TOP: silver thin solid
}
</style>
<script type="text/JavaScript" language="JavaScript">
    function switchdisplay(sid, type) {
        var obj = document.getElementById(sid);
        obj.style.display = type ? '' : 'none';
    }
</script>
</head>
<body onload="ini()" style="margin-top:0px;margin-left:2px; background-position: left top;background-color:#ffffff;BORDER-left: #5A6B8C 1px solid; " ondragenter="fOnDragOver()" ondragover="fOnDragOver()" ondrop="fOnDrop()">
    <form id="form1" method =post action="MOA00006.aspx" >
    <input type="hidden" name="stepsid">
    <input type="hidden" name="group_id">
    <!-- group id for drag to flow chart ; R1=上一級主管; A1=申請者的上一級主管   uid-3 : group id; uid-0:department id; uid-1:employee;uid-2:group leader-->
    <input type="hidden" name="group_name">
    <input type="hidden" name="MODE">
    <input type="hidden" name="scrollLeftPixel" value="<%=scrollLeftPixel%>">
    <input type="hidden" name="scrollTopPixel" value="<%=scrollTopPixel%>">
    <input type="hidden" name="organization_id" value="<%=organization_id%>" />    
    <input type="hidden" name="eformid" value="<%=eformid%>" />  
    <input type="hidden" name="eformrole" value="<%=eformrole%>" />       
    <input type="hidden" id="frm_chinese_name" name="frm_chinese_name" value="<%=frm_chinese_name%>" />    
    <input type="hidden" name="cc">  
    <% If save_flag = "True" Then %>
       <input type="button" class="buttons" onclick="reflow()" value="重新規劃" id="button2" name="button2">&nbsp;    
       <input type="button" class="buttons" onclick="all_finish()" value="完成" id="btn_finish" name="btn_finish">&nbsp;
    <% End If %>   
       
        </form>
        
    <div ID="fake" STYLE="border: solid 1px black;font-size:9pt; position:absolute;Z-INDEX:100; display:none; background-color:#ffcccc; text-align:center; color:#000000; width:80;  ">
</div>
<span style="top:3000px;left:2000;position:absolute;">&nbsp;
</span>
<!-- popup menu for flow node but 申請者除外-->
<!--<span ID="set1" STYLE="padding:0;font-size:9pt;position:absolute;width:65pt;Z-INDEX:103;display:none" onMouseOver="this.style.display=''" onMouseOut="this.style.display='none'">    -->
<div ID="set1" STYLE="border: solid 1px black;font-size:9pt; position:absolute;Z-INDEX:100; display:none; background-color:#ffcccc; text-align:center; color:#000000; width:80;  ">
   <table class="setbox" cellpadding="0" cellspacing="1" width="100%">        
		<tr><td>
		
			<table border="0" bgcolor="cccccc" cellpadding="2" cellspacing="0" width="100%">
                 <tr><td class="c2" align="center" colspan="2"><nobr>屬性設定</nobr></td></tr>
                 <tr align="center" class="nr" onmouseover="this.className='cl'" onmouseout="this.className='nr'" onmousedown="basicSetting()">
                    <td></td><td>基本設定</td>        
                 </tr>   
                 <tr align="center" class="nr" onmouseover="this.className='cl'" onmouseout="this.className='nr'" onmousedown="conditionFlow()">
                    <td></td><td>條件處理</td>      
                 </tr>     
                 <tr align="center" class="nr" onmouseover="this.className='cl'" onmouseout="this.className='nr'" onmousedown="userProcedure('1')">
                    <td></td><td>自訂程序</td>      
                 </tr>     
                 <tr align="center" class="nr" onmouseover="this.className='cl'" onmouseout="this.className='nr'" onmousedown="setMajor()">
                    <td></td><td>設主關卡</td>      
                 </tr>     
                 <tr align="center" class="nr" onmouseover="this.className='cl'" onmouseout="this.className='nr'" onmousedown="changeName()">
                    <td></td><td>更　　名</td>      
                 </tr>     
                 <tr align="center" class="nr" onmouseover="this.className='cl'" onmouseout="this.className='nr'" onmousedown="delStep()">
                    <td align="center"><img src="../../image/x.gif" WIDTH="14" HEIGHT="13"></td><td>刪除此關</td>
                 </tr>
			</table>
			
		</td></tr>
   </table>
  </div>
<!--</span>-->
<!-- popup menu for flow node 申請者 -->

<!--<span ID="set0" STYLE="padding:0;font-size:9pt;position:absolute;width:65pt;Z-INDEX:103;display:none">-->
<div ID="set0" STYLE="border: solid 1px black;font-size:9pt; position:absolute;Z-INDEX:100; display:none; background-color:#ffcccc; text-align:center; color:#000000; width:80;  ">

   <table class="setbox" cellpadding="0" cellspacing="1" width="100%">
   <tr><td>
			<table border="0" bgcolor="cccccc" cellpadding="2" cellspacing="0" width="100%">
         <tr><td class="c2" align="center"><nobr>屬性設定</nobr></td></tr>
         <tr><td class="nr" onmouseover="this.className='cl'" onmouseout="this.className='nr'" onmousedown="basicSetting()" align="center">基本設定</td></tr>
         <tr><td class="nr" onmouseover="this.className='cl'" onmouseout="this.className='nr'" onmousedown="conditionFlow()" align="center">條件處理</td></tr>
         <tr><td class="nr" onmouseover="this.className='cl'" onmouseout="this.className='nr'" onmousedown="userProcedure('0')" align="center">自訂程序</td></tr>
			</table>
		</td></tr>
   </table>
 </div>
<!--</span>-->
<!-- popup menu for flow node 結束 -->

<!--<span ID="Span1" STYLE="padding:0;font-size:9pt;position:absolute;width:65pt;Z-INDEX:103;display:none">  -->  
<div ID="Span1" STYLE="border: solid 1px black;font-size:9pt; position:absolute;Z-INDEX:100; display:none; background-color:#ffcccc; text-align:center; color:#000000; width:80;  ">
   <table class="setbox" cellpadding="0" cellspacing="1" width="100%">        
		<tr><td>
		
			<table border="0" bgcolor="cccccc" cellpadding="2" cellspacing="0" width="100%">
         <tr><td class="c2" align="center" style="height: 18px"><nobr>屬性設定</nobr></td></tr>
         <tr><td class="nr" onmouseover="this.className='cl'" onmouseout="this.className='nr'" onmousedown="userProcedure('-1')" align="center" style="height: 18px">自訂程序</td></tr>
			</table>
			
		</td></tr>
   </table>
 </div>
<!--</span>-->

<!-- start flow here -->
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
                Response.Write("/><br/><nobr>")
                'If Request.Cookies("global")("language") = "EN" Or Request.Cookies("global")("language") = "JA" Then
                '    Response.Write("Finish")
                'Else
                Response.Write("結束")
                'End If
                Response.Write("</nobr>")
                Response.Write("</span>")
                Response.Write("<img STYLE='position:absolute;top:" & CStr(CInt(Trim(dr("y").ToString)) + 44) & "px;left:" & CStr(CInt(Trim(dr("x").ToString)) - 20) & "px;width:80px;z-index:1' src='../../image/tryA.gif' WIDTH=81 HEIGHT=4/>")
                Response.Write("<img STYLE='position:absolute;top:" & CStr(CInt(Trim(dr("y").ToString)) + 45) & "px;left:" & CStr(CInt(Trim(dr("x").ToString)) - 20) & "px;width:81px;height:18px;' src='../../image/L1.gif' WIDTH=81 HEIGHT=38/>")
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
                    Response.Write("<span onMouseOver=""javascipt:switchdisplay('" & Perid & "',true);"" onMouseOut=""javascipt:switchdisplay('" & Perid & "',false);"" STYLE='padding:0;position:absolute;width=80;height:25px;top:" & CStr(CInt(Trim(dr("y").ToString)) - 12) & "px;left:" & CStr(CInt(Trim(dr("x").ToString)) - 30) & "px;text-align:center;'>")
                Else
                    Response.Write("<span STYLE='padding:0;position:absolute;width=80;height:25px;top:" & CStr(CInt(Trim(dr("y").ToString)) - 12) & "px;left:" & CStr(CInt(Trim(dr("x").ToString)) - 30) & "px;text-align:center;'>")
                End If
                
                Response.Write("   <img ID='I" & Trim(dr("stepsid").ToString) & "' style='CURSOR: hand' ")
                
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
                    Response.Write("javascipt:switchdisplay('" & Perid & "',false);")
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

<script language="javascript" src="../Inc/inc_dialogDown.js"></script>
<script language="javascript">
var curElement;
var curGroupName;
var notcorr;
var faketitle;
var PY = 0;
var PX = 0;
var apk = 0;
var setPopup = 1;
var set1cont = set1.innerHTML;
//document.onmouseup = doMouseUp;
//document.onmouseover= doMouseMove;
//document.onmousedown=doMouseDown;
//document.onmousemove= doMouseMove;
//document.ondrop = doOnDrop;
//document.ondragenter=doDragEnter;
//document.ondragover=doDragOver;
//document.ondragleave=doDragLeave;
//document.ondragend=doDragEnd;
//window.onresize = doresize;
window.onscroll = doscroll;
function ini(){
      window.scrollTo(form1.scrollLeftPixel.value,form1.scrollTopPixel.value);
      flowbody.style.display = '';
      PY = document.body.scrollTop;
      PX = document.body.scrollLeft;
      <%if err_msg <>"" then  %>  
          alert('<%= err_msg.Replace("'", "") %>');
      <%  err_msg="" 
         end if %>
}

function doscroll(){
      PY = document.body.scrollTop;
      PX = document.body.scrollLeft;
      //folder.style.top  = PY;
      //folder.style.left = PX + xt ;
      //strbar.style.top  = PY;
      //strbar.style.left = PX;
      form1.scrollLeftPixel.value = PX;
      form1.scrollTopPixel.value = PY;
}
function delStep() {

    set1.style.display = 'none';
    run.style.display="none";
    if (confirm('確定要刪除此關?')){
        form1.stepsid.value = curElement1;
        form1.MODE.value = 'DEL';
        form1.scrollLeftPixel.value = PX;
        form1.scrollTopPixel.value = PY;
        form1.submit();
    } 
}
function delCC(stepsid, group_id) {
    document.all('CC'+stepsid).style.display = 'none';
    if (confirm('確定要刪除此關?')){
        form1.stepsid.value = stepsid;
        form1.MODE.value = 'DELCC';
        form1.cc.value = group_id;
        form1.scrollLeftPixel.value = PX;
        form1.scrollTopPixel.value = PY;
        form1.submit();
    } 
}
function changeName() {
   set1.style.display = 'none';
	var GroupName = prompt("請輸入新的名稱", curGroupName);
	if (GroupName != "" && GroupName != null)
	{
        form1.stepsid.value = curElement1;
        form1.MODE.value = 'CHANGENAME';
        form1.group_name.value=GroupName;
        form1.scrollLeftPixel.value = PX;
        form1.scrollTopPixel.value = PY;
        form1.submit();
    } 
}
function reflow(){
     var save_flag;
    save_flag="<%=save_flag %>";
     if (save_flag =='False')
      {return ;}
    if (confirm('確定要重新規劃此表單流程？')){
          document.all.MODE.value='NEW';
          form1.submit();
   }
}
function all_finish(){     
    if (confirm('確定要完成此表單流程？')){          
          document.all.MODE.value='finish';
          form1.submit();
   }
}
function showPopup() {
    curElement0 = event.srcElement.id;
    curElement1 = curElement0.substring(1, curElement0.length);
    curGroupName=event.srcElement.title.substring(0,event.srcElement.title.indexOf(' ('));
    switch (setPopup) {
    case 0: 
          set1.innerHTML = set0.innerHTML;
          break;
    case 1:
          set1.innerHTML = set1cont;
          break;
    case -1:
          set1.innerHTML=set0.innerHTML;
          break;
    }
    setPopup = 1;
    set1.style.display = '';
    var iTopPosition=event.clientY + PY-5;
    if (event.clientY+145 > document.body.offsetHeight) 
        iTopPosition=document.body.offsetHeight +PY - 150;
    set1.style.pixelTop= iTopPosition;
    set1.style.pixelLeft = event.clientX + PX - 5;
}
function showPopup0() {
  setPopup = 0;
  showPopup();
}
function showPopup_1() {
  setPopup = -1;
  showPopup();
}
function basicSetting() {
   form1.stepsid.value = curElement1;
   form1.scrollLeftPixel.value = PX;
   form1.scrollTopPixel.value = PY;
   form1.action='DefineFlowBasicSetting.asp';
   form1.submit();
}
function conditionFlow() {
   form1.stepsid.value = curElement1;
   form1.scrollLeftPixel.value = PX;
   form1.scrollTopPixel.value = PY;
   form1.action='DefineFlowCondition.asp';
   form1.submit();
}
function userProcedure(nodetype) {
   form1.stepsid.value = curElement1;
   form1.scrollLeftPixel.value = PX;
   form1.scrollTopPixel.value = PY;
   form1.action='DefineFlowUserProcedure.asp?nodetype='+nodetype;
   form1.submit();
}
function setMajor() {
    form1.stepsid.value = curElement1;
    form1.MODE.value = 'MAJOR';
    form1.scrollLeftPixel.value = PX;
    form1.scrollTopPixel.value = PY;
    form1.submit();
}
function fOnDragOver() {
   event.returnValue=false;
	if (window.event.dataTransfer.getData("Text"))
	{
      event.dataTransfer.dropEffect="copy";
      doMouseMove();
   }
}
function fOnDrop() {
	if (window.event.dataTransfer.getData("Text"))
	{
       doMouseUp();
   }
}
function doMouseMove() {
   if (curElement!=null){
     parent.FlowLeft.fake.style.display='none';
     fake.style.posLeft = window.event.clientX + PX;
     fake.style.posTop = window.event.clientY + PY;
     if (1==2) {
     }
   //start node control here  
<%
	sqlstrQT = "select * from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " order by y asc, x desc"
	If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = True Then
	   n_table = do_sql.G_table
	   
	   For Each dr In n_table.Rows           
        Dim str_stmt As string
	    sqlstrQT = "select count(*) as nodecount from tempflow where eformid='" & eformid & "' and eformrole=" & eformrole & " and y=" & Trim(dr("y").ToString) 
	    
        If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = false Then
           exit for
        end if
        nodecount= cint(do_sql.G_table.Rows(0).item("nodecount").tostring)
        
        str_stmt = "    else if (fake.style.posTop >=" & cstr(cint(Trim(dr("y").ToString)) - 80) & " && fake.style.posTop <= " & cstr(cint(Trim(dr("y").ToString)) + 20) 
        if nodecount > 1 then
           str_stmt += " && fake.style.posLeft >=" &  cstr(cint(Trim(dr("x").ToString)) - 40) & " && fake.style.posLeft <= " &  cstr(cint(Trim(dr("x").ToString)) + 80)
        end if  
        str_stmt += "){" & chr(13)
        response.write (str_stmt)
        
        response.write ("        if (1==2) {" & chr(13))
        response.write ("        }" & chr(13) )
        response.write ("        else if ( fake.style.posLeft >=" &  cstr(cint(Trim(dr("x").ToString)) - 40) & " && fake.style.posLeft <= " &  cstr(cint(Trim(dr("x").ToString)) + 80) & " ){" & chr(13)  )
        response.write ("                  notcorr = 0;" & chr(13)  )
        response.write ("                  form1.stepsid.value = '" & Trim(dr("stepsid").ToString)  & "';" & chr(13) )
        response.write ("                  if (fake.style.posTop >=" & cstr(cint(Trim(dr("y").ToString)) - 80) & " && fake.style.posTop < " & cstr(cint(Trim(dr("y").ToString)) - 20) & " && fake.style.posLeft >=" & cstr(cint(Trim(dr("x").ToString)) - 20) & " && fake.style.posLeft <= " & cstr(cint(Trim(dr("x").ToString)) + 40) & ") {" & chr(13) )
        select case Trim(dr("steps").ToString) 
        case 0 '申請人
           response.write ("               fake.style.backgroundColor ='cyan'" & chr(13)  )
           response.write ("               fake.innerHTML = '串簽';" & chr(13) )
           response.write ("               form1.MODE.value ='ASER';" & chr(13) )
        case -1 '結束
           response.write ("               notcorr = 1" & chr(13) )
           response.write ("               fake.style.display = ''" & chr(13) )
           response.write ("               fake.style.backgroundColor = '#ffcccc'" & chr(13) )
           response.write ("               fake.innerHTML = faketitle " & chr(13) )
           response.write ("           }" & chr(13) )
           response.write ("           else {" & chr(13) )
           response.write ("               if (fake.style.posLeft >=" & cstr(cint(Trim(dr("x").ToString)) - 40) & " && fake.style.posLeft <= " & cstr(cint(Trim(dr("x").ToString)) + 10) & "){" & chr(13) )
           response.write ("                   notcorr = 1" & chr(13) )
           response.write ("                   fake.style.display = ''" & chr(13) )
           response.write ("                   fake.style.backgroundColor = 'ffcccc'" & chr(13) )
           response.write ("                   fake.innerHTML = faketitle " & chr(13) )
           response.write ("               }" & chr(13) )
           response.write ("               if (fake.style.posLeft >=" & cstr(cint(Trim(dr("x").ToString)) + 10) & " && fake.style.posLeft <= " & cstr(cint(Trim(dr("x").ToString)) - 40) & "){" & chr(13) )
           response.write ("                   fake.style.backgroundColor ='chartreuse';" & chr(13) )
           response.write ("                   fake.innerHTML = '通知';" & chr(13) )
           response.write ("                   form1.MODE.value ='CC';" & chr(13) )
           response.write ("               }" & chr(13) )
        case else
           response.write ("               fake.style.backgroundColor ='cyan'" & chr(13) )
           response.write ("               fake.innerHTML = '串簽';" & chr(13) )
           response.write ("               form1.MODE.value ='ASER';" & chr(13) )
           response.write ("           }" & chr(13) )
           response.write ("           else {" & chr(13) )
           response.write ("               if (fake.style.posLeft >=" & cstr(cint(Trim(dr("x").ToString)) - 40) & " && fake.style.posLeft <= " & cstr(cint(Trim(dr("x").ToString)) + 10) & "){" & chr(13) )
           response.write ("                   fake.style.backgroundColor ='ffff00';" & chr(13) )
           response.write ("                   fake.innerHTML = '會簽';" & chr(13) )
           response.write ("                   form1.MODE.value ='PAR';" & chr(13) )
           response.write ("               }" & chr(13) )
           response.write ("               if (fake.style.posLeft >=" & cstr(cint(Trim(dr("x").ToString)) + 10) & " && fake.style.posLeft <= " & cstr(cint(Trim(dr("x").ToString)) + 40) & "){" & chr(13) )
           response.write ("                   fake.style.backgroundColor ='chartreuse';" & chr(13) )
           response.write ("                   fake.innerHTML = '通知';" & chr(13) )
           response.write ("                   form1.MODE.value ='CC';" & chr(13) )
           response.write ("               }" & chr(13) )
        end select
        response.write ("               }" & chr(13) )
        response.write ("      }" & chr(13) )
        response.write ("      else {" & chr(13) )
        response.write ("          notcorr = 1" & chr(13) )
        response.write ("          fake.style.display = ''" & chr(13) )
        response.write ("          fake.style.backgroundColor = '#ffcccc'" & chr(13) )
        response.write ("          fake.innerHTML = faketitle " & chr(13) )
        response.write ("      }" & chr(13) )
        response.write ("}" & chr(13) )
        
        next
     
     end if
 %>       
 //       
     else {
            notcorr = 1
            fake.style.backgroundColor = '#ffcccc'
            fake.innerHTML = faketitle 
     }
     fake.style.display = '';
   }
   //event.returnValue = true;
   //event.cancelBubble=true;
}
function doMouseDown() {
   window.status="doMouseDown";
}
function doMouseUp() {
   fake.style.display = 'none';
   parent.FlowLeft.fake.style.display = 'none';   
   if (curElement!=null){
      curElement = null;
      parent.FlowLeft.curElement=null;
      if (notcorr == 1 ){
          fake.style.display = 'none';
          curElement = null;
          alert ('請拖拉至某關上方' + unescape('%0D%0A') + '變色時為正確位置')
      return true;
      }      
      form1.scrollLeftPixel.value =  PX;
      form1.scrollTopPixel.value =  PY;

	  var oGroupID = document.all.group_id;
      var sMode = oGroupID.value.substr(oGroupID.value.indexOf('_')+1);
      var sType = oGroupID.value.substr(0, oGroupID.value.indexOf('_'));
      switch (sMode) {
      case "Org":
      case "Prj":
      case "LDR":
		  clipDialog(sType, sMode);
		  break;
	  default :		      
	      form1.submit();	      	      
	      break;	
      }
   }
}

function clipDialog(addType,addMode) {
	if (document.readyState!='complete' || document.all.Itemifrmdiv==null){
		alert('文件尚未載入完畢');
		return false;
	}
    
   var src; 
	switch (addMode){
	case "Org":
		switch (addType){
		//行政部門/主管
		case "O1":
		case "O2":
			
			src = "DefineFlowAddGroup.asp";
			break;
		//所有員工
		case "O3":
		   
			src = "DefineFlowAddEmployee.asp?iType=Org";
			break;
		//部門關係人
		case "O4":
		   
			src = "DefineFlowAddRole.asp?iType=Org";
			break;
		}
		break;
	case "Prj":
		switch(addType){
		//專案部門/主管
		case "P1":
		case "P2":
		   
			src = "DefineFlowAddPrjGroup.asp?organization_id=<%=organization_id%>";
			break;
		//專案所有員工
		case "P3":
		   
			src = "DefineFlowAddEmployee.asp?iType=Prj&organization_id=<%=organization_id%>";
			break;
		case "P4":
		   	src = "DefineFlowAddRole.asp?iType=Prj";
			break;
		}	
		break;
	case "LDR":
	   
		src = "DefineFlowAddLeader.asp";
		break;
	}
	  
   if (window.screen.width < 835) { // 800x600
      document.all.Itemifrmdiv.style.pixelHeight=300;
      document.all.Itemifrmdiv.style.pixelWidth=240;
   } else {
      document.all.Itemifrmdiv.style.pixelHeight=320;
      document.all.Itemifrmdiv.style.pixelWidth=250;
   }
   document.all.Itemifrm.style.pixelHeight=document.all.Itemifrmdiv.style.pixelHeight;
   document.all.Itemifrm.style.pixelWidth=document.all.Itemifrmdiv.style.pixelWidth;
   DialogDown(src);	    
   
}

function SubmitForm() {
	 form1.submit();
}

function chg_s100() {
     
	 s100.style.color='blue';
	 s200.style.color='black';
	 s900.style.color='black';	 
}
function test2() {
	 alert("bbb");
}
function chg_s200() {   
    
	 s100.style.color='black';
	 s200.style.color='blue';
	 s900.style.color='black';	 
	 //run.style.display="block";
}
function chg_s900() {   
    
	 s100.style.color='black';
	 s200.style.color='black';
	 s900.style.color='blue';	 
	 //run.style.display="block";
}
function bock_div() {   
    run.style.display="none";
}
function show_div() {
    var save_flag;
    save_flag="<%=save_flag %>";    
     if (save_flag =='False')
      {return ;}
    curElement0 = event.srcElement.id;
    curElement1 = curElement0.substring(1, curElement0.length);
    curGroupName=event.srcElement.title.substring(0,event.srcElement.title.indexOf(' ('));
    
    setPopup = 1;
    run.style.left =event.clientX + PX - 5;
    run.style.top = event.clientY-40;    
    run.style.display="block";
}


</script>
<!-- finish rendering groupitem --><!-- iframe for property -->
<div id="Itemifrmdiv" style="DISPLAY: none; POSITION: absolute; HEIGHT: 310px; LEFT: -10px; TOP: -510px; WIDTH: 300px; Z-INDEX: 6;font-size:9pt;line-height:14pt" align="center">
<table width="100%" cellspacing="0" cellpadding="0"><tr><td>
<iframe id="Itemifrm" name="Itemifrm" height="310" width="300" frameborder="0" scrolling="no" marginheight="0" marginwidth="0">
</iframe></td></tr></table>
</div>
<div id="ItemifrmdivShadow" style="BACKGROUND-COLOR: #999999; DISPLAY: none; POSITION: absolute; Z-INDEX: 5;overflow:auto"></div>
    
     <div id="run" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:90pt; height:50pt; left:300; top:100; display:none;">
<br>
<table border="1";width=90;>
       <tr>
        <td>
           <ol id="ol2">
      <li id="s100" onclick=delStep();   onmouseover=chg_s100() >刪除此關</li>
      <li id="s200" onclick=test2();  onmouseover=chg_s200()>測試此關</li>
      <li id="s900" onclick=bock_div();  onmouseover=chg_s900()>取消</li>
     </ol>
        </td>
       </tr>
        
     
     </table>
</div>
    
</body>
</html>
