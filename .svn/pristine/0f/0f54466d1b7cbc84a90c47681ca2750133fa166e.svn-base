<%@ Page Language="VB" EnableEventValidation = "false" AutoEventWireup="true" CodeFile="MOA04102.aspx.vb" Inherits="M_Source_04_MOA04102" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="zh-tw" http-equiv="Content-Language" />
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
<title></title>
</head>
<body>
<form id="form1" runat="server" >
<center>
<asp:label runat="server" id="lb_EFORMSN" Visible = "false"/>
<asp:Panel ID="PanelContent" runat="server">
<asp:FormView ID="FormView1" runat="server" DataSourceID="sdsDataRecords" EnableViewState="False" >
    <ItemTemplate>
<table cellpadding="13"  border="1" cellpadding="0" cellspacing="0" bordercolor="black" style="width: 140mm; height: 300mm; line-height: 1px; vertical-align: 1%; text-align: left;">
	<tr style="border-color: black;height:40px;">
		<td colspan="4" align = "center" style="font-family: 標楷體; font-size: 36px; ">房 舍 水 電 報 修 單</td>
	</tr>
	<tr style="height:30px;">
		<td style="font-family: 標楷體;font-size: 14px">報修單號：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3"> &nbsp;<%# Eval("EFORMSN")%></td>
	</tr>
	<tr style="height:30px;">
		<td style="font-family: 標楷體;font-size: 14px">填表人單位：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3"> &nbsp;<%# Eval("PWUNIT")%></td>
	</tr>
	<tr style="height:30px;">
		<td style="font-family: 標楷體;font-size: 14px">填表人姓名：</td>
		<td style="font-family: 標楷體;font-size: 14px"> &nbsp;<%# Eval("PWNAME") %></td>
		<td style="font-family: 標楷體;font-size: 14px">級職：</td>
		<td style="font-family: 標楷體;font-size: 14px"> &nbsp;<%# Eval("PWTITLE")%></td>
	</tr>
	<tr style="height:30px;">
		<td style="font-family: 標楷體;font-size: 14px">申請人單位：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3"> &nbsp;<%# Eval("PAUNIT")%></td>
	</tr>
	<tr style="height:30px;">
		<td style="font-family: 標楷體;font-size: 14px">申請人姓名：</td>
		<td style="font-family: 標楷體;font-size: 14px"> &nbsp;<%# Eval("PANAME")%></td>
		<td style="font-family: 標楷體;font-size: 14px">級職：</td>
		<td style="font-family: 標楷體;font-size: 14px"> &nbsp;<%# Eval("PATITLE")%></td>
	</tr>
	<tr style="height:30px;">
		<td style="font-family: 標楷體;font-size: 14px">申請時間：</td>
		<td style="font-family: 標楷體;font-size: 14px"> &nbsp;<%# Eval("nAPPTIME")%></td>
		<td style="font-family: 標楷體;font-size: 14px">連絡電話：</td>
		<td style="font-family: 標楷體;font-size: 14px"> &nbsp;<%# Eval("nPHONE")%></td>
	</tr>
	<tr style="height:30px;">
		<td style="font-family: 標楷體;font-size: 14px">請修地點：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3"> &nbsp;<%# Eval("location") %></td>
	</tr>
	<tr style="height:160px;">
		<td style="font-family: 標楷體;font-size: 14px">請修事項：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3" > &nbsp;<%# Eval("nFIXITEM") %></td>
	</tr>
	<tr style="height:30px;">
		<td style="font-family: 標楷體;font-size: 14px">現勘人員：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3">&nbsp;</td>
	</tr>
	<tr style="height:30px;">
		<td style="font-family: 標楷體;font-size: 14px">現勘日期：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3">&nbsp;</td>
	</tr>
	<tr style="height:160px;">
		<td style="font-family: 標楷體;font-size: 14px">原因分析：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3">&nbsp;</td>
	</tr>
	<tr style="height:160px;">
		<td style="font-family: 標楷體;font-size: 14px" >備註：</td>
		<td colspan="3" valign ="top" style="font-family: 標楷體;font-size: 14px"> &nbsp;□ 修繕  &nbsp;□ 不修繕 
		</td>
	</tr>
    
</table>
    <table style="width: 140mm; height: 30px; >
    <tr style="height:30px;">
    <td align = "right" style="font-family: 標楷體;font-size: 14px" >列印日期:<%= printdate %></td></tr>
    </table>
   </ItemTemplate>
</asp:FormView>

</asp:Panel> 
</center>
<asp:SqlDataSource ID="sdsDataRecords" runat="server"
            ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="select a.* ,b.bd_name+' / ' + c.fl_name + ' / ' + d.rnum_name as location
                        from P_0415 a left join P_0404 b on a.nbd_code = b.bd_code
                        left join P_0406 c on a.nfl_code = c.fl_code
                        left join P_0411 d on a.nrnum_code = d.rnum_code
                        where eformsn=@EFORMSN;">
            <SelectParameters>
                <asp:ControlParameter ControlID="lb_EFORMSN" Name="EFORMSN" PropertyName="Text" type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
        
</form>
</body>
</html>
