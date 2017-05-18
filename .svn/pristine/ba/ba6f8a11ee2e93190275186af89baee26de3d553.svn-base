<%@ Page Language="VB" EnableEventValidation = "false" AutoEventWireup="true" CodeFile="MOA04103.aspx.vb" Inherits="M_Source_04_MOA04103" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
   <center>

<asp:Panel ID="PanelContent" runat="server">

<asp:FormView ID="FormView1" runat="server" DataSourceID="sdsDataRecords" EnableViewState="False" >
    <ItemTemplate>
<table cellpadding="18"  border="1" cellpadding="0" cellspacing="0" bordercolor="black" style="width: 140mm; height: 300mm; line-height: 0.1px; vertical-align: 1%; text-align: left;">
	<tr style="height:40px;">
		<td colspan="4" align = "center" style="font-family: 標楷體; font-size: 36px; ">房 舍 水 電 派 工 單</td>
	</tr>
	<tr style="height:30px;">
		<td style="font-family: 標楷體;font-size: 14px">報修單號：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3">&nbsp;<%# Eval("EFORMSN")%></td>
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
	<tr style="height:130px;">
		<td style="font-family: 標楷體;font-size: 14px">請修事項：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3" > &nbsp;<%# Eval("nFIXITEM") %></td>
	</tr>
	<tr style="height:30px;">
		<td style="font-family: 標楷體;font-size: 14px">現勘人員：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3">&nbsp;<%# Eval("nViewPerName")%></td>
	</tr>
	<tr style="height:30px;">
		<td style="font-family: 標楷體;font-size: 14px">現勘日期：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3">&nbsp;<%# Eval("nViewDate")%></td>
	</tr>
	<tr style="height:130px;">
		<td style="font-family: 標楷體;font-size: 14px">原因分析：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3">&nbsp;<%# Eval("nCause")%></td>
	</tr>
    <tr style="height:30px;">
		<td style="font-family: 標楷體;font-size: 14px">承辦類別：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3">&nbsp;<%# Eval("nExternal")%></td>
	</tr>
       <tr style="height:30px;">
		<td style="font-family: 標楷體;font-size: 14px">開工日期：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3">&nbsp;<%# Eval("nStartDATE")%></td>
	</tr>
    </tr>
       <tr style="height:30px;">
		<td style="font-family: 標楷體;font-size: 14px">設施(備)編號：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3">&nbsp;<%# Eval("nFacilityNo")%><br />&nbsp;<%# Eval("FacilityName") %></td>
	</tr>
    </tr>
       <tr style="height:30px;">
		<td style="font-family: 標楷體;font-size: 14px">故障原因：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3">&nbsp;<%# Eval("nErrCauseName")%></td>
	</tr>
      </tr>
       <tr style="height:30px;">
		<td style="font-family: 標楷體;font-size: 14px">派工人員：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3">&nbsp;<%# Eval("nPacthPer")%></td>
	</tr>
     <tr style="height:30px;">
		<td style="font-family: 標楷體;font-size: 14px">派工人數：</td>
		<td style="font-family: 標楷體;font-size: 14px" colspan="3">&nbsp;<%# Eval("nPacthCount")%></td>
	</tr>
</table>
   </ItemTemplate>
</asp:FormView>
<br />
<table style="width: 140mm; height: 30px;" >
    <tr>
        <td align="right" style="font-family: 標楷體;font-size: 14px">列印日期:<%= printdate %></td>
    </tr>
</table>
<br />
<table border="1" cellpadding="13" cellpadding="0" cellspacing="0" bordercolor="black" style="width: 140mm; line-height: 1px; vertical-align: 1%; text-align: center;">
    <tr style="border-color: black;height:45px;">
		<td colspan="4" align = "center" style="font-family: 標楷體; font-size: 36px; ">房 舍 水 電 領 料 單</td>
	</tr>
	<tr style="height:25px;">
		<td width = "20%" style="font-family: 標楷體;font-size: 14px">報修單號</td>
		<td width = "40%" style="font-family: 標楷體;font-size: 14px" >&nbsp;<asp:Label ID="lb_EFORMSN" runat="server" /></td>
        <td width = "15%" style="font-family: 標楷體;font-size: 14px">申請日期</td>
		<td width = "25%" style="font-family: 標楷體;font-size: 14px" > &nbsp;<asp:Label ID="lb_nAppStockDate" runat="server" /></td>
	</tr>

	<tr style="height:30px;">
		<td width = "20%" style="font-family: 標楷體;font-size: 14px">物品編號</td>
		<td width = "40%" style="font-family: 標楷體;font-size: 14px">品項名稱</td>
		<td width = "15%" style="font-family: 標楷體;font-size: 14px">單位</td>
		<td width = "25%"  style="font-family: 標楷體;font-size: 14px">儲庫編號</td>
	</tr>
    <asp:Repeater ID="RptitemList" runat="server" ><ItemTemplate>
    <tr style="height:30px;">
		<td width = "20%" style="font-family: 標楷體;font-size: 14px">&nbsp;<%# Eval("shcode")%></td>
		<td width = "40%" style="font-family: 標楷體;font-size: 14px">&nbsp;<%# Eval("it_name")%></td>
		<td width = "15%" style="font-family: 標楷體;font-size: 14px">&nbsp;<%# Eval("it_unit")%></td>
		<td width = "25%"  style="font-family: 標楷體;font-size: 14px">&nbsp;<%# Eval("seatnum")%></td>
	</tr>
     </ItemTemplate></asp:Repeater>
</table>
<table style="width: 140mm; height: 30px;" >
    <tr>
        <td align="right" style="font-family: 標楷體;font-size: 14px">
            列印日期:<%= printdate %></td>
    </tr>
</table>
</asp:Panel> 
<asp:SqlDataSource ID="sdsDataP_0415" runat="server"
            ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="select EFORMSN, nAppStockDate from P_0415 with(nolock)
                        where eformsn=@EFORMSN;">
            <SelectParameters>
                <asp:ControlParameter ControlID="lb_EFORMSN" Name="EFORMSN" PropertyName="Text" type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
    <asp:SqlDataSource ID="sdsDataP_0414" runat="server"
            ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="select shcode, b.seat_num+'_'+b.seat_name as seatnum,c.it_name,c.it_unit
                            from P_0414 a with(nolock) left join P_0417 b with(nolock) on a.seat_num = b.seat_num
                            left join P_0407 c with(nolock) on substring(a.shcode,1,6) = c.it_code
                            where job_Num = @EFORMSN">
            <SelectParameters>
                <asp:ControlParameter ControlID="lb_EFORMSN" Name="EFORMSN" PropertyName="Text" type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
</center>
<asp:SqlDataSource ID="sdsDataRecords" runat="server"
            ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="select a.* ,b.bd_name+' / ' + c.fl_name + ' / ' + d.rnum_name as location ,  f.House_Name as nViewPerName,
            case a.nErrCause when '1' then '人為因素' when '2' then '自然因素' when 3 then '維護查報' else '其他' end as nErrCauseName,
            b.bd_name as bdcode,c.fl_code as flcode,d.rnum_code as rnumcode,
            e.bd_name+'/'+e.fl_name+'/'+e.rnum_name+'/'+e.wa_name+'/'+e.bg_name+'/'+e.it_name as FacilityName
            from P_0415 a left join P_0404 b on a.nbd_code = b.bd_code
            left join P_0406 c on a.nfl_code = c.fl_code
            left join P_0411 d on a.nrnum_code = d.rnum_code
            left join P_0416 f on a.nViewPer = f.House_Num
            left join (
            select x.element_code,y.bd_name, z.fl_name,o.rnum_name,p.wa_name,q.bg_name,r.it_name
            from P_0405 x 
            left join P_0404 y on x.bd_code = y.bd_code
            left join P_0406 z on x.fl_code = z.fl_code
            left join P_0411 o on x.rnum_code = o.rnum_code
            left join P_0413 p on x.wa_code = p.wa_code
            left join P_0403 q on x.bg_code = q.bg_code
            left join P_0407 r on x.it_code = r.it_code) e on a.nFacilityNo = e.element_code
                        where eformsn=@EFORMSN;">
            <SelectParameters>
                <asp:ControlParameter ControlID="lb_EFORMSN" Name="EFORMSN" PropertyName="Text" type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
       
    </form>
</body>
</html>
