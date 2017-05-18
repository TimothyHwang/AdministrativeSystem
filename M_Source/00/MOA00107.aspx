<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00107.aspx.vb" Inherits="Source_00_MOA00107" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>單位行事曆</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
<script language="javascript">
window.onload = function() {
<%
    Dim cnt As Integer = dv.Count()
    Dim i As Integer = 0
    Do While i < cnt
 %>
    document.all.D<%=CDate(dv.Table.Rows(i)(0)).ToString("yyyyMMdd")  %>.style.color="<%= butColor2 %>";
<%
    i+=1
    Loop
 %>
}
function  CalendarChage(){
    MOA00107.Date.value=document.all.yy.value + "/" + document.all.mm.value + "/1";
    MOA00107.action="MOA00107.aspx";
}
</script>
</head>
<body background="../../Image/main_bg.jpg">
<form id="MOA00107" runat="server" action="MOA00107.aspx" method="get">
<input type="hidden" name="Date" value="" />
<table  border="0" width="100%" style="Z-INDEX: 107; LEFT: 104px; TOP: 33px; HEIGHT: 183px">
<tr><td align='center' colspan='3'>
<asp:Label ID="Label1" runat="server" CssClass="toptitle" Text="單位行事曆" Width="100%"></asp:Label>
</td></tr>
<tr align="center" style="height:220;"><td>
<asp:PlaceHolder ID="PlaceHolder1" runat="server" ></asp:PlaceHolder>
</td></tr>
</table>
<asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
    SelectCommand=""
    DeleteCommand=""
    InsertCommand=""
    ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
</asp:SqlDataSource>
</form>
</body>
</html>

