<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00100.aspx.vb" Inherits="Source_00_MOA00100" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>年度行事曆</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
<form id ='MOA00100' action="MOA00100.aspx" runat="server">
<table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px; HEIGHT: 183px">
<tr><td align='center'>
    <asp:Label ID="Label1" runat="server" CssClass="toptitle" Text="年度行事曆" Width="100%"></asp:Label>
</td></tr><tr ><td align='center'>
    <asp:Label ID="Lab1" runat="server" Text="選擇年度：" ></asp:Label>
    <asp:DropDownList id="Sel_Y" AutoPostBack="False" runat="server">
    </asp:DropDownList><hr />
</td></tr><tr ><td align='center'>
    <asp:Button ID="Button1" runat="server" Text="年度行事曆設定" /><hr />
</td></tr><tr><td align='center'>
    <asp:Button ID="Button2" runat="server" Text="年度假日重設"  OnClientClick="return confirm('確定重設年度假日嗎?')"/><hr />
</td></tr><tr ><td align='center'>
    <asp:Button ID="Button3" runat="server" Text="年度行事曆匯出" /><hr />
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
