<%@ Page Language="VB" AutoEventWireup="false" CodeFile="index.aspx.vb" Inherits="index" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>行政入口</title>
    <link href="Styles.css" rel="stylesheet" type="text/css" />    
</head>
<body style="background-image:url('image/background.jpg'); background-repeat:no-repeat; background-color: #3366CC;" >
    <form id="form1" runat="server">
    <div>
			<table border="0" style="Z-INDEX: 101; LEFT: 592px; WIDTH: 218px; POSITION: absolute; TOP: 176px; HEIGHT: 96px">
				<tr>
					<td><asp:label id="Label1" runat="server" CssClass="form" >帳號：</asp:label></td>
					<td><asp:textbox id="UserName" runat="server" Width="156px"></asp:textbox></td>
				</tr>
				<tr>
					<td><asp:label id="Label2" runat="server" CssClass="form">密碼：</asp:label></td>
					<td><asp:textbox id="password" runat="server" TextMode="Password" Width="156px"></asp:textbox></td>
				<tr>
				<tr>
					<td align="center" colSpan="2"><asp:button id="login" runat="server" Text="登入"></asp:button></td>
				</tr>
				<tr>
					<td align="center" colSpan="2">
                        <asp:Label ID="loginerr" runat="server" ForeColor="Red"></asp:Label>
                    </td>
				</tr>
			</table>
      </div>
    </form>
</body>
</html>
