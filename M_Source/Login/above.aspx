<%@ Page Language="VB" AutoEventWireup="false" CodeFile="above.aspx.vb" Inherits="OA_above" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>框架上方</title>
</head>
	<body style="width:100%" background="../../Image/banner.jpg" >
		<form id="Form1" method="post" runat="server">
			<asp:button id="Logout" style="Z-INDEX: 101; LEFT: 90%; POSITION: absolute; TOP: 8px" runat="server" Text="登出"></asp:button>
			<font color="#ffffff">
			<b>
			    <asp:label id="Label2" style="Z-INDEX: 103; LEFT: 74%; POSITION: absolute; TOP: 17px" runat="server" Font-Size="Small" Font-Bold="True" Width="80px">使用者：</asp:label>
			    <asp:label id="UserName" style="Z-INDEX: 102; LEFT: 80%; POSITION: absolute; TOP: 16px" runat="server" Font-Size="Small" Width="150px"></asp:label>
			</b>
			<asp:button id="GoFirst" style="Z-INDEX: 101; LEFT: 74%; POSITION: absolute; TOP: 38px" runat="server" Text="回入口網站" Width="110px"></asp:button>
			<font color="#ffffff">
			</font>
		</form>
	</body>
</html>
