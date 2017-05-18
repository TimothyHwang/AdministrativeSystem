<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00033.aspx.vb" Inherits="Source_00_MOA00033" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>單位新增</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr>
            <td align="center">
                    <asp:Label ID="LabelTitle" runat="server" CssClass="toptitle" Text="單位新增" Width="100%"></asp:Label>
            </td>
            </tr>
        </table>
        <table width="100%" border="3" bordercolor="#ccddee" align="center">
            <tr>
            <td align="center" class="CellClass" style="width: 25%">
                <FONT color="#ff0033">*</FONT><asp:Label ID="Label1" runat="server" Text="單位代號" ></asp:Label>
            </td>
            <td style="width: 75%">
                <asp:textbox id="txtorg_id" runat="server" Width="161px"></asp:textbox></td>
            </tr>
            
            <tr>
            <td align="center" class="CellClass" style="width: 25%">
                <FONT color="#ff0033">*</FONT><asp:Label ID="Label2" runat="server" Text="單位名稱"></asp:Label>
            </td>
            <td style="width: 75%">
                <asp:textbox id="txtorg_name" runat="server" Width="200px"></asp:textbox></td>
            </tr>
            
            <tr>
            <td align="center" class="CellClass" style="width: 25%">
                <asp:Label ID="Label4" runat="server" Text="單位英文名稱"></asp:Label>
            </td>
            <td style="width: 75%">
                <asp:textbox id="txtorg_enname" runat="server" Width="122px"></asp:textbox></td>
            </tr>
            
            <% If Request.QueryString("UseFlag") = "2" Then%>
            <tr>
            <td align="center" class="CellClass" style="width: 25%; height: 26px;">
                <asp:Label ID="Label5" runat="server" Text="上一級單位"></asp:Label>
            </td>
            <td style="width: 75%; height: 26px;">
                <asp:DropDownList ID="DropOrgUp" runat="server" DataSourceID="SqlDataSource1" DataTextField="ORG_NAME" DataValueField="ORG_UID">
                </asp:DropDownList></td>
            </tr>
            <% End If%>
            
            <tr>
            <td align="center" class="CellClass" style="width: 25%">
                <asp:Label ID="Label3" runat="server" Text="單位級別"></asp:Label>
            </td>
            <td style="width: 75%">
                <asp:DropDownList ID="DropOrgKind" runat="server">
                    <asp:ListItem Value="0">無</asp:ListItem>
                    <asp:ListItem Value="1">一級單位</asp:ListItem>
                    <asp:ListItem Value="2">二級單位</asp:ListItem>
                </asp:DropDownList></td>
            </tr>            
            
            <tr>
            <td align="center" class="CellClass" colspan="2">
                <asp:ImageButton ID="btnImgOK" runat="server" ImageUrl="~/Image/apply.gif" ToolTip="確認" />
                <asp:ImageButton ID="BackBtn" runat="server" AlternateText="回上頁" ImageUrl="~/Image/backtop.gif" /></td>
            </tr>
        </table>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] ORDER BY [ORG_NAME]">
        </asp:SqlDataSource>
        
    </form>
</body>
</html>
