<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA02005.aspx.vb" Inherits="Source_02_MOA02005" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>會議室新增</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />    
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label6" runat="server" CssClass="toptitle" Text="會議室管理" Width="100%"></asp:Label>
            </td></tr>
        </table>         
        <table id="Table1" border="1" cellpadding="1" cellspacing="1" style="z-index: 101;left: 137px; position: absolute; top: 67px; height: 216px" width="600">
            <tr>
            <td style="width: 25%"><FONT color="#ff0033">*</FONT><asp:Label ID="Label1" runat="server" Text="會議室名稱" CssClass="form"></asp:Label></td>
            <td style="width: 70%"><asp:textbox id="RoomName" runat="server" Width="207px"></asp:textbox></td>
            </tr>
            <tr>
            <td style="width: 25%"><FONT color="#ff0033">*</FONT><asp:Label ID="Label2" runat="server" Text="管理員" CssClass="form"></asp:Label></td>
            <td style="width: 70%"><asp:dropdownlist id="Manager" runat="server" Width="143px" DataSourceID="SqlDataSource1" DataTextField="emp_chinese_name" DataValueField="employee_id"></asp:dropdownlist><asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                SelectCommand="SELECT [employee_id], [emp_chinese_name] FROM [V_EmpInfo] WHERE ([ORG_UID] = @ORG_UID) ORDER BY [emp_chinese_name]">
                <SelectParameters>
                    <asp:SessionParameter Name="ORG_UID" SessionField="org_uid" Type="String" />
                </SelectParameters>
            </asp:SqlDataSource>
            </td>
            </tr>
            <tr>
            <td style="width: 25%; height: 28px;"><FONT color="#ff0033">*</FONT><asp:Label ID="Label3" runat="server" Text="聯絡電話" CssClass="form"></asp:Label></td>
            <td style="width: 70%; height: 28px;"><asp:textbox id="tel" runat="server" Width="123px"></asp:textbox></td>
            </tr>
            <tr>
            <td style="width: 25%"><asp:Label ID="Label4" runat="server" Text="容納人數" CssClass="form"></asp:Label></td>
            <td style="width: 70%"><asp:textbox id="ContainNum" runat="server" Width="122px"></asp:textbox></td>
            </tr>
            <tr>
            <td style="width: 25%"><asp:Label ID="Label5" runat="server" Text="共用與否" CssClass="form"></asp:Label></td>
            <td style="width: 70%"><asp:dropdownlist id="Share" runat="server" Width="143px">
                <asp:ListItem Value="1">共用</asp:ListItem>
                <asp:ListItem Value="2">不共用</asp:ListItem>
            </asp:dropdownlist>
            </td>
            </tr>
            <tr>
            <td width="30%" colspan="2" align="center">
                <asp:ImageButton ID="ImgOK" runat="server" ImageUrl="~/Image/apply.gif" />
                <asp:ImageButton ID="BackBtn" runat="server" AlternateText="回上頁" ImageUrl="~/Image/backtop.gif" /></td>
            </tr>
        </table>
    <div>
    
    </div>
    </form>
</body>
</html>
