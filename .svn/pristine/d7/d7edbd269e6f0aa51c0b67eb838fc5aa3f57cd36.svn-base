<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00040.aspx.vb" Inherits="Source_00_MOA00040" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>人員管理</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label6" runat="server" CssClass="toptitle" Text="人員管理" Width="100%"></asp:Label>
            </td></tr>
        </table>       
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap width="10%" class="form">
                    部門名稱：</td>
			    <td noWrap class="form" style="height: 25px; width: 25%;">
                    <asp:DropDownList ID="OrgSel" runat="server" DataSourceID="SqlDataSource2" DataTextField="ORG_NAME"
                        DataValueField="ORG_UID">
                    </asp:DropDownList></td>
			    <td noWrap class="form" style="width: 5%">
                    人員帳號：</td>
			    <td noWrap class="form" style="width: 15%">
                    <asp:TextBox ID="empuid" runat="server" Width="70px"></asp:TextBox></td>
			    <td noWrap class="form" style="width: 5%">
                    姓名：</td>
			    <td noWrap class="form" style="height: 25px; width: 15%;"><asp:TextBox ID="emp_chinese_name" runat="server" Width="70px"></asp:TextBox></td>
			    <td noWrap class="form" style="width: 5%">
                    狀態：</td>
			    <td noWrap class="form" style="height: 25px; width: 10%;">
                    <asp:DropDownList ID="DropLeave" runat="server">
                        <asp:ListItem Value="">全部</asp:ListItem>
                        <asp:ListItem Value="Y" Selected="True">在職</asp:ListItem>
                        <asp:ListItem Value="N">離職</asp:ListItem>
                    </asp:DropDownList></td>
			    <td noWrap class="form" style="height: 25px; width: 10%;">
                <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/Search.gif" ToolTip="查詢" />
			    <asp:ImageButton ID="ImgBtn1" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" /></td>
		    </tr>
	    </table>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
            AutoGenerateColumns="False" DataSourceID="SqlDataSource1" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None" EmptyDataText="查無任何資料">
            <Columns>
                <asp:BoundField DataField="ORG_NAME" HeaderText="部門名稱" SortExpression="ORG_NAME" >
                    <HeaderStyle HorizontalAlign="Center" Width="27%" />
                </asp:BoundField>
                <asp:BoundField DataField="employee_id" HeaderText="帳號" SortExpression="employee_id" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="empemail" HeaderText="E-Mail" SortExpression="empemail" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                </asp:BoundField>
                <asp:BoundField DataField="emp_chinese_name" HeaderText="姓名" SortExpression="emp_chinese_name" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="leave" HeaderText="在職" SortExpression="leave" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                </asp:BoundField>
                <asp:BoundField DataField="AD_TITLE" HeaderText="職稱" SortExpression="AD_TITLE">
                    <HeaderStyle HorizontalAlign="Center" Width="22%" />
                </asp:BoundField>
                <asp:TemplateField>
                <ItemTemplate>
                        <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl='~/Image/update.gif'
                            PostBackUrl='<%# "MOA00043.aspx?empuid=" & eval("employee_id")%>' ToolTip="修改" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                </asp:TemplateField>
            </Columns>
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#E9ECF1" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [ORG_NAME], [empuid], [employee_id], [empemail], [emp_chinese_name], [leave], [AD_TITLE], [member_uid], [ErrCount] FROM [V_EmpInfo] WHERE leave='Y' ORDER BY [emp_chinese_name]">
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT * FROM [ADMINGROUP] ORDER BY [ORG_NAME]"></asp:SqlDataSource>
    </form>
</body>
</html>
