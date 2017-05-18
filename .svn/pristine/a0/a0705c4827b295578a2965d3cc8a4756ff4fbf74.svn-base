<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00061.aspx.vb" Inherits="Source_00_MOA00061" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>關卡人員管理</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">    
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label6" runat="server" CssClass="toptitle" Text="關卡人員管理" Width="100%"></asp:Label>
            </td></tr>
        </table>       
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap class="form" style="height: 25px; width: 10%;">
                    <asp:Label ID="Label1" runat="server" CssClass="form" Text="關卡名稱:" Width="77px"></asp:Label></td>
			    <td noWrap width="30%" class="form" style="height: 25px">			    
                    <asp:DropDownList ID="DropDownList1" runat="server" DataSourceID="SqlDataSource2" DataTextField="object_name" DataValueField="object_uid" Width="180px">
                    </asp:DropDownList></td>
			    <td noWrap class="form" style="height: 25px; width: 10%;">
                    <asp:Label ID="Label2" runat="server" CssClass="form" Text="人員名稱:" Width="77px"></asp:Label></td>
                    <td noWrap width="30%" class="form" style="height: 25px">
                        <asp:TextBox ID="PerName" runat="server" Width="100px"></asp:TextBox></td>
                <td noWrap width="10%" class="form" style="height: 25px">
                    <asp:ImageButton ID="searchbtn" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" /></td>                    
                <td noWrap width="10%" class="form" style="height: 25px">
			    <asp:ImageButton ID="ImgBtn1" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" />
                </td>
		    </tr>
	    </table>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True" EmptyDataText="查無任何資料"
            AutoGenerateColumns="False" DataSourceID="SqlDataSource1" Width="100%" DataKeyNames="object_num" CellPadding="4" ForeColor="#333333" GridLines="None">
            <Columns>
                <asp:BoundField DataField="object_name" HeaderText="關卡名稱" SortExpression="object_name" >
                    <HeaderStyle HorizontalAlign="Center" Width="30%" />
                </asp:BoundField>
                <asp:BoundField DataField="ORG_NAME" HeaderText="部門名稱" SortExpression="ORG_NAME">
                    <HeaderStyle HorizontalAlign="Center" Width="30%" />
                </asp:BoundField>
                <asp:BoundField DataField="emp_chinese_name" HeaderText="關卡人員" SortExpression="emp_chinese_name" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:BoundField DataField="employee_id" HeaderText="帳號" SortExpression="employee_id">
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:ImageButton ID="ImageButton1" runat="server" CommandName="Delete" ImageUrl="~/Image/delete.gif"
                            OnClientClick="return confirm('確定刪除嗎?')" SkinID="刪除" ToolTip="刪除" />
                    </ItemTemplate>
                    <HeaderStyle Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
            </Columns>
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT SYSTEMOBJUSE.object_num, SYSTEMOBJ.object_name, EMPLOYEE.emp_chinese_name, EMPLOYEE.employee_id, ADMINGROUP.ORG_NAME FROM SYSTEMOBJ INNER JOIN SYSTEMOBJUSE ON SYSTEMOBJ.object_uid = SYSTEMOBJUSE.object_uid INNER JOIN EMPLOYEE ON SYSTEMOBJUSE.employee_id = EMPLOYEE.employee_id INNER JOIN ADMINGROUP ON EMPLOYEE.ORG_UID = ADMINGROUP.ORG_UID ORDER BY SYSTEMOBJ.object_name,ADMINGROUP.ORG_NAME" DeleteCommand="DELETE FROM SYSTEMOBJUSE WHERE (object_num = @object_num)">
            <DeleteParameters>
                <asp:Parameter Name="object_num" />
            </DeleteParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>" SelectCommand="SELECT [object_uid], [object_name], [object_type], [display_order] FROM [SYSTEMOBJ] WHERE ([object_type] <> @object_type)">
            <SelectParameters>
                <asp:Parameter DefaultValue="2" Name="object_type" Type="Int32" />
            </SelectParameters>
        </asp:SqlDataSource>
    </form>
</body>
</html>
