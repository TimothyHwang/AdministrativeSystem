<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00054.aspx.vb" Inherits="Source_00_MOA00054" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>群組人員管理</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">    
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label6" runat="server" CssClass="toptitle" Text="群組人員管理" Width="100%"></asp:Label>
            </td></tr>
        </table>       
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap class="form" style="height: 25px; width: 10%;">
                    <asp:Label ID="Label1" runat="server" CssClass="form" Text="群組名稱:" Width="77px"></asp:Label></td>
			    <td noWrap width="30%" class="form" style="height: 25px">			    
                    <asp:DropDownList ID="DropDownList1" runat="server" DataSourceID="SqlDataSource2" DataTextField="Group_Name" DataValueField="Group_Uid">
                    </asp:DropDownList></td>
			    <td noWrap class="form" style="height: 25px; width: 10%;">
                    <asp:Label ID="Label2" runat="server" CssClass="form" Text="人員名稱:" Width="77px"></asp:Label></td>
                    <td noWrap width="30%" class="form" style="height: 25px">
                        <asp:TextBox ID="PerName" runat="server" Width="150px"></asp:TextBox></td>
                <td noWrap width="10%" class="form">
                    <asp:ImageButton ID="searchbtn" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />
                </td>
                <td noWrap width="10%" class="form">
			        <asp:ImageButton ID="ImgBtn1" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" />
                </td>
		    </tr>
	    </table>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True" EmptyDataText="查無任何資料"
            AutoGenerateColumns="False" DataSourceID="SqlDataSource1" Width="100%" DataKeyNames="Role_Num" CellPadding="4" ForeColor="#333333" GridLines="None">
            <Columns>
                <asp:TemplateField HeaderText="群組名稱" SortExpression="Group_Name">
                    <EditItemTemplate>
                        <asp:DropDownList ID="DropDownList2" runat="server" DataSourceID="SqlDataSource2"
                            DataTextField="Group_Name" DataValueField="Group_Uid" SelectedValue='<%# Bind("Group_Uid") %>'>
                        </asp:DropDownList>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("Group_Name") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="25%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:BoundField DataField="ORG_NAME" SortExpression="ORG_NAME" HeaderText="部門名稱" >
                    <HeaderStyle HorizontalAlign="Center" Width="25%" />
                </asp:BoundField>
                <asp:BoundField DataField="emp_chinese_name" HeaderText="人員" SortExpression="emp_chinese_name" ReadOnly="True" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                </asp:BoundField>
                <asp:BoundField DataField="employee_id" HeaderText="帳號" ReadOnly="True" SortExpression="employee_id">
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <EditItemTemplate>
                        <asp:ImageButton ID="btnImgOK" runat="server" CommandName="Update" ImageUrl="~/Image/apply.gif"
                            ToolTip="確認" />
                        <asp:ImageButton ID="btnImgCancel" runat="server" CommandName="Cancel" ImageUrl="~/Image/cancel.gif"
                            ToolTip="取消" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:ImageButton ID="btnImgEdit" runat="server" CommandName="Edit" ImageUrl="~/Image/update.gif"
                            ToolTip="修改" />
                        <asp:ImageButton ID="btnImgDel" runat="server" CommandName="Delete" ImageUrl="~/Image/delete.gif"
                            OnClientClick="return confirm('確定刪除群組人員嗎?')" ToolTip="刪除" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
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
            SelectCommand="SELECT ROLEGROUP.Group_Name, EMPLOYEE.emp_chinese_name, ROLEGROUPITEM.Role_Num, ROLEGROUPITEM.Group_Uid, ROLEGROUPITEM.employee_id,ADMINGROUP.ORG_NAME FROM ROLEGROUPITEM INNER JOIN ROLEGROUP ON ROLEGROUPITEM.Group_Uid = ROLEGROUP.Group_Uid INNER JOIN EMPLOYEE ON ROLEGROUPITEM.employee_id = EMPLOYEE.employee_id INNER Join ADMINGROUP ON EMPLOYEE.ORG_UID = ADMINGROUP.ORG_UID &#13;&#10;ORDER BY  ROLEGROUP.Group_Name,ADMINGROUP.ORG_NAME" DeleteCommand="DELETE FROM ROLEGROUPITEM WHERE Role_Num= @Role_Num" UpdateCommand="UPDATE ROLEGROUPITEM SET Group_Uid=@Group_Uid WHERE (Role_Num=@Role_Num)">
            <DeleteParameters>
                <asp:Parameter Name="Role_Num" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="Group_Uid" />
                <asp:Parameter Name="Role_Num" />
            </UpdateParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [ROLEGROUPITEM] WHERE [employee_id] = @employee_id" SelectCommand="SELECT [Group_Uid], [Group_Name] FROM [ROLEGROUP] ORDER BY [Group_Order]">
        </asp:SqlDataSource>
    </form>
</body>
</html>
