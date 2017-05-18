<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00060.aspx.vb" Inherits="Source_00_MOA00060" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>關卡管理</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
    <div>
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label6" runat="server" CssClass="toptitle" Text="關卡管理" Width="100%"></asp:Label>
            </td></tr>
        </table>       
        <asp:DetailsView ID="DetailsView1" runat="server" AutoGenerateRows="False" DataKeyNames="object_uid"
            DataSourceID="SqlDataSource1" DefaultMode="Insert" Height="50px" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None" CssClass="form">
            <Fields>
                <asp:TemplateField HeaderText="object_name" ShowHeader="False" SortExpression="object_name">
                    <InsertItemTemplate>
                        <table width="100%">
                        <tr>
			            <td noWrap width="10%" class="form">
                            <asp:Label ID="Lab1" runat="server" Text="關卡名稱："></asp:Label>
			            </td>
			            <td noWrap width="20%" class="form">
                            <asp:TextBox ID="object_name" runat="server" Text='<%# Bind("object_name") %>' Width="100px"></asp:TextBox>
			            </td>
			            <td noWrap width="10%" class="form">
                            <asp:Label ID="Label5" runat="server" Text="關卡型態："></asp:Label>
			            </td>
			            <td noWrap width="20%" class="form">
                            <asp:DropDownList ID="object_type" runat="server" SelectedValue='<%# Bind("object_type") %>'>
                                <asp:ListItem Value="1">全關卡</asp:ListItem>
                                <asp:ListItem Value="2">上一級單位</asp:ListItem>
                                <asp:ListItem Value="3">同單位</asp:ListItem>
                                <asp:ListItem Value="4">指定群組</asp:ListItem>
                            </asp:DropDownList>
			            </td>			            
			            <td noWrap width="10%" class="form">
                            <asp:Label ID="Lab2" runat="server" Text="排序："></asp:Label>
			            </td>
			            <td noWrap width="10%" class="form">
                            <asp:TextBox ID="display_order" runat="server" Text='<%# Bind("display_order") %>' Width="100px"></asp:TextBox>			            
			            </td>
			            <td noWrap class="form" style="width: 20%">
                            <asp:ImageButton ID="btnImgIns" runat="server" ImageUrl="~/Image/add.gif" OnClick="btnImgIns_Click"
                                ToolTip="新增" />
                            <asp:TextBox ID="object_uid" runat="server" Text='<%# Bind("object_uid") %>' Visible="False" Width="42px"></asp:TextBox>
                            <asp:Label ID="InsErr" runat="server" ForeColor="Red" Width="102px"></asp:Label>
			            </td>
                        </tr>
                        </table>                        
                    </InsertItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("object_name") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
            </Fields>
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <CommandRowStyle BackColor="#E2DED6" Font-Bold="True" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <FieldHeaderStyle BackColor="#E9ECF1" Font-Bold="True" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#E9ECF1" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
        </asp:DetailsView>
    
    </div>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
            AutoGenerateColumns="False" DataKeyNames="object_uid" DataSourceID="SqlDataSource1" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
            <Columns>
                <asp:TemplateField HeaderText="關卡代號" SortExpression="object_uid" Visible="False">
                    <EditItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Eval("object_uid") %>'></asp:Label>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("object_uid") %>'></asp:Label>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="關卡名稱" SortExpression="object_name">
                    <EditItemTemplate>
                        <asp:TextBox ID="GP_NAME" runat="server" Text='<%# Bind("object_name") %>'></asp:TextBox>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="GP_NAME"
                            ErrorMessage="請輸入群組名稱"></asp:RequiredFieldValidator>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# Bind("object_name") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle Width="40%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="關卡型態" SortExpression="object_type">
                    <EditItemTemplate>
                            <asp:DropDownList ID="object_type" runat="server" SelectedValue='<%# Bind("object_type") %>'>
                                <asp:ListItem Value="1">全關卡</asp:ListItem>
                                <asp:ListItem Value="2">上一級單位</asp:ListItem>
                                <asp:ListItem Value="3">同單位</asp:ListItem>
                                <asp:ListItem Value="4">指定群組</asp:ListItem>
                            </asp:DropDownList>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label3" runat="server" Text='<%# FunType("object_type") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle Width="30%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="排序" SortExpression="display_order">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox3" runat="server" Text='<%# Bind("display_order") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label4" runat="server" Text='<%# Bind("display_order") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField ShowHeader="False">
                    <EditItemTemplate>
                        <asp:ImageButton ID="btnImgOK" runat="server" CommandName="Update" ImageUrl="~/Image/apply.gif"
                            ToolTip="確認" />
                        <asp:ImageButton ID="btnImgCancel" runat="server" CommandName="Cancel" ImageUrl="~/Image/cancel.gif"
                            ToolTip="取消" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:ImageButton ID="btnImgUpd" runat="server" CommandName="Edit" ImageUrl="~/Image/update.gif"
                            ToolTip="修改" />
                        <asp:ImageButton ID="btnImgDel" runat="server" CommandName="Delete" ImageUrl="~/Image/delete.gif"
                            OnClientClick="return confirm('確定刪除嗎?')" ToolTip="刪除" />
                    </ItemTemplate>
                    <HeaderStyle Width="20%" />
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
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [SYSTEMOBJ] WHERE [object_uid] = @object_uid" InsertCommand="INSERT INTO [SYSTEMOBJ] ([object_uid], [object_name], [object_type], [display_order]) VALUES (@object_uid, @object_name, @object_type, @display_order)"
            SelectCommand="SELECT [object_uid], [object_name], [object_type], [display_order] FROM [SYSTEMOBJ] ORDER BY [display_order]"
            UpdateCommand="UPDATE [SYSTEMOBJ] SET [object_name] = @object_name,[object_type] = @object_type, [display_order] = @display_order WHERE [object_uid] = @object_uid">
            <DeleteParameters>
                <asp:Parameter Name="object_uid" Type="String" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="object_name" Type="String" />
                <asp:Parameter Name="object_type" Type="Int32" />
                <asp:Parameter Name="display_order" Type="Int32" />
                <asp:Parameter Name="object_uid" Type="String" />
            </UpdateParameters>
            <InsertParameters>
                <asp:Parameter Name="object_uid" Type="String" />
                <asp:Parameter Name="object_name" Type="String" />
                <asp:Parameter Name="object_type" Type="Int32" />
                <asp:Parameter Name="display_order" Type="Int32" />
            </InsertParameters>
        </asp:SqlDataSource>
    </form>
</body>
</html>
