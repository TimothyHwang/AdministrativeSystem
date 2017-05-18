<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00050.aspx.vb" Inherits="Source_00_MOA00050" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>群組管理</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
    <div>
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label6" runat="server" CssClass="toptitle" Text="群組管理" Width="100%"></asp:Label>
            </td></tr>
        </table>       
            <asp:DetailsView ID="DetailsView1" runat="server" AutoGenerateRows="False" DataKeyNames="Group_Uid" DataSourceID="SqlDataSource1" DefaultMode="Insert" Height="50px" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None" CssClass="form">
                <Fields>
                    <asp:TemplateField SortExpression="Group_Name" ShowHeader="False">
                        <InsertItemTemplate>
                        <table>
                        <tr>
			            <td noWrap width="10%" class="form">
			                <asp:TextBox ID="Group_Uid" runat="server" Text='<%# Bind("Group_Uid") %>' Visible='False'></asp:TextBox>
                            <asp:Label ID="Label3" runat="server" Text="群組名稱:" CssClass="form"  ></asp:Label>
                        </td>
			            <td noWrap width="20%" class="form">
                            <asp:TextBox ID="Group_Name_Ins" runat="server" Text='<%# Bind("Group_Name") %>'></asp:TextBox>
                        </td>
			            <td noWrap width="10%" class="form">
			                <asp:Label ID="Lab2" runat="server" Text="排序:" CssClass="form" ></asp:Label>
                        </td>
			            <td noWrap width="10%" class="form">			            
                            <asp:TextBox ID="Group_Order" runat="server" Text='<%# Bind("Group_Order") %>'></asp:TextBox>
			            </td>
			            <td noWrap class="form">
                            <asp:ImageButton ID="btnInsert" runat="server" ImageUrl="~/Image/add.gif" OnClick="btnInsert_Click1"
                                ToolTip="新增" />
                            <asp:Label ID="InsErr" runat="server" ForeColor="Red" Width="102px"></asp:Label>
			            </td>
                        </tr>
                        </table>
                        </InsertItemTemplate>
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
		    <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" DataKeyNames="Group_Uid" DataSourceID="SqlDataSource1" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
                <Columns>
                  <asp:TemplateField >
                        <EditItemTemplate>
                            <asp:HiddenField ID="HF_Group_Uid" runat="server" Value = '<%# Bind("Group_Uid") %>' />
                        </EditItemTemplate>
                        <ItemTemplate>
                        </ItemTemplate>
                        <HeaderStyle Width="1%" HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:TemplateField>
                   
                    <asp:TemplateField HeaderText="群組名稱" SortExpression="Group_Name">
                        <EditItemTemplate>
                            <asp:TextBox ID="GV_GroupName" runat="server" Text='<%# Bind("Group_Name") %>'></asp:TextBox>
                            <asp:RequiredFieldValidator ID="GV_RF" runat="server" ControlToValidate="GV_GroupName"
                                ErrorMessage="GV_RF">請輸入群組名稱</asp:RequiredFieldValidator>
                        </EditItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label1" runat="server" Text='<%# Bind("Group_Name") %>'></asp:Label>
                        </ItemTemplate>
                        <HeaderStyle Width="40%" HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="排序" SortExpression="Group_Order">
                        <EditItemTemplate>
                            <asp:TextBox ID="GV_GroupOrder" runat="server" Text='<%# Bind("Group_Order") %>'></asp:TextBox>
                        </EditItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label2" runat="server" Text='<%# Bind("Group_Order") %>'></asp:Label>
                        </ItemTemplate>
                        <HeaderStyle Width="30%" HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:TemplateField>
                    <asp:TemplateField ShowHeader="False">
                        <EditItemTemplate>
                            <asp:ImageButton ID="ImageButton1" runat="server" CommandName="Update" ImageUrl="~/Image/apply.gif"
                                ToolTip="確認" />
                            <asp:ImageButton ID="ImageButton2" runat="server" CommandName="Cancel" ImageUrl="~/Image/cancel.gif"
                                ToolTip="取消" />
                        </EditItemTemplate>
                        <ItemTemplate>
                            <asp:ImageButton ID="ImgUpd" runat="server" CommandName="Edit" ImageUrl="~/Image/update.gif" ToolTip="修改" />
                            <asp:ImageButton ID="ImgDel" runat="server" CommandName="Delete" ImageUrl="~/Image/delete.gif"
                                OnClientClick="return confirm('確定刪除群組嗎?')" ToolTip="刪除" />
                        </ItemTemplate>
                        <HeaderStyle Width="30%" HorizontalAlign="Center" />
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
                InsertCommand="INSERT INTO [ROLEGROUP] ([Group_Name], [Group_Order], [Group_Uid]) VALUES (@Group_Name, @Group_Order, @Group_Uid)"
                SelectCommand="SELECT [Group_Name], [Group_Order], [Group_Uid] FROM [ROLEGROUP] ORDER BY [Group_Order]"
                UpdateCommand="UPDATE [ROLEGROUP] SET [Group_Name] = @Group_Name, [Group_Order] = @Group_Order WHERE [Group_Uid] = @Group_Uid">
                
                <UpdateParameters>
                    <asp:Parameter Name="Group_Name" Type="String" />
                    <asp:Parameter Name="Group_Order" Type="Int32" />
                    <asp:Parameter Name="Group_Uid" Type="String" />
                </UpdateParameters>
                <InsertParameters>
                    <asp:Parameter Name="Group_Name" Type="String" />
                    <asp:Parameter Name="Group_Order" Type="Int32" />
                    <asp:Parameter Name="Group_Uid" Type="String" />
                </InsertParameters>
            </asp:SqlDataSource>        &nbsp;
    
    </div>
    </form>
</body>
</html>
