<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00109.aspx.vb" Inherits="Source_00_MOA00109" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>系統功能設定</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />  
</head>
<body>
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="系統功能設定" Width="100%"></asp:Label>
            </td></tr>
        </table>
        <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%">
            <tr>
                <td class="form" nowrap="nowrap" style="width: 10%; height: 25px">
                    <asp:Label ID="Label1" runat="server" CssClass="form" Text="功能代碼:" Width="77px"></asp:Label></td>
                <td class="form" nowrap="nowrap" style="width: 20%; height: 25px">
                    <asp:TextBox ID="CVar" runat="server" Width="102px"></asp:TextBox></td>
                <td class="form" nowrap="nowrap" style="width: 10%; height: 25px">
                    <asp:Label ID="Label2" runat="server" CssClass="form" Text="功能名稱:" Width="77px"></asp:Label></td>
                <td class="form" nowrap="nowrap" style="width: 20%; height: 25px">
                    <asp:TextBox ID="CDesc" runat="server" Width="102px"></asp:TextBox></td>
                <td class="form" nowrap="nowrap" style="width: 10%; height: 25px">
                    <asp:Label ID="Label3" runat="server" CssClass="form" Text="功能內容:" Width="77px"></asp:Label></td>
                <td class="form" nowrap="nowrap" style="width: 20%; height: 25px">
                    <asp:TextBox ID="CVal" runat="server" Width="102px"></asp:TextBox></td>
                <td class="form" nowrap="nowrap" style="width: 10%; height: 25px">
                    <asp:ImageButton ID="ImgBtn1" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" /></td>
            </tr>
        </table>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
            AutoGenerateColumns="False" CellPadding="4" DataKeyNames="CONFIG_VAR" DataSourceID="SqlDataSource1"
            ForeColor="#333333" GridLines="None" Width="100%">
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <Columns>
                <asp:BoundField DataField="CONFIG_VAR" HeaderText="系統代碼" ReadOnly="True" SortExpression="CONFIG_VAR">
                    <HeaderStyle HorizontalAlign="Center" Width="30%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="CONFIG_DESC" HeaderText="系統名稱" SortExpression="CONFIG_DESC">
                    <HeaderStyle HorizontalAlign="Center" Width="30%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="CONFIG_VALUE" HeaderText="系統內容" SortExpression="CONFIG_VALUE">
                    <HeaderStyle HorizontalAlign="Center" Width="30%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <EditItemTemplate>
                        <asp:ImageButton ID="ImageButton1" runat="server" CausesValidation="True" CommandName="Update"
                            ImageUrl="~/Image/apply.gif" ToolTip="確認" />&nbsp;<asp:ImageButton ID="ImageButton2"
                                runat="server" CausesValidation="False" CommandName="Cancel" ImageUrl="~/Image/cancel.gif" ToolTip="取消" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:ImageButton ID="ImageButton1" runat="server" CausesValidation="False" CommandName="Edit"
                            ImageUrl="~/Image/update.gif" ToolTip="修改" />&nbsp;<asp:ImageButton ID="ImageButton2"
                                runat="server" CausesValidation="False" CommandName="Delete" ImageUrl="~/Image/delete.gif" OnClientClick="return confirm('確定刪除嗎?')" ToolTip="刪除" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
            </Columns>
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [SYSCONFIG] WHERE [CONFIG_VAR] = @CONFIG_VAR" InsertCommand="INSERT INTO [SYSCONFIG] ([CONFIG_VAR], [CONFIG_DESC], [CONFIG_VALUE]) VALUES (@CONFIG_VAR, @CONFIG_DESC, @CONFIG_VALUE)"
            SelectCommand="SELECT [CONFIG_VAR], [CONFIG_DESC], [CONFIG_VALUE] FROM [SYSCONFIG] ORDER BY [CONFIG_NUM]"
            UpdateCommand="UPDATE [SYSCONFIG] SET [CONFIG_DESC] = @CONFIG_DESC, [CONFIG_VALUE] = @CONFIG_VALUE WHERE [CONFIG_VAR] = @CONFIG_VAR">
            <DeleteParameters>
                <asp:Parameter Name="CONFIG_VAR" Type="String" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="CONFIG_DESC" Type="String" />
                <asp:Parameter Name="CONFIG_VALUE" Type="String" />
                <asp:Parameter Name="CONFIG_VAR" Type="String" />
            </UpdateParameters>
            <InsertParameters>
                <asp:Parameter Name="CONFIG_VAR" Type="String" />
                <asp:Parameter Name="CONFIG_DESC" Type="String" />
                <asp:Parameter Name="CONFIG_VALUE" Type="String" />
            </InsertParameters>
        </asp:SqlDataSource>
    </form>
</body>
</html>
