<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04016.aspx.vb" Inherits="Source_04_MOA04016" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>圖資編號管理</title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
    <script type="text/javascript" language="javascript" src="../Inc/inc_common.js"></script>
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
    <div>
        <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center">
                    <asp:Label ID="Label1" runat="server" CssClass="toptitle" Text="圖資編號管理" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>
        <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
            <tr>
                <td nowrap class="form" style="height: 25px; width: 10%;">
                    <asp:Label ID="Label2" runat="server" CssClass="form" Text="圖資編號:" Width="77px"></asp:Label>
                </td>
                <td nowrap width="15%" class="form" style="height: 25px">
                    <asp:TextBox ID="tb_map_code" runat="server" MaxLength="20" Width="150px"></asp:TextBox>
                </td>
                <td nowrap class="form" style="height: 25px; width: 10%;">
                    <asp:Label ID="Label3" runat="server" CssClass="form" Text="圖資名稱:" Width="77px"></asp:Label>
                </td>
                <td nowrap width="15%" class="form" style="height: 25px">
                    <asp:TextBox ID="tb_map_name" runat="server" MaxLength="50" Width="150px"></asp:TextBox>
                </td>
                <td nowrap width="1%" class="form" style="padding-right: 5px;">
                    <asp:ImageButton ID="ibtnSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />
                </td>
                <td nowrap width="1%" class="form">
                    <asp:ImageButton ID="ibtnAdd" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" />
                </td>
            </tr>
        </table>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True" EmptyDataText="查無任何資料" AutoGenerateColumns="False" DataSourceID="SqlDataSource1" Width="100%" DataKeyNames="map_code" CellPadding="4" ForeColor="#333333" GridLines="None">
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <Columns>
                <asp:BoundField DataField="map_code" SortExpression="map_code" HeaderText="圖資編號" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="圖資名稱" SortExpression="map_name">
                    <HeaderStyle HorizontalAlign="Center" Width="30%" />
                    <ItemStyle HorizontalAlign="Left" />
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox2" runat="server" MaxLength="50" Text='<%# Bind("map_name") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# Bind("map_name") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="map_no" SortExpression="map_no" HeaderText="圖資流水號" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" Width="25%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                    <ItemStyle HorizontalAlign="Center" Wrap="false" />
                    <EditItemTemplate>
                        <asp:ImageButton ID="btnImgOK" runat="server" CommandName="Update" ImageUrl="~/Image/apply.gif" ToolTip="確認" />
                        <asp:ImageButton ID="btnImgCancel" runat="server" CommandName="Cancel" ImageUrl="~/Image/cancel.gif" ToolTip="取消" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:ImageButton ID="btnImgEdit" runat="server" CommandName="Edit" ImageUrl="~/Image/update.gif" ToolTip="修改" />
                        <asp:ImageButton ID="btnImgDel" runat="server" CommandName="Delete" ImageUrl="~/Image/delete.gif" OnClientClick="return confirm('確定刪除此圖資資料嗎?')" ToolTip="刪除" />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server"

            ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="Select map_code, map_name, map_no From P_0409 Order by map_no"
            DeleteCommand="Delete From P_0409 Where map_code = @map_code"
            UpdateCommand="Update P_0409 Set map_name = @map_name Where (map_code = @map_code)">

            <DeleteParameters>
                <asp:Parameter Name="map_code" />
            </DeleteParameters>

            <UpdateParameters>
                <asp:Parameter Name="map_name" />
                <asp:Parameter Name="map_code" />
                <asp:Parameter Name="map_no" />
            </UpdateParameters>
        </asp:SqlDataSource>
    </div>
    </form>
</body>
</html>