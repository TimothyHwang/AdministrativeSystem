<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04012.aspx.vb" Inherits="Source_04_MOA04012" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>樓層編號管理</title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
    <script type="text/javascript" language="javascript" src="../Inc/inc_common.js"></script>
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
    <div>
        <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center">
                    <asp:Label ID="Label1" runat="server" CssClass="toptitle" Text="樓層編號管理" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>
        <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
            <tr>
                <td nowrap class="form" style="height: 25px; width: 10%;">
                    <asp:Label ID="Label2" runat="server" CssClass="form" Text="樓層編號:" Width="77px"></asp:Label>
                </td>
                <td nowrap width="15%" class="form" style="height: 25px">
                    <asp:TextBox ID="tb_fl_code" runat="server" MaxLength="2" Width="150px"></asp:TextBox>
                </td>
                <td nowrap class="form" style="height: 25px; width: 10%;">
                    <asp:Label ID="Label3" runat="server" CssClass="form" Text="樓層名稱:" Width="77px"></asp:Label>
                </td>
                <td nowrap width="15%" class="form" style="height: 25px">
                    <asp:TextBox ID="tb_fl_name" runat="server" MaxLength="50" Width="150px"></asp:TextBox>
                </td>
                <td nowrap class="form" style="height: 25px; width: 10%;">
                    <asp:Label ID="Label4" runat="server" CssClass="form" Text="排序順序:" Width="77px"></asp:Label>
                </td>
                <td nowrap width="15%" class="form" style="height: 25px">
                    <asp:TextBox ID="tb_fl_no" runat="server" Width="150px"></asp:TextBox>
                </td>
                <td nowrap width="1%" class="form" style="padding-right: 5px;">
                    <asp:ImageButton ID="ibtnSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />
                </td>
                <td nowrap width="1%" class="form">
                    <asp:ImageButton ID="ibtnAdd" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" />
                </td>
            </tr>
        </table>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True" EmptyDataText="查無任何資料" AutoGenerateColumns="False" DataSourceID="SqlDataSource1" Width="100%" DataKeyNames="fl_code" CellPadding="4" ForeColor="#333333" GridLines="None">
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <Columns>
                <asp:BoundField DataField="fl_code" SortExpression="fl_code" HeaderText="樓層編號" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" Width="25%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="樓層名稱" SortExpression="fl_name">
                    <HeaderStyle HorizontalAlign="Center" Width="25%" />
                    <ItemStyle HorizontalAlign="Left" />
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox2" runat="server" MaxLength="50" Text='<%# Bind("fl_name") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# Bind("fl_name") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="排序順序" SortExpression="fl_no">
                    <HeaderStyle HorizontalAlign="Center" Width="25%" />
                    <ItemStyle HorizontalAlign="Center" />
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox3" runat="server" Text='<%# Bind("fl_no") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label3" runat="server" Text='<%# Bind("fl_no") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField ShowHeader="False">
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                    <ItemStyle HorizontalAlign="Center" Wrap="false" />
                    <EditItemTemplate>
                        <asp:ImageButton ID="btnImgOK" runat="server" CommandName="Update" ImageUrl="~/Image/apply.gif" ToolTip="確認" />
                        <asp:ImageButton ID="btnImgCancel" runat="server" CommandName="Cancel" ImageUrl="~/Image/cancel.gif" ToolTip="取消" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:ImageButton ID="btnImgEdit" runat="server" CommandName="Edit" ImageUrl="~/Image/update.gif" ToolTip="修改" />
                        <asp:ImageButton ID="btnImgDel" runat="server" CommandName="Delete" ImageUrl="~/Image/delete.gif" OnClientClick="return confirm('確定刪除樓層資料嗎?')" ToolTip="刪除" />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server"

            ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="Select fl_code, fl_name, fl_no From P_0406 Order by fl_no"
            DeleteCommand="Delete From P_0406 Where fl_code = @fl_code"
            UpdateCommand="Update P_0406 Set fl_name = @fl_name, fl_no = @fl_no Where (fl_code = @fl_code)">

            <DeleteParameters>
                <asp:Parameter Name="fl_code" />
            </DeleteParameters>

            <UpdateParameters>
                <asp:Parameter Name="fl_name" />
                <asp:Parameter Name="fl_code" />
                <asp:Parameter Name="fl_no" />
            </UpdateParameters>
        </asp:SqlDataSource>
    </div>
    </form>
</body>
</html>