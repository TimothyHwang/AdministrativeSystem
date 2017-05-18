<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA01007.aspx.vb" Inherits="Source_04_MOA01007" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>文職人員薪資表</title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
    <script type="text/javascript" language="javascript" src="../Inc/inc_common.js"></script>
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
    <div>
        <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center">
                    <asp:Label ID="Label1" runat="server" CssClass="toptitle" Text="文職人員薪資表" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>
        <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
            <tr>
                <td nowrap class="form" style="height: 25px; width: 10%;">
                    <asp:Label ID="Label2" runat="server" CssClass="form" Text="薪資:" Width="77px"></asp:Label>
                </td>
                <td nowrap width="15%" class="form" style="height: 25px">
                    <asp:TextBox ID="tb_p_money" runat="server" Width="50%"></asp:TextBox>
                </td>
                <td nowrap class="form" style="height: 25px; width: 10%;">
                    <asp:Label ID="Label3" runat="server" CssClass="form" Text="薪資種類:" Width="77px"></asp:Label>
                </td>
                <td nowrap width="15%" class="form" style="height: 25px">
                    <asp:DropDownList ID="ddl_p_kind" runat="server" Width="50%">
                        <asp:ListItem Selected="True"></asp:ListItem>
                        <asp:ListItem Value="1">俸額</asp:ListItem>
                        <asp:ListItem Value="2">專業加給</asp:ListItem>
                        <asp:ListItem Value="3">主管加給</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td nowrap width="1%" class="form" style="padding-right: 5px;">
                    <asp:ImageButton ID="ibtnSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />
                </td>
                <td nowrap width="1%" class="form">
                    <asp:ImageButton ID="ibtnAdd" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" />
                </td>
            </tr>
        </table>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True" EmptyDataText="查無任何資料" AutoGenerateColumns="False" DataSourceID="SqlDataSource1" Width="100%" DataKeyNames="P_Num" CellPadding="4" ForeColor="#333333" GridLines="None">
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <Columns>
                <asp:BoundField DataField="P_Num" SortExpression="P_Num" HeaderText="流水號" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" Width="25%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="薪資" SortExpression="P_Money">
                    <HeaderStyle HorizontalAlign="Center" Width="25%" />
                    <ItemStyle HorizontalAlign="Left" />
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox1" runat="server" MaxLength="50" Text='<%# Bind("P_Money") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("P_Money") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="薪資種類" SortExpression="P_KIND">
                    <HeaderStyle HorizontalAlign="Center" Width="25%" />
                    <ItemStyle HorizontalAlign="Center" />
                    <EditItemTemplate>
                        <asp:DropDownList ID="gv_ddl_p_kind" runat="server" Width="60%">
                            <asp:ListItem></asp:ListItem>
                            <asp:ListItem Value="1">俸額</asp:ListItem>
                            <asp:ListItem Value="2">專業加給</asp:ListItem>
                            <asp:ListItem Value="3">主管加給</asp:ListItem>
                        </asp:DropDownList>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# ShowMoneyKindText(DataBinder.Eval(Container, "DataItem.P_KIND", "{0}")) %>'></asp:Label>
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
                        <asp:ImageButton ID="btnImgDel" runat="server" CommandName="Delete" ImageUrl="~/Image/delete.gif" OnClientClick="return confirm('確定刪除文職人員薪資資料嗎?')" ToolTip="刪除" />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server"

            ConnectionString="<%$ ConnectionStrings:ConnectionString %>"

            SelectCommand="Select P_Num, P_Money, P_KIND From P_0102 Order by P_Num"
            DeleteCommand="Delete From P_0102 Where P_Num = @P_Num"
            UpdateCommand="Update P_0102 Set P_Money = @P_Money, P_KIND = @P_KIND Where (P_Num = @P_Num)">

            <DeleteParameters>
                <asp:Parameter Name="P_Num" />
            </DeleteParameters>

            <UpdateParameters>
                <asp:Parameter Name="P_Money" />
                <asp:Parameter Name="P_KIND" />
                <asp:Parameter Name="P_Num" />
            </UpdateParameters>
        </asp:SqlDataSource>
    </div>
    </form>
</body>
</html>