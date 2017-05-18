<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04017.aspx.vb" Inherits="Source_04_MOA04017" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>房間(區域)編號管理</title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
    <script type="text/javascript" language="javascript" src="../Inc/inc_common.js"></script>
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
    <div>
        <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center">
                    <asp:Label ID="Label1" runat="server" CssClass="toptitle" Text="房間(區域)編號管理" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>
        <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
            <tr>
                <td nowrap class="form" style="height: 25px; width: 5%;">
                    <asp:Label ID="Label2" runat="server" CssClass="form" Text="編號:" Width="77px"></asp:Label>
                </td>
                <td nowrap width="15%" class="form" style="height: 25px">
                    <asp:TextBox ID="tb_rnum_code" runat="server" MaxLength="5" Width="90%"></asp:TextBox>
                </td>
                <td nowrap class="form" style="height: 25px; width: 6%;">
                    <asp:Label ID="Label3" runat="server" CssClass="form" Text="名稱:" Width="77px"></asp:Label>
                </td>
                <td nowrap width="15%" class="form" style="height: 25px">
                    <asp:TextBox ID="tb_rnum_name" runat="server" MaxLength="100" Width="90%"></asp:TextBox>
                </td>
                <td nowrap class="form" style="height: 25px; width: 5%;">
                    <asp:Label ID="Label4" runat="server" CssClass="form" Text="坪數:" ></asp:Label>
                </td>
                <td nowrap width="10%" class="form" style="height: 25px">
                    <asp:TextBox ID="tb_rnum_spec" runat="server" MaxLength="29" Width="50%"></asp:TextBox>
                    </td>
                <td nowrap width="1%" class="form" style="padding-right: 5px;">&nbsp;</td>
                <td nowrap width="1%" class="form">&nbsp;</td>
            </tr>
            <tr>
                <td nowrap class="form" style="height: 25px; width: 5%;">
                    <asp:Label ID="Label5" runat="server" CssClass="form" Text="單位編號:" Width="77px"></asp:Label>
                </td>
                <td nowrap width="15%" class="form" style="height: 25px">
                    <asp:TextBox ID="tb_unit_code" runat="server" MaxLength="4" Width="90%"></asp:TextBox>
                </td>
                <td nowrap class="form" style="height: 25px; width: 6%;">
                    <asp:Label ID="Label6" runat="server" CssClass="form" Text="圖資編號:" Width="77px"></asp:Label>
                </td>
                <td nowrap width="15%" class="form" style="height: 25px" colspan="4" >
                    <asp:DropDownList ID="ddl_map_code" runat="server" DataSourceID="sds_map_code" DataTextField="map_name" DataValueField="map_code"></asp:DropDownList>
                    <asp:SqlDataSource ID="sds_map_code" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>" SelectCommand="Select map_name, map_code From P_0409 Order by map_name, map_code"></asp:SqlDataSource>
                    <asp:Label ID="lbl_map_code" runat="server" Visible="false"></asp:Label>
                </td>
                <td nowrap width="1%" class="form" style="padding-right: 5px;">
                    <asp:ImageButton ID="ibtnSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />
                </td>
                <td nowrap width="1%" class="form">
                    <asp:ImageButton ID="ibtnAdd" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" />
                </td>
            </tr>
        </table>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True" EmptyDataText="查無任何資料" AutoGenerateColumns="False" DataSourceID="SqlDataSource1" Width="100%" DataKeyNames="rnum_code" CellPadding="4" ForeColor="#333333" GridLines="None">
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <Columns>
                <asp:BoundField DataField="rnum_code" SortExpression="rnum_code" HeaderText="編號" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="名稱" SortExpression="rnum_name">
                    <HeaderStyle HorizontalAlign="Center" Width="30%" />
                    <ItemStyle HorizontalAlign="Left" />
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox1" runat="server" MaxLength="100" Text='<%# Bind("rnum_name") %>' Width = "180px" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("rnum_name") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="坪數" SortExpression="rnum_spec">
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox2" runat="server" Text='<%# Bind("rnum_spec") %>' Width = "30px" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# Bind("rnum_spec") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="單位編號" SortExpression="unit_code">
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                    <ItemStyle HorizontalAlign="Center" />
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox3" runat="server" MaxLength="4" Text='<%# Bind("unit_code") %>' Width = "50px"  />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label3" runat="server" Text='<%# Bind("unit_code") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="圖資編號" SortExpression="map_code">
                    <HeaderStyle HorizontalAlign="Center" Width="35%" />
                    <ItemStyle HorizontalAlign="Center" Wrap="false" />
                    <EditItemTemplate>
                        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="text-align: left;">
                            <tr>
                                <td style="width: 100%; text-align: center;">
                                    <asp:DropDownList ID="gv_ddl_map_code" runat="server" Width="200px" DataSourceID="sds_map_code" DataTextField="map_name" DataValueField="map_code"></asp:DropDownList>
                                </td>
                                <td style="width: 0%;">
                                    <asp:Label ID="gv_lbl_map_code" runat="server" Visible="false"></asp:Label>
                                </td>
                            </tr>
                        </table>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label4" runat="server" Text='<%# Bind("map_code") %>'></asp:Label>
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
                        <asp:ImageButton ID="btnImgDel" runat="server" CommandName="Delete" ImageUrl="~/Image/delete.gif" OnClientClick="return confirm('確定刪除這筆房間(區域)編號資料嗎?')" ToolTip="刪除" />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server"

            ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="Select rnum_code, rnum_name, rnum_spec, unit_code, map_code From P_0411 Order by rnum_code"
            DeleteCommand="Delete From P_0411 Where rnum_code = @rnum_code"
            UpdateCommand="Update P_0411 Set rnum_name = @rnum_name, rnum_spec = @rnum_spec, unit_code = @unit_code, map_code = @map_code Where (rnum_code = @rnum_code)">

            <DeleteParameters>
                <asp:Parameter Name="rnum_code" />
            </DeleteParameters>

            <UpdateParameters>
                <asp:Parameter Name="rnum_name" />
                <asp:Parameter Name="rnum_spec" />
                <asp:Parameter Name="unit_code" />
                <asp:Parameter Name="map_code" />
                <asp:Parameter Name="rnum_code" />
            </UpdateParameters>
        </asp:SqlDataSource>
    </div>
    </form>
</body>
</html>