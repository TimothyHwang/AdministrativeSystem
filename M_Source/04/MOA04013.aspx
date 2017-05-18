<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04013.aspx.vb" Inherits="Source_04_MOA04013" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>設備物料分類管理</title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
    <script type="text/JavaScript">
        function Check() {
            var d = document.form1;
            var re = /</;
            if (re.test(d.tb_it_code.value)) {
                alert("您輸入的物料編號有包含危險字元！");
                d.tb_it_code.value = "";
                return false;
            }

            if (re.test(d.tb_it_name.value)) {
                alert("您輸入的物料名稱有包含危險字元！");
                d.tb_it_name.value = "";
                return false;
            }
            return true;
        }
</script>
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
    <div>
        <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center">
                    <asp:Label ID="Label1" runat="server" CssClass="toptitle" Text="設備物料分類管理" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>
        <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
            <tr>
                <td nowrap class="form" style="height: 25px; width: 10%;">
                    <asp:Label ID="Label2" runat="server" CssClass="form" Text="物料編號:" Width="77px"></asp:Label>
                    <asp:TextBox ID="tb_it_code" runat="server" MaxLength="6" Width="150px"></asp:TextBox>
                </td>
                <td nowrap class="form" style="height: 25px; width: 10%;">
                    <asp:Label ID="Label3" runat="server" CssClass="form" Text="物料名稱:" Width="77px"></asp:Label>
                    <asp:TextBox ID="tb_it_name" runat="server" MaxLength="255" Width="150px"></asp:TextBox>
                </td>
                <td nowrap width="1%" class="form" style="padding-right: 5px;">
                    <asp:ImageButton ID="ibtnSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" OnClientClick = "return Check();" />
                </td>
                <td nowrap width="1%" class="form">
                    <asp:ImageButton ID="ibtnAdd" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" onclientclick="form1.reset();" />
                </td>
            </tr>
        </table>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True" EmptyDataText="查無任何資料" AutoGenerateColumns="False" DataSourceID="SqlDataSource1" Width="100%" DataKeyNames="it_code" CellPadding="4" ForeColor="#333333" GridLines="None">
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <Columns>
                <asp:BoundField DataField="it_code" SortExpression="it_code" HeaderText="編號" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" Wrap="false" />
                </asp:BoundField>
                <asp:BoundField DataField="it_name" SortExpression="it_name" HeaderText="名稱" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Left" Wrap="false" />
                </asp:BoundField>
                <asp:BoundField DataField="it_spec" SortExpression="it_spec" HeaderText="規格" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Left" Wrap="false" />
                </asp:BoundField>
                <asp:BoundField DataField="it_cost" SortExpression="it_cost" HeaderText="價格" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Left" />
                    <ItemStyle HorizontalAlign="Left" Wrap="false" />
                </asp:BoundField>
                <asp:BoundField DataField="manufacturer" SortExpression="manufacturer" HeaderText="公司" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Left" />
                    <ItemStyle HorizontalAlign="Left" Wrap="false" />
                </asp:BoundField>
                <asp:BoundField DataField="memo" SortExpression="memo" HeaderText="備註" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Left" />
                    <ItemStyle HorizontalAlign="Left" Wrap="false" />
                </asp:BoundField>
                <asp:TemplateField  HeaderText="圖檔一">
                    <HeaderStyle HorizontalAlign="Left" Width="10%" />
                    <ItemStyle HorizontalAlign="Left" Wrap="false" />
                    <ItemTemplate>
                        <A HREF="#" onClick="window.open('<%# sPicPath + Eval("it_code") + "A_" + Eval("file_a") %>', '物料圖檔一', config='height=200,width=200,menubar=no,toolbar=no,location=no,status=no,resizable=yes')"> <asp:Label ID="lb1" runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "file_a") %>' /></A> 
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="圖檔二">
                    <HeaderStyle HorizontalAlign="Left" Width="10%" />
                    <ItemStyle HorizontalAlign="Left" Wrap="false" />
                    <ItemTemplate>
                        <A HREF="#" onClick="window.open('<%# sPicPath + Eval("it_code") + "B_" + Eval("file_b") %>', '物料圖檔二', config='height=200,width=200,menubar=no,toolbar=no,location=no,status=no,resizable=yes')"> <asp:Label ID="lb2" runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "file_b") %>' /></A> 
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="snum" SortExpression="snum" HeaderText="安全數量" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Left" />
                    <ItemStyle HorizontalAlign="Left" Wrap="false" />
                </asp:BoundField>
                <asp:BoundField DataField="expired_y" SortExpression="expired_y" HeaderText="有效期" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Left" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="it_sort" SortExpression="it_sort" HeaderText="類別" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="it_unit" SortExpression="it_unit" HeaderText="單位" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Left" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="operator" SortExpression="operator" HeaderText="建置人員" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Left" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <HeaderStyle HorizontalAlign="Center" Width="1%" />
                    <ItemStyle HorizontalAlign="Center" Wrap="false" />
                    <EditItemTemplate>
                        <asp:ImageButton ID="btnImgOK" runat="server" CommandName="Update" ImageUrl="~/Image/apply.gif" ToolTip="確認" />
                        <asp:ImageButton ID="btnImgCancel" runat="server" CommandName="Cancel" ImageUrl="~/Image/cancel.gif" ToolTip="取消" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:ImageButton ID="btnImgEdit" runat="server" CommandName="Edit" ImageUrl="~/Image/update.gif" ToolTip="修改" />
                        <asp:ImageButton ID="btnImgDel" runat="server" CommandName="Delete" ImageUrl="~/Image/delete.gif" OnClientClick="return confirm('確定刪除該筆物料資料嗎?')" ToolTip="刪除" />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server"
            ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="Select it_code, it_name, it_spec, it_cost, manufacturer, memo, file_a, file_b, snum, expired_y, it_sort, it_unit, operator From P_0407 Order by it_code"
            DeleteCommand="Delete From P_0407 Where it_code = @it_code">
            <DeleteParameters>
                <asp:Parameter Name="it_code" />
            </DeleteParameters>
        </asp:SqlDataSource>
    </div>
    </form>
</body>
</html>