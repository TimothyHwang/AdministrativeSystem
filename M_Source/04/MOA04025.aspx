<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04025.aspx.vb" Inherits="M_Source_04_MOA04025" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>廢料管理查詢</title>
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

             if (re.test(d.tb_itname.value)) {
                 alert("您輸入的品項名稱有包含危險字元！");
                 d.tb_itname.value = "";
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
                    <asp:Label ID="lblTitle" runat="server" CssClass="toptitle" Text="廢料管理查詢" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>
        <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
            <tr>
                <td nowrap class="form" style="height: 25px; width: 10%;">
                    <asp:Label ID="Label1" runat="server" CssClass="form" Text="物料編號:" Width="77px"></asp:Label>
                    <asp:TextBox ID="tb_it_code" runat="server" MaxLength="6" Width="100px"></asp:TextBox>
                </td>
                <td nowrap class="form" style="height: 25px; width: 10%;">
                    <asp:Label ID="Label2" runat="server" CssClass="form" Text="品項名稱:" Width="77px"></asp:Label>
                    <asp:TextBox ID="tb_itname" runat="server" MaxLength="255" Width="150px"></asp:TextBox>
                </td>
                 <td nowrap class="form" style="height: 25px; width: 10%;">
                    <asp:Label ID="Label3" runat="server" CssClass="form" Text="狀態:" Width="40px"></asp:Label>
                    <asp:DropDownList ID="ddlUseCheck" runat="server">
                        <asp:ListItem Value="-1" Selected = "True" >全部</asp:ListItem>
                        <asp:ListItem Value="0">庫存</asp:ListItem>
                        <asp:ListItem Value="2">出庫</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td nowrap width="1%" class="form" style="padding-right: 5px;">
                    <asp:ImageButton ID="ibtnSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" OnClientClick ="return Check();" />
                </td>
                <td nowrap width="1%" class="form">
                    <asp:Button ID="ibtnAdd" runat="server" Text="入庫" />
                </td>
            </tr>
        </table>
        <asp:GridView ID="gvDataRecords" runat="server" AllowPaging="True" AllowSorting="True" AutoGenerateColumns="False" 
        CellPadding="4" DataKeyNames="it_code" DataSourceID="sdsDataRecords" EmptyDataText="查無任何資料" ForeColor="#333333" GridLines="None" Width="100%">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EditRowStyle BackColor="#999999" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <Columns>
                <asp:TemplateField HeaderText="物料編號" SortExpression="it_code">
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" Wrap="False" />
                    <ItemTemplate>
                        <asp:Label ID="lb_it_code" runat="server" text='<%# Bind("it_code") %>'  />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="品項名稱" SortExpression="it_name">
                    <HeaderStyle HorizontalAlign="left" />
                    <ItemStyle HorizontalAlign="left" Wrap="False" />
                    <ItemTemplate>
                    <A HREF="#" onClick="window.open('<%# "MOA04025_1.aspx?itcode=" + Eval("it_code") + "&usecheck=" + Eval("UseCheck") %>', '物料品項資訊', config='height=800,width=1200,scrollbars=yes,status=yes')"> <asp:Label ID="Label3" runat="server" Text='<%# Bind("it_name") %>' /></A> 
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="狀態" SortExpression="UseCheck">
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" Wrap="False" />
                    <ItemTemplate>
                        <asp:Label ID="Label4" runat="server" Text='<%# ShowStocksStateText(DataBinder.Eval(Container.DataItem, "UseCheck")) %>' />
                        <asp:HiddenField ID="hf_UseCheck" runat="server" Value='<%# Bind("UseCheck") %>' />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="總計" SortExpression="cntUseCheck">
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" Wrap="False" />
                    <ItemTemplate>
                        <asp:Label ID="lbcntUseCheck" runat="server" text='<%# Bind("cntUseCheck") %>'  />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>

        <asp:SqlDataSource ID="sdsDataRecords" runat="server"
            ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT substring(a.shcode,1,6) as it_code,b.it_name, Usecheck, count(Usecheck) as cntUseCheck
            FROM  P_0414 a left join P_0407 b on substring(a.shcode,1,6) = b.it_code
            WHERE (1=1) and shtype = '1' and usecheck in (0,2) group by Usecheck ,b.it_name, substring(a.shcode,1,6) Order by it_code;">
        </asp:SqlDataSource>
    </div>
    </form>
</body>
</html>
