<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04021.aspx.vb" Inherits="Source_04_MOA04021" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>倉儲用料統計</title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
    <script type="text/javascript" language="javascript" src="../Inc/inc_common.js"></script>
    <style type="text/css">
        .style_records_count { }
        .style1
        {
            font: 13px Verdana, Arial, Helvetica, sans-serif;
            color: #666666;
            text-decoration: none;
            width: 1%;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center">
                    <asp:Label ID="lblTitle" runat="server" CssClass="toptitle" Text="倉儲用料統計" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>
        <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
            <tr>
                <td nowrap class="form" style="height: 25px; ">
                    <asp:Label ID="Label1" runat="server" CssClass="form" Text="設備分層編號:" Width="90px"></asp:Label>
                    <asp:TextBox ID="tb_it_code" runat="server" MaxLength="6" Width="60px"></asp:TextBox>
                </td>
                <td nowrap class="form" style="height: 25px; width: 5%;">&nbsp;
                    <asp:Label ID="Label2" runat="server" CssClass="form" Text="倉儲年份(西元):" />
                    <asp:TextBox ID="tb_year" runat="server" MaxLength="4" Width="60px"></asp:TextBox>
                </td>
                <td nowrap class="form" style="height: 25px; width: 10%;">&nbsp;
                    <asp:Label ID="Label3" runat="server" CssClass="form" Text="倉儲狀況:" Width="70px" />
                    <asp:DropDownList ID="ddl_stocks_status" runat="server">
                        <asp:ListItem Selected="True" Value="">全部</asp:ListItem>
                        <asp:ListItem Value="0">庫存</asp:ListItem>
                        <asp:ListItem Value="1">待出庫</asp:ListItem>
                        <asp:ListItem Value="2">出庫</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td width="10px">&nbsp;</td>
                <td nowrap align="left"  
                    style="padding-right: 50px; padding-top: 4px; padding-bottom: 4px;">
                    <asp:ImageButton ID="ibtnSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />
                </td>    
            </tr>
        </table>
        <asp:GridView ID="gvDataRecords" runat="server" AllowPaging="True" AllowSorting="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="sh_name" DataSourceID="sdsDataRecords" EmptyDataText="查無任何資料" ForeColor="#333333" GridLines="None" Width="100%">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EditRowStyle BackColor="#999999" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <Columns>
                <asp:BoundField DataField="sh_name" SortExpression="sh_name" HeaderText="倉儲品項" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Left" Wrap="False" Width="60%" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="狀態" SortExpression="UseCheck">
                    <HeaderStyle HorizontalAlign="Left" />
                    <ItemStyle HorizontalAlign="Left" Wrap="False" />
                    <ItemTemplate>
                        <%# ShowStocksStateText(DataBinder.Eval(Container.DataItem, "UseCheck")) %>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="Records_Count" SortExpression="Records_Count" HeaderText="總計數量" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" Wrap="False" CssClass="style_records_count" />
                </asp:BoundField>
            </Columns>
        </asp:GridView>
    </div>

    <asp:SqlDataSource ID="sdsDataRecords" runat="server"

        ConnectionString="<%$ ConnectionStrings:ConnectionString %>"

        SelectCommand="Select sh_name = [TMP].it_code + ' / ' + (Case IsNull(P07.it_spec, '') When '' Then P07.it_name Else (
                       P07.it_name + ' (' + P07.it_spec + ')') End), [TMP].UseCheck, [TMP].Records_Count From (Select Left(shcode, 6) 
                       As it_code, UseCheck, Count(*) As Records_Count From P_0414 Where shcode Like (IsNull(@shcode_partial, shcode) + '%') 
                       And UseCheck = IsNull(@stocks_status, UseCheck) Group by Left(shcode, 6), UseCheck) As [TMP] Left Join P_0407 As P07 
                       On [TMP].it_code = P07.it_code Order by [TMP].it_code, [TMP].UseCheck;">

        <SelectParameters>
            <asp:Parameter Name="shcode_partial" DbType="String" DefaultValue="%%" Size="10" />
            <asp:Parameter Name="stocks_status" DbType="String" Size="1" />
        </SelectParameters>

    </asp:SqlDataSource>
    </form>
</body>
</html>