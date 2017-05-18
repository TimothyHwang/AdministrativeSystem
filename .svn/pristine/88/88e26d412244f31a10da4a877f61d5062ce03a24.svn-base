<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04019_2.aspx.vb" Inherits="M_Source_04_MOA04019_2" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>物品詳細資料</title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
</head>
<body background="../../Image/main_bg.jpg">
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Label ID="lbl_it_code" runat="server" Visible="false" />
        <asp:Label ID="lbl_usecheck" runat="server" Visible="false" />
          <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center" colspan = "2">
                    <asp:Label ID="lblTitle" runat="server" CssClass="toptitle" Width="100%" />
                </td>
            </tr>
            </table>
            <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True" DataKeyNames="shcode" DataSourceID="sqldatalist" 
            EmptyDataText="查無任何資料" AutoGenerateColumns="False" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <Columns>
                  <asp:TemplateField ShowHeader="False">
                        <EditItemTemplate>
                            <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="True" CommandName="Update" Text="更新" />
                            <asp:LinkButton ID="LinkButton2" runat="server" CausesValidation="False" CommandName="Cancel" Text="取消" />
                        </EditItemTemplate>
                        <ItemTemplate>
                            <asp:ImageButton ID="ImageButton1" runat="server" CausesValidation="False" CommandName="Edit" ImageUrl ="~/Image/update.gif" Visible='<%# showeditdelLB(Eval("usecheck"))%>' />
                            <asp:ImageButton ID="ImageButton2" runat="server" CausesValidation="False" CommandName="Delete" OnClientClick="return confirm('您確定要刪除嗎?')" ImageUrl ="~/Image/delete.gif" Visible='<%# showeditdelLB(Eval("usecheck"))%>' />
                        </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="shcode" HeaderText="物品編號" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                 <asp:TemplateField HeaderText="廠商名稱" >
                     <EditItemTemplate>
                        <asp:TextBox id = "tb_company" runat="server" Text='<%# Bind("company") %>' />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label id = "lb_company" runat="server" Text='<%# Bind("company") %>' />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                 <asp:TemplateField HeaderText="物料價格" >
                     <EditItemTemplate>
                        <asp:TextBox id = "tb_shcost" runat="server" Text='<%# Bind("shcost") %>'  width = "40px" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label id = "lb_shcost" runat="server" Text='<%# Bind("shcost") %>' />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="有效期">
                     <EditItemTemplate>
                        <asp:TextBox id = "tb_expired_y" runat="server" Text='<%# Bind("expired_y") %>' width = "40px" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label id = "lb_sexpired_y" runat="server" Text='<%# Bind("expired_y") %>' />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                 <asp:TemplateField HeaderText="儲庫編號">
                     <EditItemTemplate>
                         <asp:DropDownList ID="ddl_seat_num" runat="server" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label id = "lb_Seat_Num" runat="server" Text='<%# Bind("seatnum") %>' />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="採購編號">
                     <EditItemTemplate>
                        <asp:TextBox id = "tb_Buy_Num" runat="server" Text='<%# Bind("Buy_Num") %>'  Width = "50px" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label id = "lb_Buy_Num" runat="server" Text='<%# Bind("Buy_Num") %>' Width = "50px" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                        <asp:TemplateField HeaderText="廠商產品貨號">
                     <EditItemTemplate>
                        <asp:TextBox id = "tb_company_Num" runat="server" Text='<%# Bind("company_Num") %>' Width = "40px" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label id = "lb_company_Num" runat="server" Text='<%# Bind("company_Num") %>'  Width = "40px"  />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="備註">
                     <EditItemTemplate>
                        <asp:TextBox id = "tb_memo" runat="server" Text='<%# Bind("memo") %>' Width = "40px" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label id = "lb_memo" runat="server" Text='<%# Bind("memo") %>' Width = "40px" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                 <asp:BoundField DataField="itemdate" HeaderText="日期" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                 <asp:BoundField DataField="itemcreator" HeaderText="人員" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
            </Columns>
        </asp:GridView>
        <br />
        <center><a herf ="#" onclick = "window.close();"><asp:Button ID="btClose" runat="server" Text="關閉視窗" /></a> <asp:Label ID="lbErrMsg" runat="server" Width="100%" Font-Bold="True" ForeColor="Red" /></center>
    </div>
           <asp:SqlDataSource ID="sqldatalist" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [P_0414] WHERE [shcode] = @shcode" 
            SelectCommand="select shcode,shcost,company,expired_y,UseCheck,Buy_num,a.Seat_num+'_'+Seat_name as seatnum,a.Seat_num,company_Num,memo,
                    case usecheck when '0' then StockDate when '1' then AppDate when '2' then UseDate end as itemdate,c.emp_chinese_name as itemcreator
                from P_0414 a with(nolock) left join P_0417 b with(nolock) on a.Seat_num = b.Seat_num
                left join employee c with(nolock) on (case a.usecheck when '0' then StockKeyIn when '1' then AppKeyIn when '2' then UseKeyIn end) = c.employee_id
                where shtype = '0' and substring(shcode,1,6) =@it_code and usecheck=@usecheck
                Order by shcode DESC;"
            UpdateCommand="Update P_0414 Set shcost = @shcost, company = @company, expired_y = @expired_y, 
                            Buy_Num = @Buy_Num, Seat_Num = @Seat_Num, company_Num = @company_Num,
                            memo = @memo Where shcode = @shcode;"
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
            <SelectParameters>
                <asp:ControlParameter ControlID="lbl_it_code" Name="it_code" PropertyName="Text" type="String" />
                <asp:ControlParameter ControlID="lbl_usecheck" Name="usecheck" PropertyName="Text" type="String" />
            </SelectParameters>
            <DeleteParameters>
                <asp:Parameter Name="shcode" Type="String" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="shcost" Type="String" />
                <asp:Parameter Name="company" Type="String" />
                <asp:Parameter Name="expired_y" Type="String" />
                <asp:Parameter Name="Buy_Num" Type="String" />
                <asp:Parameter Name="Seat_Num" Type="String" />
                <asp:Parameter Name="company_Num" Type="String" />
                <asp:Parameter Name="memo" Type="String" />
            </UpdateParameters>
        </asp:SqlDataSource>    
          <asp:SqlDataSource ID="sqlds_Seat_Num" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [seat_num]+'_'+[seat_name] as seatnum,[seat_num], [seat_name] FROM [P_0417] ORDER BY [seat_num]"
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
        </asp:SqlDataSource>        
    </form>
</body>
</html>

