<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04022_1.aspx.vb" Inherits="M_Source_04_MOA04022_1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>領料物品詳細資料</title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
    <div>
          <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center">
                    <asp:Label ID="lblTitle" runat="server" CssClass="toptitle" Text="倉儲領料-物品詳細資料" Width="100%"></asp:Label>
                </td>
            </tr>
            <tr>
            <td align="center"> <asp:Label ID="ErrMsg" runat="server" ForeColor="Red" Font-Bold="True" /></td>
            </tr>
            
            <asp:Label ID="lbl_EFORMSN" runat="server" Visible="false" />

            <asp:GridView ID="GVItemList" runat="server" AllowSorting="True" EmptyDataText="查無任何資料" AutoGenerateColumns="False" 
            Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None" DataKeyNames="shcode" DataSourceID="sqldatalist" >
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <Columns>
                <asp:TemplateField HeaderText="領料狀態" >
                    <ItemTemplate>
                        <asp:CheckBox ID="ck_GetForm" runat="server" text="領料" Enabled='<%# enableGetCK(Eval("usecheck"))%>' Checked ='<%# showcheckbox(Eval("usecheck"))%>'  />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                  <asp:TemplateField HeaderText="物品編號" >
                    <ItemTemplate>
                        <asp:Label ID="lb_shcode" runat="server" Text='<%# Eval("shcode") %>' />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
               <asp:BoundField DataField="it_name" HeaderText="物料名稱" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" Width="30%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                  <asp:BoundField DataField="expired_y" HeaderText="有效期" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                    <asp:BoundField DataField="seat_num" HeaderText="儲庫編號" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
            </Columns>
        </asp:GridView>
        </table>
        <br />
        <center>
        <asp:Button ID="btGetItem" runat="server" Text="勾選者領料" Enabled = "false"/>
        <asp:Button ID="btFinishEForm" runat="server" Text="完結" />
        <a herf ="#" onclick = "window.close();"><asp:Button ID="btClose" runat="server" Text="關閉視窗" /></a></center>
    </div>
    <asp:SqlDataSource ID="sqldatalist" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="select shcode,b.it_name,a.expired_y,c.seat_num+'_'+c.seat_name as seat_num,usecheck
                            from P_0414 a with(nolock) left join P_0407 b with(nolock) on  substring(a.shcode,1,6) = b.it_code
                            left join P_0417 c with(nolock) on a.seat_num = c.seat_num
                            where job_num = @EFORMSN and usecheck in (1,2) Order by usecheck"
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
            <SelectParameters>
                <asp:ControlParameter ControlID="lbl_EFORMSN" Name="EFORMSN" PropertyName="Text" type="String" />
            </SelectParameters>
            </asp:SqlDataSource>
    </form>
</body>
</html>

