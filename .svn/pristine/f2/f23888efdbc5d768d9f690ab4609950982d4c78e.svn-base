<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04025_1.aspx.vb" Inherits="M_Source_04_MOA04025_1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>物料詳細資料</title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
    <div>
          <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center" colspan = "4">
                    <asp:Label ID="lblTitle" runat="server" CssClass="toptitle" Width="100%"></asp:Label>
                </td>
            </tr>
            <tr>
            <td align="center" colspan = "4"> <asp:Label ID="ErrMsg" runat="server" ForeColor="Red" Font-Bold="True" /></td>
            </tr>
            <asp:Repeater ID="RptitemList" runat="server" ><ItemTemplate>
            <tr>
                <td colspan = "4">物料名稱： <asp:Label ID="lb_it_name" runat="server" Text ='<%# Bind("it_name") %>' /></td>
            </tr>
            <tr>
                <td align ="left" valign ="middle">物料圖片1：</td>
                <td><asp:Image ID="Image1" runat="server" ImageUrl ='<%# sPicPath + Eval("it_code")+ "A_" + Eval("file_a") %>' Visible='<%# showPic(Eval("file_a"))%>' />
                <asp:Label ID="lbImage1" runat="server" Text ="此物料無上傳圖片一" Visible='<%# showPicLB(Eval("file_a"))%>' /></td>
                <td align ="left" valign ="middle">物料圖片2：</td>
                <td><asp:Image ID="Image2" runat="server" ImageUrl ='<%# sPicPath + Eval("it_code")+ "B_" + Eval("file_b") %>' Visible='<%# showPic(Eval("file_b"))%>' />
                <asp:Label ID="lbImage2" runat="server" Text ="此物料無上傳圖片二" Visible='<%# showPicLB(Eval("file_a"))%>'/></td>
            </tr>
            </ItemTemplate></asp:Repeater>
            <asp:Label ID="lbl_it_code" runat="server" Visible="false" />
            <asp:Label ID="lbl_usecheck" runat="server" Visible="false" />

            <asp:GridView ID="GVItemList" runat="server" AllowPaging="True" AllowSorting="True" EmptyDataText="查無任何資料" AutoGenerateColumns="False" 
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
                 <asp:TemplateField ShowHeader="False">
                        <EditItemTemplate>
                            <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="True" CommandName="Update" Text="更新" />
                            <asp:LinkButton ID="LinkButton2" runat="server" CausesValidation="False" CommandName="Cancel" Text="取消" />
                        </EditItemTemplate>
                        <ItemTemplate>
                            <asp:ImageButton ID="ImageButton1" runat="server" CausesValidation="False" CommandName="Edit" ImageUrl ="~/Image/update.gif" Visible='<%# showeditdelLB(Eval("usecheck"))%>'  />
                            <asp:ImageButton ID="ImageButton2" runat="server" CausesValidation="False" CommandName="Delete" OnClientClick="return confirm('您確定要刪除嗎?')" ImageUrl ="~/Image/delete.gif"  Visible='<%# showeditdelLB(Eval("usecheck"))%>' />
                        </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="shcode"  HeaderText="物品編號" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                  <asp:TemplateField HeaderText="物料規格">
                     <EditItemTemplate>
                        <asp:TextBox id = "tb_it_spec" runat="server" Text='<%# Bind("it_spec") %>' />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label id = "lb_it_spec" runat="server" Text='<%# Bind("it_spec") %>' />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
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
                <asp:BoundField DataField="editor" HeaderText="建檔人員" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="領取人員">
                     <EditItemTemplate>
                        <asp:TextBox id = "tb_receive" runat="server" Text='<%# Bind("receive") %>' Width = "50px" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label id = "lb_receive" runat="server" Text='<%# Bind("receive") %>' />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>

                  <asp:BoundField DataField="editdate" HeaderText="時間" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                 <asp:TemplateField>
                    <EditItemTemplate>
                        
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Button ID="BTOut" runat="server" CausesValidation="False" CommandArgument=<%#Eval("shcode") %> CommandName="UPOut" text="出庫" Visible='<%# showOutLB(Eval("usecheck"))%>'  />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
        </table>
        <br />
        <center><a herf ="#" onclick = "window.close();"><asp:Button ID="btClose" runat="server" Text="關閉視窗" /></a></center>
    </div>
    <asp:SqlDataSource ID="sqldatalist" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [P_0414] WHERE [shcode] = @shcode" 
            SelectCommand="select a.shcode,a.it_spec,a.Seat_num+'_'+c.Seat_name as seatnum,a.seat_num,a.receive,
                            case usecheck when '0' then StockDate when '2' then UseDate end as editdate,
                            case usecheck when '0' then StockKeyIn when '2' then UseKeyIn end as editor,usecheck
                            from P_0414 a with(nolock) left join P_0407 b with(nolock) on substring(a.shcode,1,6) = b.it_code
                            left join P_0417 c with(nolock) on a.Seat_num = c.Seat_num
                            where b.it_code =@it_code and usecheck=@usecheck and shtype = '1'
                            Order by shcode DESC"
            UpdateCommand="Update P_0414 Set it_spec=@it_spec, Seat_Num = @Seat_Num , receive=@receive Where shcode = @shcode"
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
            <SelectParameters>
                <asp:ControlParameter ControlID="lbl_it_code" Name="it_code" PropertyName="Text" type="String" />
                <asp:ControlParameter ControlID="lbl_usecheck" Name="usecheck" PropertyName="Text" type="String" />
            </SelectParameters>
            <DeleteParameters>
                <asp:Parameter Name="shcode" Type="String" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="shcode" Type="String" />
                <asp:Parameter Name="it_spec" Type="String" />
                <asp:Parameter Name="Seat_Num" Type="String" />
                <asp:Parameter Name="receive" Type="String" />
            </UpdateParameters>
            </asp:SqlDataSource>
    <asp:SqlDataSource ID="sqlds_Seat_Num" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
    SelectCommand="SELECT [seat_num]+'_'+[seat_name] as seatnum,[seat_num], [seat_name] FROM [P_0417] ORDER BY [seat_num]"
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
        </asp:SqlDataSource>        
    </form>
</body>
</html>
