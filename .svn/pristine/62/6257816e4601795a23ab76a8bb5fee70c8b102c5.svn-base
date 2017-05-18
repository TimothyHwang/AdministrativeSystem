<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04024.aspx.vb" Inherits="M_Source_04_MOA04024" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>儲庫編號管理</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
      <script type="text/JavaScript">
          function Check() {
              var d = document.form1;
              var re = /</;
              if (re.test(d.DetailsView1$tb_seat_num.value)) {
                  alert("您輸入的儲庫編號有包含危險字元！");
                  d.DetailsView1$tb_seat_num.value = "";
                  return false;
              }

              if (re.test(d.DetailsView1$tb_seat_name.value)) {
                  alert("您輸入的儲庫名稱有包含危險字元！");
                  d.DetailsView1$tb_seat_name.value = "";
                  return false;
              }
              return true;
          }
</script>
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="儲庫編號管理" Width="100%"></asp:Label>
            </td></tr>
        </table>
        <asp:DetailsView ID="DetailsView1" runat="server" AutoGenerateRows="False" DataKeyNames="seat_Num" DataSourceID="SqlDataSource1" DefaultMode="Insert" Width="100%" CssClass="form">
            <Fields>
                <asp:TemplateField SortExpression="seat_Num" ShowHeader="False">
                    <InsertItemTemplate>                    
                        <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
                            <tr>
                                <td noWrap width="5%" class="form">
                                    <asp:Label ID="Lab1" runat="server" Text="儲庫編號：" />
                                </td>
                                <td noWrap width="25%" class="form">
                                    <asp:TextBox ID="tb_seat_num" runat="server" Text='<%# Bind("seat_num") %>' MaxLength='10' Width="150px"></asp:TextBox>
                                </td>
                                <td noWrap width="5%" class="form">
                                    <asp:Label ID="Lab2" runat="server" Text="儲庫名稱：" />
                                </td>
                                <td noWrap width="25%" class="form">
                                    <asp:TextBox ID="tb_seat_name" runat="server" Text='<%# Bind("seat_name") %>'  MaxLength='50' Width="100px"></asp:TextBox>
                                </td>
                                <td noWrap width="10%" class="form">
                                    <asp:ImageButton ID="ImgInsert" runat="server" ImageUrl="~/Image/add.gif" OnClick="ImgInsert_Click"
                                        ToolTip="新增" OnClientClick = "return Check();" />
                                    <asp:Label ID="ErrMsg" runat="server" ForeColor='Red' Text="" ></asp:Label>
                                </td>
                            </tr>
                        </table> 
                    </InsertItemTemplate>
                </asp:TemplateField>
                <asp:CommandField />
            </Fields>
        </asp:DetailsView>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" DataKeyNames="Seat_Num" DataSourceID="SqlDataSource1" EmptyDataText="查無任何資料" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
            <Columns>
                <asp:templatefield HeaderText="儲庫編號"  SortExpression="seat_num">
                    <itemtemplate>
                        <asp:label id="lb_seat_num"
                        text= '<%# Bind("seat_num") %>'
                        runat="server"/>
                    </itemtemplate>
                    <ItemStyle HorizontalAlign="Center" Width="40%" />
                    </asp:templatefield> 
                <asp:BoundField  DataField="seat_name" HeaderText="儲庫名稱" SortExpression="seat_name">
                    <ItemStyle HorizontalAlign="Center" Width="40%" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <EditItemTemplate>
                        <asp:ImageButton ID="btnImgOK" runat="server" CommandName="Update" ImageUrl="~/Image/apply.gif"
                            ToolTip="確認" />&nbsp;<asp:ImageButton ID="btnImgCancel" runat="server" CommandName="Cancel"
                                ImageUrl="~/Image/cancel.gif" ToolTip="取消" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:ImageButton ID="ImgUpd" runat="server" CommandName="Edit" ImageUrl="~/Image/update.gif"
                            ToolTip="修改" />&nbsp;<asp:ImageButton ID="ImgDel" runat="server" CommandName="Delete"
                                ImageUrl="~/Image/delete.gif" OnClientClick="return confirm('確定刪除嗎?')" ToolTip="刪除" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:TemplateField>
            </Columns>
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [P_0417] WHERE [seat_num] = @seat_num" 
            InsertCommand="INSERT INTO [P_0417] ([seat_num], [seat_name]) VALUES (@seat_num, @seat_name)"
            SelectCommand="SELECT [seat_num], [seat_name] FROM [P_0417] ORDER BY [seat_num]"
            UpdateCommand="UPDATE [P_0417] SET [seat_name] = @seat_name WHERE [seat_num] = @seat_num"
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
            <DeleteParameters>
                <asp:Parameter Name="seat_num" Type="String" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="seat_num" Type="String" />
                <asp:Parameter Name="seat_name" Type="String" />
            </UpdateParameters>
            <InsertParameters>
                <asp:Parameter Name="seat_name" Type="String" />
                <asp:Parameter Name="seat_num" Type="String" />
            </InsertParameters>
        </asp:SqlDataSource>        
    </form>
</body>
</html>

