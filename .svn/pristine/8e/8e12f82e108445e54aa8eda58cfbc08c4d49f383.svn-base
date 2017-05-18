<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA03007.aspx.vb" Inherits="Source_03_MOA03007" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>駕駛登錄</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="駕駛登錄" Width="100%"></asp:Label>
            </td></tr>
        </table>
        <asp:DetailsView ID="DetailsView1" runat="server" AutoGenerateRows="False" DataKeyNames="Drive_Num" DataSourceID="SqlDataSource1" DefaultMode="Insert" Width="100%" CssClass="form">
            <Fields>
                <asp:TemplateField SortExpression="Drive_Num" ShowHeader="False">
                    <InsertItemTemplate>                    
                        <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
                            <tr>
                                <td noWrap width="5%" class="form">
                                    <asp:Label ID="Lab1" runat="server" Text="駕駛名稱："></asp:Label>
                                </td>
                                <td noWrap width="25%" class="form">
                                    <asp:TextBox ID="Drive_Name" runat="server" Text='<%# Bind("Drive_Name") %>' MaxLength='20' Width="150px"></asp:TextBox>
                                </td>
                                <td noWrap width="5%" class="form">
                                    <asp:Label ID="Lab2" runat="server" Text="駕駛電話：" ></asp:Label>
                                </td>
                                <td noWrap width="25%" class="form">
                                    <asp:TextBox ID="Drive_Tel" runat="server" Text='<%# Bind("Drive_Tel") %>'  MaxLength='10' Width="100px"></asp:TextBox>
                                </td>
                                <td noWrap width="5%" class="form">
                                    <asp:Label ID="Label1" runat="server" Text="駕駛狀況：" ></asp:Label>
                                </td>
                                <td noWrap width="25%" class="form">
                                    <asp:DropDownList id="Drive_Status"  SelectedValue='<%# Bind("Drive_Status") %>'
                                        AutoPostBack="True"
                                        runat="server">
                                        <asp:ListItem Value="">  </asp:ListItem>
                                        <asp:ListItem Value="1"> 到勤 </asp:ListItem>
                                        <asp:ListItem Value="2"> 休假 </asp:ListItem>
                                        <asp:ListItem Value="3"> 公差 </asp:ListItem>
                                    </asp:DropDownList>
                                </td>
                                <td noWrap width="10%" class="form">
                                    <asp:ImageButton ID="ImgInsert" runat="server" ImageUrl="~/Image/add.gif" OnClick="ImgInsert_Click"
                                        ToolTip="新增" />
                                    <asp:Label ID="ErrMsg" runat="server" ForeColor='Red' Text="" ></asp:Label>
                                </td>
                            </tr>
                        </table> 
                    </InsertItemTemplate>
                </asp:TemplateField>
                <asp:CommandField />
            </Fields>
        </asp:DetailsView>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" DataKeyNames="Drive_Num" DataSourceID="SqlDataSource1" EmptyDataText="查無任何資料" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
            <Columns>
                <asp:BoundField DataField="Drive_Name" HeaderText="駕駛名稱" SortExpression="Drive_Name">
                    <ItemStyle HorizontalAlign="Center" Width="30%" />
                </asp:BoundField>
                <asp:BoundField  DataField="Drive_Tel" HeaderText="駕駛電話" SortExpression="Drive_Tel">
                    <ItemStyle HorizontalAlign="Center" Width="30%" />
                </asp:BoundField>
                <asp:templatefield HeaderText="駕駛狀況"  SortExpression="Drive_Status">
                    <itemtemplate>
                        <asp:label id="LastNameLabel"
                        text= '<%# Chg_Drive_Status("Drive_Status") %>'
                        runat="server"/>
                    </itemtemplate>
                    <EditItemTemplate>
                        <asp:DropDownList id="Drive_Status2" SelectedValue='<%# Bind("Drive_Status") %>'
                            AutoPostBack="True"
                            runat="server" >
                            <asp:ListItem Value=""></asp:ListItem>
                            <asp:ListItem Value="1"> 到勤 </asp:ListItem>
                            <asp:ListItem Value="2"> 休假 </asp:ListItem>
                            <asp:ListItem Value="3"> 公差 </asp:ListItem>
                        </asp:DropDownList>
                    </EditItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:templatefield>
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
            DeleteCommand="DELETE FROM [P_0304] WHERE [Drive_Num] = @Drive_Num" 
            InsertCommand="INSERT INTO [P_0304] ([Drive_Name], [Drive_Tel],[Drive_Status]) VALUES (@Drive_Name, @Drive_Tel,@Drive_Status)"
            SelectCommand="SELECT [Drive_Name], [Drive_Tel], [Drive_Num],[Drive_Status] FROM [P_0304] ORDER BY [Drive_Num]"
            UpdateCommand="UPDATE [P_0304] SET [Drive_Name] = @Drive_Name, [Drive_Tel] = @Drive_Tel, [Drive_Status]=@Drive_Status WHERE [Drive_Num] = @Drive_Num"
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
            <DeleteParameters>
                <asp:Parameter Name="Drive_Num" Type="Int32" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="Drive_Name" Type="String" />
                <asp:Parameter Name="Drive_Tel" Type="String" />
                <asp:Parameter Name="Drive_Status" Type="String" />
                <asp:Parameter Name="Drive_Num" Type="Int32" />
            </UpdateParameters>
            <InsertParameters>
                <asp:Parameter Name="Drive_Name" Type="String" />
                <asp:Parameter Name="Drive_Tel" Type="String" />
                <asp:Parameter Name="Drive_Status" Type="String" />
            </InsertParameters>
        </asp:SqlDataSource>        
    </form>
</body>
</html>
