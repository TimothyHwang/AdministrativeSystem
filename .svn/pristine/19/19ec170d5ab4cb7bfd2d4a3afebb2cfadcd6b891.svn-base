<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04023.aspx.vb" Inherits="Source_04_MOA04023" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>現勘人員管理</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="現勘人員管理" Width="100%"></asp:Label>
            </td></tr>
        </table>
        <asp:DetailsView ID="DetailsView1" runat="server" AutoGenerateRows="False" DataKeyNames="House_Num" DataSourceID="SqlDataSource1" DefaultMode="Insert" Width="100%" CssClass="form">
            <Fields>
                <asp:TemplateField SortExpression="House_Num" ShowHeader="False">
                    <InsertItemTemplate>                   
                        <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
                            <tr>
                                <td noWrap width="5%" class="form">
                                    <asp:Label ID="Lab1" runat="server" Text="現勘人員名稱："></asp:Label>
                                </td>
                                <td noWrap width="25%" class="form">
                                    <asp:TextBox ID="House_Name" runat="server" Text='<%# Bind("House_Name") %>' MaxLength='20' Width="150px"></asp:TextBox>
                                </td>
                                <td noWrap width="5%" class="form">
                                    <asp:Label ID="Lab2" runat="server" Text="現勘人員電話：" ></asp:Label>
                                </td>
                                <td noWrap width="25%" class="form">
                                    <asp:TextBox ID="House_Tel" runat="server" Text='<%# Bind("House_Tel") %>'  MaxLength='10' Width="100px"></asp:TextBox>
                                </td>
                                <td noWrap width="5%" class="form">
                                    <asp:Label ID="Label1" runat="server" Text="現勘人員狀況：" ></asp:Label>
                                </td>
                                <td noWrap width="25%" class="form">
                                    <asp:DropDownList id="House_Status"  SelectedValue='<%# Bind("House_Status") %>'
                                        AutoPostBack="True"
                                        runat="server">
                                        <asp:ListItem Value="">  </asp:ListItem>
                                        <asp:ListItem Value="1"> 內派 </asp:ListItem>
                                        <asp:ListItem Value="2"> 外包 </asp:ListItem>
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
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" DataKeyNames="House_Num" DataSourceID="SqlDataSource1" EmptyDataText="查無任何資料" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
            <Columns>
                <asp:BoundField DataField="House_Name" HeaderText="現勘人員名稱" SortExpression="House_Name">
                    <ItemStyle HorizontalAlign="Center" Width="30%" />
                </asp:BoundField>
                <asp:BoundField  DataField="House_Tel" HeaderText="現勘人員電話" SortExpression="House_Tel">
                    <ItemStyle HorizontalAlign="Center" Width="30%" />
                </asp:BoundField>
                <asp:templatefield HeaderText="現勘人員狀況"  SortExpression="House_Status">
                    <itemtemplate>
                        <asp:label id="LastNameLabel"
                        text= '<%# Chg_House_Status("House_Status") %>'
                        runat="server"/>
                    </itemtemplate>
                    <EditItemTemplate>
                        <asp:DropDownList id="House_Status2" SelectedValue='<%# Bind("House_Status") %>'
                            AutoPostBack="True"
                            runat="server" >
                            <asp:ListItem Value=""></asp:ListItem>
                            <asp:ListItem Value="1"> 內派 </asp:ListItem>
                            <asp:ListItem Value="2"> 外包 </asp:ListItem>
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
            DeleteCommand="DELETE FROM [P_0416] WHERE [House_Num] = @House_Num" 
            InsertCommand="INSERT INTO [P_0416] ([House_Name], [House_Tel],[House_Status]) VALUES (@House_Name, @House_Tel,@House_Status)"
            SelectCommand="SELECT [House_Name], [House_Tel], [House_Num],[House_Status] FROM [P_0416] ORDER BY [House_Num]"
            UpdateCommand="UPDATE [P_0416] SET [House_Name] = @House_Name, [House_Tel] = @House_Tel, [House_Status]=@House_Status WHERE [House_Num] = @House_Num"
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
            <DeleteParameters>
                <asp:Parameter Name="House_Num" Type="Int32" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="House_Name" Type="String" />
                <asp:Parameter Name="House_Tel" Type="String" />
                <asp:Parameter Name="House_Status" Type="String" />
                <asp:Parameter Name="House_Num" Type="Int32" />
            </UpdateParameters>
            <InsertParameters>
                <asp:Parameter Name="House_Name" Type="String" />
                <asp:Parameter Name="House_Tel" Type="String" />
                <asp:Parameter Name="House_Status" Type="String" />
            </InsertParameters>
        </asp:SqlDataSource>        
    </form>
</body>
</html>
