<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA03004.aspx.vb" Inherits="Source_03_MOA03004" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>車種登錄</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="車種登錄" Width="100%"></asp:Label>
            </td></tr>
        </table>
        <asp:DetailsView ID="DetailsView1" runat="server" AutoGenerateRows="False" DataKeyNames="PCK_Num" DataSourceID="SqlDataSource1" DefaultMode="Insert"  Width="100%" CssClass="form">
            <Fields>
                <asp:TemplateField SortExpression="PCK_Num" ShowHeader="False">
                    <InsertItemTemplate>
                        <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
                            <tr>
                                <td noWrap width="5%" class="form">
                                    <asp:Label ID="Lab1" runat="server" Text="車種："></asp:Label>
                                </td>
                                <td noWrap width="25%" class="form">
                                    <asp:TextBox ID="PCK_Name" runat="server" Text='<%# Bind("PCK_Name") %>' MaxLength='25' Width="150px"></asp:TextBox>
                                </td>
                                <td noWrap width="5%" class="form">
                                    <asp:Label ID="Lab2" runat="server" Text="耗油量：" ></asp:Label>
                                </td>
                                <td noWrap width="25%" class="form">
                                    <asp:TextBox ID="PCK_OilLose" runat="server" Text='<%# Bind("PCK_OilLose") %>'  MaxLength='15' Width="100px"></asp:TextBox>
                                </td>
                                <td noWrap width="5%" class="form">
                                    <asp:Label ID="Lab3" runat="server" Text="油品：" ></asp:Label>
                                </td>
                                <td noWrap width="25%" class="form">
                                    <asp:DropDownList id="PCK_OilItem"  SelectedValue='<%# Bind("PCK_OilItem") %>'
                                        AutoPostBack="True" runat="server">
                                        <asp:ListItem Value="">  </asp:ListItem>
                                        <asp:ListItem Value="1"> 汽油 </asp:ListItem>
                                        <asp:ListItem Value="2"> 柴油 </asp:ListItem>
                                    </asp:DropDownList>
                                </td>
                                <td noWrap width="10%" class="form">
                                    <asp:ImageButton ID="ImgInsert" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" OnClick="ImgInsert_Click" />
                                    <asp:Label ID="ErrMsg" runat="server" ForeColor='Red' Text="" ></asp:Label>
                                </td>
                            </tr>
                        </table>                        
                        
                    </InsertItemTemplate>
                </asp:TemplateField>
                <asp:CommandField />
            </Fields>
        </asp:DetailsView>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" DataKeyNames="PCK_Num" DataSourceID="SqlDataSource1" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
            <Columns>
                <asp:BoundField DataField="PCK_Name" HeaderText="車種" SortExpression="PCK_Name">
                    <ItemStyle HorizontalAlign="Center" Width="30%" />
                </asp:BoundField>
                <asp:BoundField  DataField="PCK_OilLose" HeaderText="耗油量" SortExpression="PCK_OilLose">
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:templatefield HeaderText="油品"  SortExpression="PCK_OilItem">
                    <itemtemplate>
                        <asp:label id="LastNameLabel"
                        text= '<%# Chg_OilItem("PCK_OilItem") %>'
                        runat="server"/>
                    </itemtemplate>
                    <EditItemTemplate>
                        <asp:DropDownList id="PCK_OilItem2" SelectedValue='<%# Bind("PCK_OilItem") %>'
                            AutoPostBack="True"
                            runat="server" >
                            <asp:ListItem Value=""></asp:ListItem>
                            <asp:ListItem Value="1"> 汽油 </asp:ListItem>
                            <asp:ListItem Value="2"> 柴油 </asp:ListItem>
                        </asp:DropDownList>
                    </EditItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="30%" />
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
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [P_0302] WHERE [PCK_Num] = @PCK_Num" 
            InsertCommand="INSERT INTO [P_0302] ([PCK_Name], [PCK_OilLose],[PCK_OilItem]) VALUES (@PCK_Name, @PCK_OilLose, @PCK_OilItem)"
            SelectCommand="SELECT [PCK_Name], [PCK_OilLose], [PCK_Num],[PCK_OilItem] FROM [P_0302] ORDER BY [PCK_Num]"
            UpdateCommand="UPDATE [P_0302] SET [PCK_Name] = @PCK_Name, [PCK_OilLose] = @PCK_OilLose,[PCK_OilItem]=@PCK_OilItem WHERE [PCK_Num] = @PCK_Num"
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
            <DeleteParameters>
                <asp:Parameter Name="PCK_Num" Type="Int32" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="PCK_Name" Type="String" />
                <asp:Parameter Name="PCK_OilLose" Type="String" />
                <asp:Parameter Name="PCK_OilItem"  Type="String" />
                <asp:Parameter Name="PCK_Num" Type="Int32" />
            </UpdateParameters>
            <InsertParameters>
                <asp:Parameter Name="PCK_Name" Type="String" />
                <asp:Parameter Name="PCK_OilLose" Type="String" />
                <asp:Parameter Name="PCK_OilItem"  Type="String" />
            </InsertParameters>
        </asp:SqlDataSource>        
    </form>
</body>
</html>
