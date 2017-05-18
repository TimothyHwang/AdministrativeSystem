<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00001.aspx.vb" Inherits="Source_00_MOA00001" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>表單製作</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />    
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
    <div>
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label6" runat="server" CssClass="toptitle" Text="表單製作" Width="100%"></asp:Label>
            </td></tr>
        </table>
        
        <asp:ImageButton ID="ImgBtn1" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" /><br />
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
            AutoGenerateColumns="False" DataKeyNames="eformid" DataSourceID="SqlDataSource1"
            Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None" HorizontalAlign="Center">
            <Columns>
                <asp:BoundField DataField="eformid" HeaderText="eformid" ReadOnly="True" SortExpression="eformid" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="frm_chinese_name" HeaderText="表單名稱" SortExpression="frm_chinese_name" >
                    <HeaderStyle HorizontalAlign="Center" Width="40%" />
                </asp:BoundField>
                <asp:BoundField DataField="frm_english_name" HeaderText="frm_english_name" SortExpression="frm_english_name" Visible="False" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="organization_id" HeaderText="organization_id" SortExpression="organization_id" Visible="False" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="PRIMARY_TABLE" HeaderText="表單資料表" SortExpression="PRIMARY_TABLE" >
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <EditItemTemplate>
                        <asp:ImageButton ID="ImageButton1" runat="server" CommandName="Update" ImageUrl="~/Image/apply.gif"
                            ToolTip="確認" />&nbsp;<asp:ImageButton ID="ImageButton2" runat="server" CommandName="Cancel"
                                ImageUrl="~/Image/cancel.gif" ToolTip="取消" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:ImageButton ID="ImageButton1" runat="server" CommandName="Edit" ImageUrl="~/Image/update.gif" ToolTip="修改" />
                        <asp:ImageButton ID="ImageButton2" runat="server" CommandName="Delete" ImageUrl="~/Image/delete.gif"
                            OnClientClick="return confirm('確定刪除表單嗎!!!\n刪除將會影響舊有表單資料!!!')" ToolTip="刪除" 
                            Visible="False" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                </asp:TemplateField>
                <asp:CommandField ShowSelectButton="True" ButtonType="Button" SelectText="表單流程" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                </asp:CommandField>
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
            DeleteCommand="DELETE FROM [EFORMS] WHERE [eformid] = @eformid" InsertCommand="INSERT INTO [EFORMS] ([eformid], [frm_chinese_name], [frm_english_name], [PRIMARY_TABLE], [organization_id]) VALUES (@eformid, @frm_chinese_name, @frm_english_name, @PRIMARY_TABLE, @organization_id)"
            SelectCommand="SELECT [eformid], [frm_chinese_name], [frm_english_name], [PRIMARY_TABLE], [organization_id] FROM [EFORMS]"
            UpdateCommand="UPDATE [EFORMS] SET [frm_chinese_name] = @frm_chinese_name, [PRIMARY_TABLE] = @PRIMARY_TABLE WHERE [eformid] = @eformid">
            <DeleteParameters>
                <asp:Parameter Name="eformid" Type="String" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="frm_chinese_name" Type="String" />
                <asp:Parameter Name="frm_english_name" Type="String" />
                <asp:Parameter Name="PRIMARY_TABLE" Type="String" />
                <asp:Parameter Name="organization_id" Type="String" />
                <asp:Parameter Name="eformid" Type="String" />
            </UpdateParameters>
            <InsertParameters>
                <asp:Parameter Name="eformid" Type="String" />
                <asp:Parameter Name="frm_chinese_name" Type="String" />
                <asp:Parameter Name="frm_english_name" Type="String" />
                <asp:Parameter Name="PRIMARY_TABLE" Type="String" />
                <asp:Parameter Name="organization_id" Type="String" />
            </InsertParameters>
        </asp:SqlDataSource>
    
    </div>
    </form>
</body>
</html>
