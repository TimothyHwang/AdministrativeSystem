<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA10007.aspx.vb" Inherits="M_Source_10_MOA10007" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>主官在營公告維護</title>
    <link href='<%#ResolveUrl("~/css/site.css") %>' rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:DetailsView ID="dvP_10Mqrquee" runat="server" AutoGenerateRows="False" 
            DataKeyNames="CONFIG_VAR" DataSourceID="SqlDataSourceMarquee" 
            EnableModelValidation="True" Height="50px" Caption="主官在營公告維護" 
            CaptionAlign="Top" CssClass="tableView" BorderStyle="Solid" 
            BorderWidth="1px" CellPadding="0" Width="100%">
            <Fields>
                <asp:BoundField DataField="CONFIG_NUM" HeaderText="CONFIG_NUM" 
                    InsertVisible="False" ReadOnly="True" SortExpression="CONFIG_NUM" 
                    Visible="False" />
                <asp:BoundField DataField="CONFIG_VAR" HeaderText="CONFIG_VAR" ReadOnly="True" 
                    SortExpression="CONFIG_VAR" Visible="False" />
                <asp:BoundField DataField="CONFIG_DESC" HeaderText="公告說明" ReadOnly="True" 
                    SortExpression="CONFIG_DESC" >
                <ControlStyle Width="100px" />
                <HeaderStyle Width="100px" />
                <ItemStyle Width="450px" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="公告內容" SortExpression="CONFIG_VALUE">
                    <EditItemTemplate>
                        <asp:TextBox ID="txtMarquee" runat="server" Text='<%# Bind("CONFIG_VALUE") %>' 
                            Rows="10" TextMode="MultiLine" Width="100%"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lblMarquee" runat="server" Text='<%# Bind("CONFIG_VALUE") %>' 
                            Width="100%"></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle Width="100px" />
                </asp:TemplateField>
                <asp:TemplateField InsertVisible="False" ShowHeader="False">
                    <EditItemTemplate>
                        <asp:ImageButton ID="ImageButton1" runat="server" CausesValidation="True" 
                            CommandArgument='<%#Eval("CONFIG_NUM")%>' CommandName="Update" 
                            ImageUrl="~/Image/apply.gif" Text="更新" />
                        &nbsp;<asp:ImageButton ID="ImageButton2" runat="server" CausesValidation="False" 
                            CommandName="Cancel" ImageUrl="~/Image/cancel.gif" Text="取消" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:ImageButton ID="ImageButton1" runat="server" CausesValidation="False" 
                            CommandName="Edit" ImageUrl="~/Image/update.gif" Text="編輯" />
                    </ItemTemplate>
                </asp:TemplateField>
            </Fields>
        </asp:DetailsView>
        <asp:Label ID="Label1" runat="server" ForeColor="Red" 
            Text="*公告內容不可超過50個中文字"></asp:Label>
        <asp:SqlDataSource ID="SqlDataSourceMarquee" runat="server" 
            ConnectionString="<%$ ConnectionStrings:ConnectionString %>" 
            
            SelectCommand="SELECT * FROM [SYSCONFIG] WHERE ([CONFIG_VAR] = @CONFIG_VAR)" 
            UpdateCommand="UPDATE [SYSCONFIG] SET [CONFIG_DESC] = @CONFIG_DESC, [CONFIG_VALUE] = @CONFIG_VALUE WHERE [CONFIG_VAR] = @original_CONFIG_VAR AND [CONFIG_NUM] = @original_CONFIG_NUM AND (([CONFIG_DESC] = @original_CONFIG_DESC) OR ([CONFIG_DESC] IS NULL AND @original_CONFIG_DESC IS NULL)) AND (([CONFIG_VALUE] = @original_CONFIG_VALUE) OR ([CONFIG_VALUE] IS NULL AND @original_CONFIG_VALUE IS NULL))"
            ConflictDetection="CompareAllValues" OldValuesParameterFormatString="original_{0}"
            >

            <SelectParameters>
                <asp:Parameter DefaultValue="P_10Marquee" Name="CONFIG_VAR" Type="String" />
            </SelectParameters>
            <UpdateParameters>
                <asp:Parameter Name="CONFIG_NUM" Type="Decimal" />
                <asp:Parameter Name="CONFIG_DESC" Type="String" />
                <asp:Parameter Name="CONFIG_VALUE" Type="String" />
                <asp:Parameter Name="original_CONFIG_VAR" Type="String" />
                <asp:Parameter Name="original_CONFIG_NUM" Type="Decimal" />
                <asp:Parameter Name="original_CONFIG_DESC" Type="String" />
                <asp:Parameter Name="original_CONFIG_VALUE" Type="String" />
            </UpdateParameters>
        </asp:SqlDataSource>
    
    </div>
    </form>
</body>
</html>
