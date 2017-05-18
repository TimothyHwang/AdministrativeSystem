<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA10005.aspx.vb" Inherits="M_Source_10_MOA10005" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>主官在營作業操作</title>
    <link href='<%#ResolveUrl("~/css/site.css") %>' rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    
        <asp:GridView ID="gvManagerEmployee" runat="server" 
            AllowSorting="True" EmptyDataText="查無任何資料"
            AutoGenerateColumns="False" DataSourceID="SqlDataSourceP_10" 
            Width="100%" DataKeyNames="P_Num" CellPadding="4" ForeColor="#333333" 
            GridLines="None" EnableModelValidation="True" Caption="主官在營作業操作" 
            CaptionAlign="Top" CssClass="tableView">
            <Columns>
                <asp:BoundField DataField="P_Num" SortExpression="P_Num" HeaderText="P_Num" 
                    InsertVisible="False" ReadOnly="True" Visible="False" >
                </asp:BoundField>
                <asp:TemplateField HeaderText="主官名稱" SortExpression="Manager_Id">
                    <ItemTemplate>
                        <asp:Label ID="lblManager" runat="server" Text='<%# Bind("Name") %>'></asp:Label>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:TemplateField>
                <asp:TemplateField ShowHeader="False" HeaderText="在營狀態">
                    <ItemTemplate>
                        <asp:ImageButton ID="ibnOffStatus" runat="server" CausesValidation="False" 
                            CommandName="Offline" ImageUrl="~/Image/ManagerOffline.gif" 
                            Text="選取" Visible="False" Width="50px" 
                            CommandArgument="<%# CType(Container,GridViewRow).RowIndex %>" />
                        <asp:ImageButton ID="ibnOnStatus" runat="server" CausesValidation="False" 
                            CommandName="Online" ImageUrl="~/Image/ManagerOnline.jpg" 
                            Text="選取" Visible="False" Width="50px" 
                            CommandArgument="<%# CType(Container,GridViewRow).RowIndex %>" />
                        <asp:Label ID="lblStatus" runat="server" Text='<%# Eval("Status") %>' 
                            Visible="False"></asp:Label>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
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
        <asp:SqlDataSource ID="SqlDataSourceP_10" runat="server" 
            ConnectionString="<%$ ConnectionStrings:ConnectionString %>" 
            
        
            
            SelectCommand="SELECT A.* FROM [P_10] A,[P_1001] B,[EMPLOYEE] C WHERE (A.[P_NUM]=B.MANAGER_ID AND B.[Employee_Id]=C.[EMPLOYEE_ID]) AND (C.[EMPLOYEE_ID] = @Employee_Id)">
            <SelectParameters>
                <asp:SessionParameter Name="Employee_Id" SessionField="user_id" Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
    
    <div>
    
    </div>
    </form>
</body>
</html>
