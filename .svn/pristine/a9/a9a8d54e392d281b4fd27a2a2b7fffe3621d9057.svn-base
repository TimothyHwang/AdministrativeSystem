<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA10004.aspx.vb" Inherits="M_Source_10_MOA10004" %>

<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="cc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>主官在營清單</title>
    
    <link href="../../css/site.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table style="width: 100%;">
            <tr>
                <td align="left" width="33%">
                    <asp:Label ID="Label1" runat="server" Text="主官名稱："></asp:Label>
                    &nbsp;
                    <asp:DropDownList ID="ddlManager" runat="server" 
                        DataSourceID="SqlDataSourceManager" DataTextField="Name" DataValueField="P_Num">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceManager" runat="server" 
                        ConnectionString="<%$ ConnectionStrings:ConnectionString %>" 
                        SelectCommand="SELECT * FROM [P_10]"></asp:SqlDataSource>
                </td>
                <td align="left" width="33%">
                    &nbsp;
                    <asp:Label ID="Label2" runat="server" Text="負責人名稱："></asp:Label>
                    &nbsp;
                    <asp:DropDownList ID="ddlEmployee" runat="server" 
                        DataSourceID="SqlDataSourceEmployee" DataTextField="name" 
                        DataValueField="employee_id">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceEmployee" runat="server" 
                        ConnectionString="<%$ ConnectionStrings:ConnectionString %>" 
                        
                        
                        SelectCommand="select a.employee_id,c.org_name+'-'+b.emp_chinese_name as name from p_1001 a left join employee b on a.employee_id= str(b.empuid) left join admingroup c on b.org_uid=c.org_uid  GROUP BY A.EMPLOYEE_ID,B.EMP_CHINESE_NAME,C.ORG_UID,C.ORG_NAME order by c.org_uid">
                    </asp:SqlDataSource>
                </td>
                <td align="left">
                    &nbsp;
                    <asp:ImageButton ID="imbQuery" runat="server" ImageUrl="~/Image/search.gif" />
                    <asp:ImageButton ID="imbClear" runat="server" ImageUrl="~/Image/ClearAll.gif" />
                </td>
            </tr>
            </table>
        <asp:GridView ID="gvManagerEmployee" runat="server" AllowPaging="True" 
            AllowSorting="True" EmptyDataText="查無任何資料"
            AutoGenerateColumns="False" DataSourceID="SqlDataSourceP_10" 
            Width="100%" DataKeyNames="P_Num" CellPadding="4" ForeColor="#333333" 
            GridLines="None" EnableModelValidation="True" Caption="主官在營清單" 
            CaptionAlign="Top" CssClass="tableView">
            <Columns>
                <asp:BoundField DataField="P_Num" SortExpression="P_Num" HeaderText="P_Num" 
                    InsertVisible="False" ReadOnly="True" Visible="False" >
                </asp:BoundField>
                <asp:TemplateField HeaderText="主官名稱" SortExpression="Manager_Id">
                    <ItemTemplate>
                        <asp:Label ID="lblManager" runat="server" Text='<%# Bind("NAME") %>'></asp:Label>
                        <asp:Label ID="lblManagerID" runat="server" Visible="False" 
                            Text='<%# Eval("P_NUM") %>'></asp:Label>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="負責人名稱">
                    <ItemTemplate>
                        <asp:GridView ID="gvEmployee" runat="server" AutoGenerateColumns="False" 
                            DataSourceID="SqlDataSourceP_1001" EnableModelValidation="True" 
                            GridLines="None" onrowdatabound="gvEmployee_RowDataBound" ShowHeader="False">
                            <Columns>
                                <asp:BoundField DataField="P_NUM" HeaderText="P_NUM" InsertVisible="False" 
                                    ReadOnly="True" ShowHeader="False" SortExpression="P_NUM" Visible="False" />
                                <asp:BoundField DataField="MANAGER_ID" HeaderText="MANAGER_ID" 
                                    ShowHeader="False" SortExpression="MANAGER_ID" Visible="False" />
                                <asp:BoundField DataField="EMPLOYEE_ID" HeaderText="負責人" ShowHeader="False" 
                                    SortExpression="EMPLOYEE_ID" />
                            </Columns>
                        </asp:GridView>
                        <asp:SqlDataSource ID="SqlDataSourceP_1001" runat="server" 
                            ConnectionString="<%$ ConnectionStrings:ConnectionString %>" 
                            SelectCommand="SELECT * FROM [P_1001]"></asp:SqlDataSource>
                    </ItemTemplate>
                    <ItemStyle BorderStyle="None" HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:TemplateField>
                <asp:TemplateField>
                    <ItemTemplate>
                        &nbsp;<asp:ImageButton ID="ibnSelectEdit" runat="server" CausesValidation="False" 
                            CommandArgument="" 
                            CommandName="Select" ImageUrl="~/Image/APY.gif" Text="選取" />
                        &nbsp;
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:TemplateField>
            </Columns>
            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#EFF3FB" />
            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#2461BF" />
            <AlternatingRowStyle BackColor="White" />
                <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSourceP_10" runat="server" 
            ConnectionString="<%$ ConnectionStrings:ConnectionString %>" 
            SelectCommand="SELECT * FROM [P_10] ">
        </asp:SqlDataSource>
    
    </div>
    </form>
</body>
</html>
