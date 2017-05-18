<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA01011.aspx.vb" Inherits="Source_01_MOA01011" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>文職人員加班設定</title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
    <script type="text/javascript" language="javascript" src="../Inc/inc_common.js"></script>
    <script type="text/javascript" language="javascript">
        document.oncontextmenu = new Function("return false");
        document.onselectstart = new Function("return false");
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center">
                    <asp:Label ID="lblTitle" runat="server" CssClass="toptitle" Text="文職人員加班設定" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>
        <asp:Panel ID="plConditionsContainer" runat="server">
            <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
                <tr>
                    <td style="width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label1" runat="server" Text="人員單位:" CssClass="form"></asp:Label>
                    </td>
                    <td style="width: 49%; height: 25px; padding-left: 20px; padding-right: 44px;">
                        <asp:DropDownList ID="ddl_org_picker" runat="server" AutoPostBack="true" DataSourceID="sdsOrganization" DataTextField="ORG_NAME" DataValueField="ORG_UID" Width="100%"></asp:DropDownList>
                    </td>
                    <td style="width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label2" runat="server" Text="人員姓名:" CssClass="form"></asp:Label>
                    </td>
                    <td style="width: 49%; height: 25px; padding-left: 20px; padding-right: 49px;">
                        <asp:DropDownList ID="ddl_user_picker" runat="server" DataSourceID="sdsUsersList" DataTextField="emp_chinese_name" DataValueField="employee_id" Width="100%"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label3" runat="server" CssClass="form" Text="人員編號:" Width="77px"></asp:Label>
                    </td>
                    <td width="15%" class="form" style="height: 25px; white-space: nowrap; padding-left: 20px; padding-right: 50px;">
                        <asp:TextBox ID="tb_employee_id" runat="server" MaxLength="10" Width="100%"></asp:TextBox>
                    </td>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label4" runat="server" CssClass="form" Text="加班總時數限制:" Width="77px"></asp:Label>
                    </td>
                    <td width="15%" class="form" style="height: 25px; white-space: nowrap; padding-left: 20px; padding-right: 55px;">
                        <asp:TextBox ID="tb_hour_limit" runat="server" Width="100%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td colspan="3"></td>
                    <td align="right" width="1%" class="form" style="white-space: nowrap; padding-right: 50px; padding-top: 4px; padding-bottom: 4px;">
                        <asp:ImageButton ID="ibSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />
                        <asp:ImageButton ID="ibApply" runat="server" ImageUrl="~/Image/apply.gif" ToolTip="確認" />
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:GridView ID="gvDataRecords" runat="server" AllowPaging="False" AllowSorting="True" AutoGenerateColumns="False" EmptyDataText="查無任何資料" DataSourceID="sdsDataRecords" Width="100%" DataKeyNames="employee_id" CellPadding="4" ForeColor="#333333" GridLines="None">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EditRowStyle BackColor="#999999" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <Columns>
                <asp:BoundField DataField="emp_chinese_name" HeaderText="人員姓名" SortExpression="emp_chinese_name">
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="employee_id" HeaderText="人員編號" SortExpression="employee_id">
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="加班總時數限制">
                    <HeaderStyle HorizontalAlign="Center" Width="60%" />
                    <ItemStyle HorizontalAlign="Center" />
                    <ItemTemplate>
                        <asp:TextBox ID="gv_tb_hour_limit" runat="server" Text='<%# Bind("HourLimit") %>'></asp:TextBox>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
    </div>

    <asp:SqlDataSource ID="sdsOrganization" runat="server"
        ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
        SelectCommand="Select ORG_NAME, ORG_UID From [ADMINGROUP] Where (1 = 2)">
    </asp:SqlDataSource>

    <asp:SqlDataSource ID="sdsUsersList" runat="server"

        ConnectionString="<%$ ConnectionStrings:ConnectionString %>"

        SelectCommand="Select emp_chinese_name, employee_id From EMPLOYEE Where ORG_UID = @ORG_UID And 
                       AD_TITLE Like (@AD_TITLE + '%') And (1 = 1) Order by emp_chinese_name">

        <SelectParameters>
            <asp:Parameter Name="AD_TITLE" DbType="String" Size="50" />
            <asp:ControlParameter ControlID="ddl_org_picker" Name="ORG_UID" Type="String" PropertyName="SelectedValue" />
        </SelectParameters>

    </asp:SqlDataSource>

    <asp:SqlDataSource ID="sdsDataRecords" runat="server"

        ConnectionString="<%$ ConnectionStrings:ConnectionString %>"

        SelectCommand="Select E.emp_chinese_name, E.employee_id, P.HourLimit From EMPLOYEE As E Left Join P_0103 As P On E.employee_id = P.employee_id 
                       Where E.AD_TITLE Like (@AD_TITLE + '%') And E.ORG_UID In ('') And (1 = 1) Order by E.emp_chinese_name;">

        <SelectParameters>
            <asp:Parameter Name="AD_TITLE" DbType="String" Size="50" />
        </SelectParameters>

    </asp:SqlDataSource>
    </form>
</body>
</html>