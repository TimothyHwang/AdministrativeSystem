<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA01008.aspx.vb" Inherits="Source_04_MOA01008" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>文職人員加班管理</title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
    <script type="text/javascript" language="javascript" src="../Inc/inc_common.js"></script>
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
    <div>
        <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center">
                    <asp:Label ID="Label1" runat="server" CssClass="toptitle" Text="文職人員加班管理" Width="100%"></asp:Label>
                </td>
            </tr>
            <tr>
                <td style="padding-top: 10px;">
                    <asp:GridView ID="GridView1" runat="server" AllowPaging="False" AllowSorting="True" AutoGenerateColumns="False" EmptyDataText="查無任何資料" DataSourceID="SqlDataSource1" Width="100%" DataKeyNames="employee_id" CellPadding="4" ForeColor="#333333" GridLines="None">
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <EditRowStyle BackColor="#999999" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
                        <Columns>
                            <asp:TemplateField HeaderText="人員姓名" SortExpression="emp_chinese_name">
                                <HeaderStyle HorizontalAlign="Center" Width="20%" />
                                <ItemStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <%# DataBinder.Eval(Container.DataItem, "emp_chinese_name") %>
                                    <asp:HiddenField ID="hf_employee_id" runat="server" Value='<%# Bind("employee_id") %>' />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="俸額">
                                <HeaderStyle HorizontalAlign="Center" Width="20%" />
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:DropDownList ID="gv_ddl_p_money1" runat="server" Width="60%"></asp:DropDownList>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="專業加給">
                                <HeaderStyle HorizontalAlign="Center" Width="16%" />
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:DropDownList ID="gv_ddl_p_money2" runat="server" Width="60%"></asp:DropDownList>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="主管加給">
                                <HeaderStyle HorizontalAlign="Center" Width="16%" />
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:DropDownList ID="gv_ddl_p_money3" runat="server" Width="60%"></asp:DropDownList>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="俸額級數">
                                <HeaderStyle HorizontalAlign="Center" Width="16%" />
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:DropDownList ID="gv_ddl_money_kind" runat="server" Width="80%"></asp:DropDownList>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="加班總時數限制" Visible="False">
                                <HeaderStyle HorizontalAlign="Center" Width="20%" />
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:TextBox ID="gv_tb_hour_limit" runat="server" Text='<%# Bind("HourLimit") %>'></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td align="center" style="background-color: #5D7B9D;">
                    <asp:ImageButton ID="btnImgOK" runat="server" ImageUrl="~/Image/apply.gif" ToolTip="確認" />
                </td>
            </tr>
        </table>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server"

            ConnectionString="<%$ ConnectionStrings:ConnectionString %>"

            SelectCommand="Select E.emp_chinese_name, E.employee_id, E.ORG_UID, P.P_Money1, P.P_Money2, P.P_Money3, P.MoneyKind, P.HourLimit From 
                          (Select employee_id, emp_chinese_name, ORG_UID From EMPLOYEE Where AD_TITLE Like (@AD_TITLE + '%') And ORG_UID In ('') 
                           And (1 = 1)) As E Left Join P_0103 As P On E.employee_id = P.employee_id Order by E.emp_chinese_name;">

            <SelectParameters>
                <asp:Parameter Name="AD_TITLE" DbType="String" Size="50" />
            </SelectParameters>

        </asp:SqlDataSource>
    </div>
    </form>
</body>
</html>