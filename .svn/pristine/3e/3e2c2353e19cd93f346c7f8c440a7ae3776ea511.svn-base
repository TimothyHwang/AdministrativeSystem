<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04011_BK.aspx.vb" Inherits="Source_04_MOA04011_BK" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>設備編碼新增</title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
    <script type="text/javascript" language="javascript" src="../../Inc/inc_common.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center">
                    <asp:Label ID="Label1" runat="server" CssClass="toptitle" Text="設備編碼新增" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>
        <div style="width: 100; text-align: center; padding-top: 10px;">
            <table cellspacing="0" cellpadding="0" width="60%" align="center" border="0" style="text-align: left;">
                <tr>
                    <td nowrap class="form" style="height: 25px; width: 10%;">
                        <asp:Label ID="Label2" runat="server" CssClass="form" Text="預算代碼:" Width="77px"></asp:Label>
                    </td>
                    <td nowrap width="15%" class="form" style="height: 25px">
                        <asp:TextBox ID="tb_bd_code" runat="server" MaxLength="1" Width="90%"></asp:TextBox>
                    </td>
                    <td nowrap class="form" style="height: 25px; width: 10%;">
                        <asp:Label ID="Label3" runat="server" CssClass="form" Text="樓層代碼:" Width="77px"></asp:Label>
                    </td>
                    <td nowrap width="15%" class="form" style="height: 25px">
                        <asp:TextBox ID="tb_fl_code" runat="server" MaxLength="2" Width="90%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td nowrap class="form" style="height: 25px; width: 10%;">
                        <asp:Label ID="Label4" runat="server" CssClass="form" Text="房間地理位置代碼:" Width="77px"></asp:Label>
                    </td>
                    <td nowrap width="15%" class="form" style="height: 25px">
                        <asp:TextBox ID="tb_rnum_code" runat="server" MaxLength="5" Width="90%"></asp:TextBox>
                    </td>
                    <td nowrap class="form" style="height: 25px; width: 10%;">
                        <asp:Label ID="Label5" runat="server" CssClass="form" Text="房間規格代碼:" Width="77px"></asp:Label>
                    </td>
                    <td nowrap width="15%" class="form" style="height: 25px">
                        <asp:TextBox ID="tb_wa_code" runat="server" MaxLength="1" Width="90%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td nowrap class="form" style="height: 25px; width: 10%;">
                        <asp:Label ID="Label6" runat="server" CssClass="form" Text="建物代碼:" Width="77px"></asp:Label>
                    </td>
                    <td nowrap width="15%" class="form" style="height: 25px">
                        <asp:TextBox ID="tb_bg_code" runat="server" MaxLength="2" Width="90%"></asp:TextBox>
                    </td>
                    <td nowrap class="form" style="height: 25px; width: 10%;">
                        <asp:Label ID="Label7" runat="server" CssClass="form" Text="物料分類代碼:" Width="77px"></asp:Label>
                    </td>
                    <td nowrap width="15%" class="form" style="height: 25px">
                        <asp:TextBox ID="tb_it_code" runat="server" MaxLength="6" Width="90%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td nowrap class="form" style="height: 25px; width: 10%;">
                        <asp:Label ID="Label8" runat="server" CssClass="form" Text="設備編碼流水號:" Width="77px"></asp:Label>
                    </td>
                    <td nowrap width="15%" class="form" style="height: 25px">
                        <asp:TextBox ID="tb_element_no" runat="server" MaxLength="3" Width="90%"></asp:TextBox>
                    </td>
                    <td colspan="2">&nbsp;</td>
                </tr>
                <tr>
                    <td nowrap class="form" style="height: 25px; width: 10%;">
                        <asp:Label ID="Label9" runat="server" CssClass="form" Text="人員帳號:" Width="77px"></asp:Label>
                    </td>
                    <td nowrap width="15%" class="form" style="height: 25px">
                        <asp:TextBox ID="tb_operator" runat="server" MaxLength="10" Width="90%"></asp:TextBox>
                    </td>
                    <td colspan="2" align="left">
                        <asp:ImageButton ID="ibtnSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />
                    </td>
                </tr>
                <tr>
                    <td nowrap class="form" style="height: 25px; width: 10%;">&nbsp;</td>
                    <td colspan="3" style="padding-top: 4px;">
                        <asp:ListBox ID="liEmployee" runat="server" DataSourceID="SqlDataSource1" DataTextField="emp_chinese_name" DataValueField="employee_id" Height="158px" SelectionMode="Single" Width="97%"></asp:ListBox>

                        <asp:SqlDataSource ID="SqlDataSource1" runat="server"
                                ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                                SelectCommand="Select ADMINGROUP.ORG_UID, ADMINGROUP.ORG_NAME, EMPLOYEE.employee_id, EMPLOYEE.emp_chinese_name From EMPLOYEE Inner Join ADMINGROUP ON EMPLOYEE.ORG_UID = ADMINGROUP.ORG_UID And 1 = 2">
                        </asp:SqlDataSource>
                    </td>
                </tr>
                <tr>
                    <td colspan="3">&nbsp;</td>
                    <td>
                        <div style="width: 92%; padding-top: 5px; text-align: right;">
                            <asp:ImageButton ID="ibtnAdd" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" />
                            <asp:ImageButton ID="ibtnPrevious" runat="server" AlternateText="回上頁" ImageUrl="~/Image/backtop.gif" />
                        </div>
                    </td>
                </tr>
            </table>
        </div>
    </div>
    </form>
</body>
</html>