<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00042.aspx.vb" Inherits="Source_00_MOA00042" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>人員待派確認</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" language="javascript">
            document.oncontextmenu = new Function("return false");
    </script>
</head>
<body>
    <form id="form1" runat="server">      
        <table border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr>
            <td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="人員待派確認" Width="100%"></asp:Label>
            </td>
            </tr>
        </table>
        <table width="100%" border="3" bordercolor="#ccddee" align="center">
            <tr>
            <td align="center" class="CellClass" style="width: 25%">
                <asp:Label ID="Label1" runat="server" Text="單位：" ></asp:Label>
            </td>
            <td style="width: 75%">
                <asp:DropDownList ID="OrgSel" runat="server" AutoPostBack="True" DataSourceID="SqlDataSource1"
                    DataTextField="ORG_NAME" DataValueField="ORG_UID">
                </asp:DropDownList></td>
            </tr>
            
            <tr>
            <td align="center" class="CellClass" style="width: 25%">
                <asp:Label ID="Label3" runat="server" Text="表單分派人員：" ></asp:Label>
            </td>
            <td style="width: 75%">
                <asp:DropDownList ID="UserSel" runat="server" DataSourceID="SqlDataSource3" DataTextField="emp_chinese_name"
                    DataValueField="employee_id">
                </asp:DropDownList>
                <asp:Label ID="ErrName" runat="server" ForeColor="Red" Text="請選擇姓名" Visible="False"
                    Width="100px"></asp:Label></td>
            </tr>
        </table>
        <table width="100%" align="center">
            <tr>
            <td align="center">
                <asp:ImageButton ID="ImgApply" runat="server" OnClientClick="return confirm('確定將將人員移到待派區並且重新分派表單嗎?')" ImageUrl="~/Image/apply.gif" />
                <asp:ImageButton ID="BackBtn" runat="server" AlternateText="回上頁" ImageUrl="~/Image/backtop.gif" />
            </tr>
        </table>        
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] ORDER BY [ORG_NAME]">
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [employee_id], [emp_chinese_name], [ORG_UID] FROM [EMPLOYEE] WHERE ([ORG_UID] = @ORG_UID) ORDER BY [emp_chinese_name]">
            <SelectParameters>
                <asp:ControlParameter ControlID="OrgSel" Name="ORG_UID" PropertyName="SelectedValue"
                    Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
    </form>
</body>
</html>
