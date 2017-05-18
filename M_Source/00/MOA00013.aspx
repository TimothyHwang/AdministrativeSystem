<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00013.aspx.vb" Inherits="Source_00_MOA00013" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>選擇上一級主管</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">    
        <table border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr>
            <td align="center">
                    <asp:Label ID="LabTitle" runat="server" CssClass="toptitle" Text="選擇上一級主管" Width="100%"></asp:Label>
            </td>
            </tr>
        </table>
        <table width="100%" border="3" bordercolor="#ccddee" align="center">
            <tr>
            <td align="center" class="CellClass" style="width: 25%">
                <asp:Label ID="LabName" runat="server" Text="上一級主管：" ></asp:Label>
            </td>
            <td style="width: 75%">
                <asp:DropDownList ID="DDLUser" runat="server" DataSourceID="SqlDataSource1" DataTextField="emp_chinese_name"
                    DataValueField="employee_id">
                </asp:DropDownList>
            </td>
            </tr>
            <tr>
            <td colspan="2" align="center">
                <asp:Button ID="btnSend" runat="server" Text="送件" /></td>
            </tr>
        </table>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="select employee_id,emp_chinese_name from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = (select ORG_UID from EMPLOYEE where employee_id = @user_id))">
            <SelectParameters>
                <asp:SessionParameter Name="user_id" SessionField="user_id" />
            </SelectParameters>
        </asp:SqlDataSource>
    </form>
</body>
</html>
