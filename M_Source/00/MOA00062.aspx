<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00062.aspx.vb" Inherits="Source_00_MOA00062" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>e關卡人員新增</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
    <div>
        <table  border="0" style="Z-INDEX: 101; LEFT: 158px; WIDTH: 403px; POSITION: absolute; TOP: 22px; HEIGHT: 197px">
            <tr>
                <td style="height: 400px; width: 412px;">
            <asp:Label ID="Label1" runat="server" Text="關卡名稱:" CssClass="form" Width="81px"></asp:Label><asp:DropDownList ID="DropDownList1" runat="server" DataSourceID="SqlDataSource2" DataTextField="object_name" DataValueField="object_uid" Width="200px">
            </asp:DropDownList><asp:RadioButtonList ID="RadioButtonList1" runat="server" RepeatDirection="Horizontal"
                Width="390px" CssClass="form">
                <asp:ListItem Selected="True" Value="1">組織名稱</asp:ListItem>
                <asp:ListItem Value="2">人員名稱</asp:ListItem>
                <asp:ListItem Value="3">人員帳號</asp:ListItem>
            </asp:RadioButtonList><asp:TextBox ID="SearchValue" runat="server" Width="300px"></asp:TextBox>
                    <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" /><br />
        <asp:Panel ID="Panel1" runat="server" Height="50px" Width="361px">
            <asp:ListBox ID="empList" runat="server" DataSourceID="SqlDataSource1" DataTextField="emp_chinese_name"
                DataValueField="employee_id" Height="158px" SelectionMode="Multiple" Width="390px">
            </asp:ListBox><asp:ImageButton ID="btnImgIns" runat="server" ImageUrl="~/Image/add.gif"
                ToolTip="新增" />&nbsp;<asp:ImageButton ID="BackBtn" runat="server" AlternateText="回上頁"
                    ImageUrl="~/Image/backtop.gif" /><br />
            <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                SelectCommand="SELECT ADMINGROUP.ORG_UID, ADMINGROUP.ORG_NAME, EMPLOYEE.employee_id, EMPLOYEE.emp_chinese_name FROM EMPLOYEE INNER JOIN ADMINGROUP ON EMPLOYEE.ORG_UID = ADMINGROUP.ORG_UID">
            </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                SelectCommand="SELECT [object_uid], [object_name], [object_type], [display_order] FROM [SYSTEMOBJ] WHERE ([object_type] <> @object_type) ORDER BY [display_order]">
            <SelectParameters>
                <asp:Parameter DefaultValue="2" Name="object_type" Type="Int32" />
            </SelectParameters>
            </asp:SqlDataSource>
        </asp:Panel>
                </td>
            </tr>
            
        </table>
    
    </div>
    </form>
</body>
</html>
