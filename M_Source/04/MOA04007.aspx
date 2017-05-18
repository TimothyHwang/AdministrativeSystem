<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04007.aspx.vb" Inherits="M_Source_04_MOA04007" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>房屋水電修繕詳細資料</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
       <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="房屋水電修繕詳細資料" Width="100%"></asp:Label>
            </td></tr>
        </table> 
       <asp:GridView ID="GridView1" runat="server" EmptyDataText="查無任何資料" AllowPaging="True" AllowSorting="True"
            AutoGenerateColumns="False" CellPadding="4" DataSourceID="SqlDataSource2" ForeColor="#333333"
            GridLines="None" Width="100%">
            <Columns>
                <asp:BoundField DataField="PAUNIT" HeaderText="部門名稱" SortExpression="PAUNIT">
                    <HeaderStyle HorizontalAlign="Center" Width="30%" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="PANAME" HeaderText="申請人" SortExpression="PANAME" >
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="nPHONE" HeaderText="電話" SortExpression="nPHONE" >
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nPLACE" HeaderText="請修地點" SortExpression="nPLACE">
                    <ItemStyle HorizontalAlign="Left" />
                    <HeaderStyle Width="20%" />
                </asp:BoundField>
                <asp:BoundField DataField="nFIXITEM" HeaderText="請修事項" SortExpression="nFIXITEM" >
                    <HeaderStyle Width="20%" />
                </asp:BoundField>
                <asp:BoundField DataField="PENDFLAG" HeaderText="完成" SortExpression="PENDFLAG">
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
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
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr>
            <td align="center">
                <asp:ImageButton ID="BackBtn" runat="server" ImageUrl="~/Image/backtop.gif" AlternateText="回上頁" />
            </td>
            </tr>
        </table>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT * FROM T_SIGN_RECORD,EMPLOYEE WHERE T_SIGN_RECORD.LOGONID_nvc = EMPLOYEE.employee_id AND 1=2"></asp:SqlDataSource>
        
    </form>
</body>
</html>
