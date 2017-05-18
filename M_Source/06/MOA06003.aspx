<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA06003.aspx.vb" Inherits="Source_06_MOA06003" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>資訊設備媒體攜出入查詢</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="left" style="width:3%" >
                <asp:ImageButton id="ImgPrint" runat="server" ImageUrl="../../Image/print.gif" ToolTip="列印"></asp:ImageButton>
            </td><td align="left" style="width:3%" >
                <asp:ImageButton ID="BackBtn" runat="server" AlternateText="回上頁" ImageUrl="~/Image/backtop.gif" ToolTip="回上頁" />
            </td><td align="center" style="width:94%" >
                    <asp:Label ID="Label2" runat="server" CssClass="toptitle" Text="資訊設備媒體攜出入查詢" Width="100%"></asp:Label>
            </td></tr>
        </table>
        <table width="100%">
            <tr><td>
                <asp:DetailsView ID="DetailsView1" runat="server" DataSourceID="SqlDataSource1"
                    Height="50px" Width="100%" AutoGenerateRows="False">
                    <Fields>
                        <asp:BoundField DataField="PWUNIT" HeaderText="填表人單位" ReadOnly="True" SortExpression="PWUNIT" />
                        <asp:BoundField DataField="PWNAME" HeaderText="填表人姓名" ReadOnly="True" SortExpression="PWNAME" />
                        <asp:BoundField DataField="PWTITLE" HeaderText="填表人級職" ReadOnly="True" SortExpression="PWTITLE" />
                        <asp:BoundField DataField="PAUNIT" HeaderText="申請人單位" ReadOnly="True" SortExpression="PAUNIT" />
                        <asp:BoundField DataField="PANAME" HeaderText="申請人姓名" ReadOnly="True" SortExpression="PANAME" />
                        <asp:BoundField DataField="PATITLE" HeaderText="申請人級職" ReadOnly="True" SortExpression="PATITLE" />
                        <asp:BoundField DataField="nAPPTIME" HeaderText="申請時間" ReadOnly="True" SortExpression="nAPPTIME" />
                        <asp:BoundField DataField="nREASON" HeaderText="申請事由" ReadOnly="True" SortExpression="nREASON" />
                        <asp:BoundField DataField="nDATE" HeaderText="申請出入日期" ReadOnly="True" SortExpression="nDATE" />
                        <asp:BoundField DataField="nPLACE" HeaderText="使用地點" ReadOnly="True" SortExpression="nPLACE" />
                    </Fields>
                </asp:DetailsView>
                <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True" AutoGenerateColumns="False"
                    DataSourceID="SqlDataSource2" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
                    <Columns>
                        <asp:BoundField DataField="nKind" HeaderText="區分" ReadOnly="True" SortExpression="nKind" >
                            <ItemStyle HorizontalAlign="Center" />
                            <HeaderStyle HorizontalAlign="Center" Width="15%" />
                        </asp:BoundField>
                        <asp:BoundField DataField="nMName" HeaderText="名稱機型" ReadOnly="True" SortExpression="nMName" >
                            <HeaderStyle HorizontalAlign="Center" Width="20%" />
                        </asp:BoundField>
                        <asp:BoundField DataField="nDocnum" HeaderText="編(文)號" ReadOnly="True" SortExpression="nDocnum" >
                            <HeaderStyle HorizontalAlign="Center" Width="15%" />
                        </asp:BoundField>
                        <asp:BoundField DataField="nAmount" HeaderText="數量" ReadOnly="True" SortExpression="nAmount" >
                            <ItemStyle HorizontalAlign="Right" />
                            <HeaderStyle HorizontalAlign="Center" Width="10%" />
                        </asp:BoundField>
                        <asp:BoundField DataField="nContent" HeaderText="內容摘要" ReadOnly="True" SortExpression="nContent" >
                            <HeaderStyle HorizontalAlign="Center" Width="30%" />
                        </asp:BoundField>
                        <asp:BoundField DataField="nClass" HeaderText="機密等級" ReadOnly="True" SortExpression="nClass" >
                            <HeaderStyle HorizontalAlign="Center" Width="10%" />
                        </asp:BoundField>
                    </Columns>
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <EditRowStyle BackColor="#999999" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                </asp:GridView>
            </td></tr>
        </table>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand=""></asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand=""></asp:SqlDataSource>
    </form>
</body>
</html>
