<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA07003.aspx.vb" Inherits="Source_07_MOA07003" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>報修詳細資料</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="left" style="width:3%" >
                <a href="javascript:window.history.go(<%= HistoryGo %>);">
                <img  align='absmiddle' title="回上頁" border="0" src="../../Image/backtop.gif" class="btn_image"/></a>
            </td><td align="left" style="width:97%" >
                    <asp:Label ID="Label2" runat="server" CssClass="toptitle" Text="報修詳細資料" Width="100%"></asp:Label>
            </td></tr>
        </table>
        <table width="100%">
            <tr><td>
                <asp:DetailsView ID="DetailsView1" runat="server" DataSourceID="SqlDataSource1"
                    Height="50px" Width="100%" AutoGenerateRows="False">
                    <Fields>
                        <asp:BoundField DataField="PWUNIT" HeaderText="填表人單位" ReadOnly="True" SortExpression="PWUNIT" />
                        <asp:BoundField DataField="PWTITLE" HeaderText="填表人級職" ReadOnly="True" SortExpression="PWTITLE" />
                        <asp:BoundField DataField="PWNAME" HeaderText="填表人姓名" ReadOnly="True" SortExpression="PWNAME" />
                        <asp:BoundField DataField="PWIDNO" HeaderText="填表人帳號" ReadOnly="True" SortExpression="PWIDNO" />
                        <asp:BoundField DataField="PAUNIT" HeaderText="申請人單位" ReadOnly="True" SortExpression="PAUNIT" />
                        <asp:BoundField DataField="PANAME" HeaderText="申請人姓名" ReadOnly="True" SortExpression="PANAME" />
                        <asp:BoundField DataField="PATITLE" HeaderText="申請人級職" ReadOnly="True" SortExpression="PATITLE" />
                        <asp:BoundField DataField="PAIDNO" HeaderText="申請人帳號" ReadOnly="True" SortExpression="PAIDNO" />
                        <asp:BoundField DataField="nAPPTIME" HeaderText="申請時間" ReadOnly="True" SortExpression="nAPPTIME" />
                        <asp:BoundField DataField="nTel" HeaderText="軍線電話" ReadOnly="True" SortExpression="nTel" />
                        <asp:BoundField DataField="nSeat" HeaderText="儲位" ReadOnly="True" SortExpression="nSeat" />
                        <asp:BoundField DataField="PENDFLAG" HeaderText="流程結束" ReadOnly="True" SortExpression="PENDFLAG" />
                    </Fields>
                </asp:DetailsView>
                <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True" AutoGenerateColumns="False"
                    DataSourceID="SqlDataSource2" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
                    <Columns>
                        <asp:BoundField DataField="nAssetNum" HeaderText="財產編號" ReadOnly="True" SortExpression="nAssetNum" >
                            <ItemStyle HorizontalAlign="Center" Width="15%" />
                        </asp:BoundField>
                        <asp:BoundField DataField="nAssetName" HeaderText="財產名稱" ReadOnly="True" SortExpression="nAssetName" >
                            <ItemStyle HorizontalAlign="Center" Width="15%" />
                        </asp:BoundField>
                        <asp:BoundField DataField="nBios" HeaderText="Bios密碼" ReadOnly="True" SortExpression="nBios" >
                            <ItemStyle HorizontalAlign="Center" Width="10%" />
                        </asp:BoundField>
                        <asp:BoundField DataField="nLabel" HeaderText="標籤" ReadOnly="True" SortExpression="nLabel" >
                            <ItemStyle HorizontalAlign="Center" Width="15%" />
                        </asp:BoundField>
                        <asp:BoundField DataField="nLabelNum" HeaderText="標籤號碼" ReadOnly="True" SortExpression="nLabelNum" >
                            <ItemStyle HorizontalAlign="Center" Width="10%" />
                        </asp:BoundField>
                        <asp:BoundField DataField="nAmount" HeaderText="數量" ReadOnly="True" SortExpression="nAmount" >
                            <ItemStyle HorizontalAlign="Center" Width="10%" />
                        </asp:BoundField>
                        <asp:BoundField DataField="nREASON" HeaderText="問題類別" ReadOnly="True" SortExpression="nREASON" >
                            <ItemStyle HorizontalAlign="Center" Width="10%" />
                        </asp:BoundField>
                        <asp:BoundField DataField="nContent" HeaderText="問題描述" ReadOnly="True" SortExpression="nContent" >
                            <ItemStyle HorizontalAlign="Center" Width="15%" />
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
