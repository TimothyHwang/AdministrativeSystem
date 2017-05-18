<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA05005.aspx.vb" Inherits="Source_05_MOA05005" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>會客洽公統計</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <link href="../../css/jquery.datepick.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.10.2.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick-zh-TW.js" type="text/javascript"></script>
    <script type="text/javascript" language="javascript">
        $(function () {
            $("#nAPPTIME1").datepick({ formats: 'yyyy/m/d', defaultDate: $("#nAPPTIME1").val(), showTrigger: '#calImg' });
            $("#nAPPTIME2").datepick({ formats: 'yyyy/m/d', defaultDate: $("#nAPPTIME2").val(), showTrigger: '#calImg' });
        }); 
    </script>
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
    <table border="0" width="100%" style="z-index: 101; left: 105px; top: 33px;">
        <tr>
            <td align="center">
                <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="會客洽公統計" Width="100%"></asp:Label>
            </td>
        </tr>
    </table>
    <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
        <tr>
            <td nowrap align='right'>
                <asp:Label ID="Label3" runat="server" Text="日期：" CssClass="form"></asp:Label>
            </td>
            <td>
                <asp:TextBox ID="nAPPTIME1" runat="server" Width="60px" OnKeyDown="return false"></asp:TextBox>
                <div style="display: none;">
                    <img id="calImg" src="../../Image/calendar.gif" alt="選擇日期" />
                </div>
                <asp:Label ID="Label4" runat="server" Text="~" CssClass="form"></asp:Label>
                <asp:TextBox ID="nAPPTIME2" runat="server" Width="60px" OnKeyDown="return false"></asp:TextBox>
            </td>
            <td>
                <asp:Label ID="Label1" runat="server" Text="部門名稱：" CssClass="form"></asp:Label>
                <asp:DropDownList ID="PAUNIT" DataSourceID="SqlDataSource1" DataValueField="ORG_NAME"
                    DataTextField="ORG_NAME" runat="server" AutoPostBack="True">
                </asp:DropDownList>
            </td>
            <td>
                <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" title="查詢"
                    ToolTip="查詢" />
                <asp:ImageButton ID="ImgClearAll" runat="server" ImageUrl="~/Image/clearall.gif"
                    title="畫面清除" ToolTip="清除" />
            </td>
        </tr>
    </table>
    <asp:GridView ID="GridView1" runat="server" EmptyDataText="查無任何資料" AutoGenerateColumns="False"
        AllowPaging="True" AllowSorting="True" DataSourceID="SqlDataSource2" Width="100%"
        CellPadding="4" ForeColor="#333333" GridLines="None">
        <Columns>
            <asp:BoundField DataField="PAUNIT" HeaderText="部門名稱" SortExpression="PAUNIT">
                <HeaderStyle HorizontalAlign="Center" Width="50%" />
            </asp:BoundField>
            <asp:BoundField DataField="count" HeaderText="來賓人數" SortExpression="count">
                <ItemStyle HorizontalAlign="Right" />
                <HeaderStyle HorizontalAlign="Center" Width="20%" />
            </asp:BoundField>
            <asp:BoundField DataField="nRECROOM" HeaderText="會客室" SortExpression="nRECROOM">
                <HeaderStyle HorizontalAlign="Center" Width="30%" />
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
    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
        SelectCommand="select '' as ORG_UID,'' as ORG_NAME union SELECT [ORG_NAME], [ORG_NAME] FROM [ADMINGROUP] ORDER BY [ORG_NAME]">
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
        SelectCommand=""></asp:SqlDataSource>
    </form>
</body>
</html>
