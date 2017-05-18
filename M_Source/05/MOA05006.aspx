<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA05006.aspx.vb" Inherits="M_Source_05_MOA05006" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>會客洽公人員統計表</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <link href="../../css/jquery.datepick.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.10.2.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick-zh-TW.js" type="text/javascript"></script>
    <script type="text/javascript" language="javascript">
        $(function () {
            $("#Sdate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#Sdate").val(), showTrigger: '#calImg' });
            $("#Edate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#Edate").val(), showTrigger: '#calImg' });
        }); 
    </script>
</head>
<body language="javascript" onload="return window_onload()">
    <form id="form1" runat="server">
    <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
        <tr>
            <td align="center">
                <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="會客洽公人員統計表" Width="100%"></asp:Label>
            </td>
        </tr>
    </table>
    <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
        <tr>
            <td nowrap style="height: 25px; width: 5%;">
                <asp:Label ID="Label2" runat="server" Text="日期：" CssClass="form"></asp:Label>
            </td>
            <td nowrap style="height: 25px; width: 45%;">
                <asp:TextBox ID="Sdate" runat="server" Width="80px" OnKeyDown="return false"></asp:TextBox>
                ~
                <asp:TextBox ID="Edate" runat="server" Width="80px" OnKeyDown="return false"></asp:TextBox>
                <div style="display: none;">
                    <img id="calImg" src="../../Image/calendar.gif" alt="選擇日期" />
                </div>
            </td>
            <td nowrap style="height: 25px; width: 10%;">
                <asp:Label ID="Label3" runat="server" CssClass="form" Text="會客室："></asp:Label>
            </td>
            <td nowrap style="height: 25px; width: 20%;">
                <asp:DropDownList ID="DrDown_nRECROOM" runat="server" CssClass="form" ForeColor="Black"
                    Width="90px" DataSourceID="SqlDataSource1" DataTextField="State_Name" DataValueField="State_Name">
                </asp:DropDownList>
            </td>
            <td nowrap style="height: 25px; width: 20%;" align="center">
                <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />&nbsp;<asp:ImageButton
                    ID="Img_Export" runat="server" ImageUrl="~/Image/ExportFile.gif" />
                <asp:ImageButton ID="ImagePrint" runat="server" ImageUrl="~/Image/print.gif" ToolTip="列印" />
            </td>
        </tr>
    </table>
    <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
        AutoGenerateColumns="False" CellPadding="4" DataSourceID="SqlDataSource2" EmptyDataText="查無任何資料"
        ForeColor="#333333" GridLines="None" PageSize="31" Width="100%">
        <Columns>
            <asp:BoundField DataField="nRECDATE" HeaderText="日期" SortExpression="nRECDATE">
                <HeaderStyle HorizontalAlign="Center" Width="15%" />
                <ItemStyle HorizontalAlign="Center" />
            </asp:BoundField>
            <asp:BoundField DataField="DWEEK" HeaderText="星期" SortExpression="DWEEK">
                <HeaderStyle HorizontalAlign="Center" Width="10%" />
                <ItemStyle HorizontalAlign="Center" />
            </asp:BoundField>
            <asp:BoundField DataField="PN01" HeaderText="件數" NullDisplayText="件數" SortExpression="PN01">
                <HeaderStyle HorizontalAlign="Center" Width="15%" />
                <ItemStyle HorizontalAlign="Center" />
            </asp:BoundField>
            <asp:BoundField DataField="PN03" HeaderText="軍人" SortExpression="PN03">
                <HeaderStyle HorizontalAlign="Center" Width="20%" />
                <ItemStyle HorizontalAlign="Center" />
            </asp:BoundField>
            <asp:BoundField DataField="PN02" HeaderText="民人" SortExpression="PN02">
                <HeaderStyle HorizontalAlign="Center" Width="20%" />
                <ItemStyle HorizontalAlign="Center" />
            </asp:BoundField>
            <asp:BoundField DataField="PN04" HeaderText="合計" SortExpression="PN04">
                <HeaderStyle HorizontalAlign="Center" Width="20%" />
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
    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
        SelectCommand="SELECT * FROM T_SIGN_RECORD,EMPLOYEE WHERE T_SIGN_RECORD.LOGONID_nvc = EMPLOYEE.employee_id AND 1=2">
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
        SelectCommand="select State_Name from SYSKIND where Kind_Num =3 and State_Enabled='1'"></asp:SqlDataSource>
    </form>
    <iframe id="lst" frameborder="0" width="0" height="0" src="/blank.htm"></iframe>
    <script language="javascript">
        var errmsg = '<%= do_sql.G_errmsg%>';
        var print_file = '<%= print_file%>';

        function window_onload() {
            if (errmsg != '') {
                alert(errmsg);
            }

            if (print_file != '') {
                lst.navigate(print_file);
            }
        }

    </script>
</body>
</html>
