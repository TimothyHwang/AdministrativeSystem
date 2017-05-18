<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA05003.aspx.vb" Inherits="Source_05_MOA05003" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>憲兵換證作業</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <link href="../../css/jquery.datepick.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.10.2.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick-zh-TW.js" type="text/javascript"></script>
    <script type="text/javascript" language="javascript">
        $(function () {
            $("#nRECDATE").datepick({ formats: 'yyyy/m/d', defaultDate: $("#nRECDATE").val(), showTrigger: '#calImg' });
        });                
    </script>
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server" >     
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="憲兵換證作業" Width="100%"></asp:Label>
            </td></tr>
        </table>
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
            <tr><td noWrap style="width: 40px">
                <asp:Label ID="Label3" runat="server" Text="日期：" CssClass="form"></asp:Label>
            </td><td style="width: 90px">
                <asp:TextBox ID="nRECDATE" runat="server" Width="60px" OnKeyDown="return false"></asp:TextBox>
            <div style="display: none;">
                <img id="calImg" src="../../Image/calendar.gif" alt="選擇日期" />
            </div>
            </td><td noWrap style="width: 70px">
                <asp:Label ID="Label5" runat="server" Text="來賓姓名：" CssClass="form"></asp:Label></td><td style="width: 70px">
                <asp:TextBox ID="nName" runat="server" Width="60px"></asp:TextBox>
            </td><td noWrap style="width: 110px">
                <asp:Label ID="Label1" runat="server" Text="來賓身分證號碼：" CssClass="form"></asp:Label>                    
            </td><td style="width: 70px">
                <asp:TextBox ID="nID" runat="server" Width="60px"></asp:TextBox>
            </td><td noWrap style="width: 60px">
                <asp:Label ID="Label6" runat="server" Text="被會人：" CssClass="form"></asp:Label></td><td style="width: 90px">
                <asp:TextBox ID="PANAME" runat="server" Width="60px"></asp:TextBox></td>
            <td style="width: 70px">
                <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" title="查詢" ToolTip="查詢"/></td>
            </tr>
            <tr><td noWrap style="width: 40px">
                <asp:Label ID="Label7" runat="server" Text="電話：" CssClass="form"></asp:Label></td><td style="width: 90px">
                <asp:TextBox ID="nPHONE" runat="server" Width="70px"></asp:TextBox></td><td noWrap style="width: 70px">
                <asp:Label ID="Label4" runat="server" Text="部門名稱：" CssClass="form"></asp:Label></td><td style="width: 70px">
                <asp:TextBox ID="PAUNIT" runat="server" Width="60px"></asp:TextBox></td><td noWrap style="width: 110px">
                <asp:Label ID="Label2" runat="server" Text="會客入口：" CssClass="form"></asp:Label></td><td style="width: 70px">
                    <asp:DropDownList ID="DDLnRECEXIT" runat="server" DataSourceID="SqlDataSource2" DataTextField="State_Name"
                        DataValueField="State_Name">
                    </asp:DropDownList></td><td noWrap style="width: 60px">
                <asp:Label ID="Label10" runat="server" Text="狀態：" CssClass="form"></asp:Label>                    
            </td><td style="width: 90px">
                <asp:DropDownList id="nSTATUS" runat="server">
                    <asp:ListItem Value="0">尚未進營</asp:ListItem>
                    <asp:ListItem Value="1">正在營中</asp:ListItem>
                    <asp:ListItem Value="2">已經出營</asp:ListItem>
                </asp:DropDownList></td><td style="width: 70px">
                <asp:ImageButton ID="ImgClearAll" runat="server" ImageUrl="~/Image/clearall.gif" title="畫面清除" ToolTip="清除"/></td></tr>
		</table>
	    <asp:GridView ID="GridView1" runat="server" EmptyDataText="查無任何資料" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" DataKeyNames="EFORMSN" DataSourceID="SqlDataSource1" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
            <Columns>
                <asp:BoundField DataField="nRECDATE" HeaderText="會客日期" SortExpression="nRECDATE" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="nREASON" HeaderText="事由" SortExpression="nREASON" >
                    <HeaderStyle HorizontalAlign="Center" Width="22%" />
                </asp:BoundField>
                <asp:BoundField DataField="PAUNIT" HeaderText="部門名稱" SortExpression="PAUNIT" >
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:BoundField DataField="PANAME" HeaderText="被會人" SortExpression="PANAME" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                </asp:BoundField>
                <asp:BoundField DataField="nPHONE" HeaderText="電話" SortExpression="nPHONE" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                </asp:BoundField>
                <asp:BoundField DataField="nRECEXIT" HeaderText="會客入口" SortExpression="nRECEXIT" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="12%" />
                </asp:BoundField>
                <asp:templatefield HeaderText="狀態"  SortExpression="nSTATUS">
                    <itemtemplate>
                        <asp:label id="LastNameLabel"
                        text= '<%# Chg_nSTATUS("nSTATUS") %>'
                        runat="server"/>
                    </itemtemplate>
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:templatefield>
                <asp:CommandField ShowSelectButton="True" ButtonType="Image" SelectText="詳細資料" SelectImageUrl="~/Image/List.gif" >
                    <ItemStyle HorizontalAlign="Center"  />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:CommandField>
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
            SelectCommand="">
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="select State_Name from SYSKIND where Kind_Num =4"></asp:SqlDataSource>

    </form>
</body>
</html>
