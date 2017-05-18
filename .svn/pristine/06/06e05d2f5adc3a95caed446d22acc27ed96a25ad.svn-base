<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA07002.aspx.vb" Inherits="Source_07_MOA07002" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>報修查詢</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server" onmousedown="SetPos(event);">     
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="報修查詢" Width="100%"></asp:Label>
            </td></tr>
        </table>
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
            <tr><td noWrap style="width: 70px">
                <asp:Label ID="Label6" runat="server" Text="部門名稱：" CssClass="form"></asp:Label>                    
            </td><td style="width: 320px">
                <asp:DropDownList id="PAUNIT"
                    DataSourceID="SqlDataSource1"
                    DataValueField="ORG_NAME"
                    DataTextField="ORG_NAME"
                    runat="server">
                </asp:DropDownList>
            </td><td noWrap style="width: 70px">
                <asp:Label ID="Label7" runat="server" Text="財產編號：" CssClass="form"></asp:Label>                    
            </td><td style="width: 90px">
                <asp:TextBox ID="nAssetNum" runat="server" Width="70px"></asp:TextBox>
            </td><td noWrap style="width: 70px">
                <asp:Label ID="Label11" runat="server" Text="標籤：" CssClass="form"></asp:Label>                    
            </td><td style="width: 90px">
                <asp:DropDownList id="nLabel"
                    DataSourceID="SqlDataSource4"
                    DataValueField="State_Name"
                    DataTextField="Kind_Name"
                    runat="server">
                </asp:DropDownList>
            </td><td noWrap style="width: 70px">
                <asp:Label ID="Label1" runat="server" Text="標籤號碼：" CssClass="form"></asp:Label>                    
            </td><td style="width: 100px">
                <asp:TextBox ID="nLabelNum" runat="server" Width="70px"></asp:TextBox>
            </td></tr> 
            <tr><td noWrap style="width: 70px">
                <asp:Label ID="Label3" runat="server" Text="申請日期：" CssClass="form"></asp:Label>
            </td><td style="width: 320px">
                <asp:TextBox ID="nAPPTIME1" runat="server" Width="60px" OnKeyDown="return false"></asp:TextBox>&nbsp;<asp:ImageButton
                    ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />
                <asp:Label ID="Label4" runat="server" Text="~" CssClass="form"></asp:Label>
                <asp:TextBox ID="nAPPTIME2" runat="server" Width="60px" OnKeyDown="return false"></asp:TextBox>&nbsp;<asp:ImageButton
                    ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" /></td><td noWrap style="width: 70px">
                <asp:Label ID="Label8" runat="server" Text="財產名稱：" CssClass="form"></asp:Label>                    
            </td><td style="width: 90px">
                <asp:TextBox ID="nAssetName" runat="server" Width="70px"></asp:TextBox>
            </td><td noWrap style="width: 70px">
                <asp:Label ID="Label10" runat="server" Text="問題類別：" CssClass="form"></asp:Label>                    
            </td><td style="width: 90px">
                <asp:DropDownList id="nREASON"
                    DataSourceID="SqlDataSource3"
                    DataValueField="State_Name"
                    DataTextField="Kind_Name"
                    runat="server">
                </asp:DropDownList>
            </td><td colspan=2>
                <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢"/>
                <asp:ImageButton ID="ImgClearAll" runat="server" ImageUrl="~/Image/clearall.gif" ToolTip="清除"/></td></tr>
		</table>
	    <asp:GridView ID="GridView1" runat="server" EmptyDataText="查無任何資料" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" DataKeyNames="EFORMSN" DataSourceID="SqlDataSource2" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
            <Columns>
                <asp:BoundField DataField="P_NUM" HeaderText="流水號" SortExpression="P_NUM">
                    <ItemStyle HorizontalAlign="Center" Width="15%" />
                </asp:BoundField>
                <asp:BoundField DataField="PAUNIT" HeaderText="部門名稱" SortExpression="PAUNIT" >
                    <ItemStyle HorizontalAlign="Center" Width="15%" />
                </asp:BoundField>
                <asp:BoundField DataField="PANAME" HeaderText="姓名" SortExpression="PANAME" >
                    <ItemStyle HorizontalAlign="Center" Width="15%" />
                </asp:BoundField>
                <asp:BoundField DataField="nAPPTIME" HeaderText="申請日期" SortExpression="nAPPTIME" >
                    <ItemStyle HorizontalAlign="Center" Width="15%" />
                </asp:BoundField>
                <asp:BoundField DataField="nTel" HeaderText="軍線電話" SortExpression="nTel" >
                    <ItemStyle HorizontalAlign="Center" Width="15%" />
                </asp:BoundField>
                <asp:BoundField DataField="nSeat" HeaderText="儲位" SortExpression="nSeat" >
                    <ItemStyle HorizontalAlign="Center" Width="15%" />
                </asp:BoundField>
                <asp:CommandField ShowSelectButton="True" ButtonType="Image" SelectText="詳細資料" SelectImageUrl="~/Image/List.gif" >
                    <ItemStyle HorizontalAlign="Center" Width="10%"  />
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
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="">
        </asp:SqlDataSource>
        
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="select '' as ORG_UID,'' as ORG_NAME union SELECT [ORG_NAME], [ORG_NAME] FROM [ADMINGROUP] ORDER BY [ORG_NAME]">
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="select '' as State_Name,'' as Kind_Name union SELECT [State_Name], [Kind_Name]+'-'+[State_Name] Kind_Name FROM [SYSKIND] WHERE Kind_Num in ('7','8','9','10') ORDER BY [Kind_Name]">
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource4" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="select '' as State_Name,'' as Kind_Name union SELECT [State_Name], [Kind_Name]+'-'+[State_Name] Kind_Name FROM [SYSKIND] WHERE Kind_Num ='11' ORDER BY [Kind_Name]">
        </asp:SqlDataSource>
        
        <div id="Div_grid" runat="server" style="border-right: lightslategray 2px solid;
            border-top: lightslategray 2px solid; display: block; z-index: 3; left: 191px;
            border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
            position: absolute; top: 458px; height: 150pt; background-color: white" visible="false">
            <asp:Calendar ID="Calendar1" runat="server" BackColor="White" BorderColor="#3366CC"
                BorderWidth="1px" Caption="" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana"
                Font-Size="8pt" ForeColor="#003399" Height="200px" ShowGridLines="True" Width="220px">
                <SelectedDayStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
                <TodayDayStyle BackColor="#99CCCC" ForeColor="White" />
                <SelectorStyle BackColor="#99CCCC" ForeColor="#336666" />
                <WeekendDayStyle BackColor="#CCCCFF" />
                <OtherMonthDayStyle ForeColor="#999999" />
                <NextPrevStyle Font-Size="8pt" ForeColor="#CCCCFF" />
                <DayHeaderStyle BackColor="#99CCCC" ForeColor="#336666" Height="1px" />
                <TitleStyle BackColor="#003399" BorderColor="#3366CC" BorderWidth="1px" Font-Bold="True"
                    Font-Size="10pt" ForeColor="#CCCCFF" Height="25px" />
            </asp:Calendar>
            <asp:Button ID="btnClose1" runat="server" Text="關閉" Width="220px" /></div>
        <div id="Div_grid2" runat="server" style="border-right: lightslategray 2px solid;
            border-top: lightslategray 2px solid; display: block; z-index: 3; left: 428px;
            border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
            position: absolute; top: 460px; height: 150pt; background-color: white" visible="false">
            <asp:Calendar ID="Calendar2" runat="server" BackColor="White" BorderColor="#3366CC"
                BorderWidth="1px" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana"
                Font-Size="8pt" ForeColor="#003399" Height="200px" ShowGridLines="True" Width="220px">
                <SelectedDayStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
                <TodayDayStyle BackColor="#99CCCC" ForeColor="White" />
                <SelectorStyle BackColor="#99CCCC" ForeColor="#336666" />
                <WeekendDayStyle BackColor="#CCCCFF" />
                <OtherMonthDayStyle ForeColor="#999999" />
                <NextPrevStyle Font-Size="8pt" ForeColor="#CCCCFF" />
                <DayHeaderStyle BackColor="#99CCCC" ForeColor="#336666" Height="1px" />
                <TitleStyle BackColor="#003399" BorderColor="#3366CC" BorderWidth="1px" Font-Bold="True"
                    Font-Size="10pt" ForeColor="#CCCCFF" Height="25px" />
            </asp:Calendar>
            <asp:Button ID="btnClose2" runat="server" Text="關閉" Width="220px" /></div>
    </form>
</body>
</html>
