<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA06002.aspx.vb" Inherits="Source_06_MOA06002" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>資訊設備媒體攜出入查詢</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server" >     
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="資訊設備媒體攜出入查詢" Width="100%"></asp:Label>
            </td></tr>
        </table>
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
            <tr><td noWrap style="height: 43px; width: 70px;">
                <asp:Label ID="Label1" runat="server" Text="申請日期：" CssClass="form"></asp:Label>
            </td><td noWrap style="width: 200px; height: 43px;">
                <asp:TextBox ID="nAPPTIME1" runat="server" Width="60px" OnKeyDown="return false"></asp:TextBox>
                <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />
                <asp:Label ID="Label2" runat="server" Text="~" CssClass="form"></asp:Label>
                <asp:TextBox ID="nAPPTIME2" runat="server" Width="60px" OnKeyDown="return false"></asp:TextBox>
                <asp:ImageButton ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" /></td><td noWrap style="height: 43px; width: 55px;">
                <asp:Label ID="Label5" runat="server" Text="申請人：" CssClass="form"></asp:Label>                    
            </td><td style="height: 43px; width: 75px;">
                <asp:TextBox ID="PANAME" runat="server" Width="60px"></asp:TextBox>
            </td><td noWrap style="height: 43px; width: 70px;">
                <asp:Label ID="Label6" runat="server" Text="部門名稱：" CssClass="form"></asp:Label>                    
            </td><td style="height: 43px; width: 90px;" colspan="2">
                    <asp:DropDownList id="PAUNIT"
                        DataSourceID="SqlDataSource1"
                        DataValueField="ORG_NAME"
                        DataTextField="ORG_NAME"
                        runat="server">
                    </asp:DropDownList></td>
                <td noWrap style="height: 43px; width: 50px;">
                    <asp:Label ID="Label11" runat="server" Text="地點：" CssClass="form"></asp:Label></td>
                <td noWrap style="height: 43px; width: 60px;">
                <asp:TextBox ID="nPLACE" runat="server" Width="50px"></asp:TextBox></td>
            </tr> 
            
            <tr><td noWrap style="width: 70px">
                <asp:Label ID="Label3" runat="server" Text="出入日期：" CssClass="form"></asp:Label>
            </td><td style="width: 200px">
                <asp:TextBox ID="nDATE1" runat="server" Width="60px" OnKeyDown="return false"></asp:TextBox>
                <asp:ImageButton ID="ImgDate3" runat="server" ImageUrl="~/Image/calendar.gif" />
                <asp:Label ID="Label4" runat="server" Text="~" CssClass="form"></asp:Label>
                <asp:TextBox ID="nDATE2" runat="server" Width="60px" OnKeyDown="return false"></asp:TextBox>
                <asp:ImageButton ID="ImgDate4" runat="server" ImageUrl="~/Image/calendar.gif" /></td><td noWrap style="width: 55px">
                    <asp:Label ID="Label10" runat="server" Text="事由：" CssClass="form"></asp:Label></td><td style="width: 75px">
                    <asp:DropDownList id="nREASON" runat="server" Width="70px">
                    <asp:ListItem></asp:ListItem>
                    <asp:ListItem Value="演訓運用"> 演訓運用</asp:ListItem>
                    <asp:ListItem Value="業務督考">業務督考</asp:ListItem>
                    <asp:ListItem Value="上級調閱">上級調閱</asp:ListItem>
                    <asp:ListItem Value="營外會報">營外會報</asp:ListItem>
                    <asp:ListItem Value="任務協調">任務協調</asp:ListItem>
                    <asp:ListItem Value="國會備詢">國會備詢</asp:ListItem>
                    <asp:ListItem Value="司法約詢">司法約詢</asp:ListItem>
                    <asp:ListItem Value="公務出國">公務出國</asp:ListItem>
                    <asp:ListItem Value="保修維護">保修維護</asp:ListItem>
                    <asp:ListItem Value="其他">其他</asp:ListItem>
                </asp:DropDownList></td><td noWrap style="width: 70px">
                    &nbsp;<asp:Label ID="Label9" runat="server" Text="機密等級：" CssClass="form"></asp:Label></td><td colspan="2" style="width: 90px">
                    &nbsp;<asp:DropDownList id="nClass" runat="server" Width="70px">
                    <asp:ListItem></asp:ListItem>
                    <asp:ListItem Value="普通">普通</asp:ListItem>
                    <asp:ListItem Value="密">密</asp:ListItem>
                    <asp:ListItem Value="機密">機密</asp:ListItem>
                    <asp:ListItem>極機密</asp:ListItem>
                    <asp:ListItem Value="絕對機密">絕對機密</asp:ListItem>
                </asp:DropDownList></td>
                    
                <td noWrap style="height: 43px;" colspan="2">
                <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢"/>
                    <asp:ImageButton ID="ImgClearAll" runat="server" ImageUrl="~/Image/clearall.gif" ToolTip="清除"/></td>
            </tr>
		</table>
	    <asp:GridView ID="GridView1" runat="server" EmptyDataText="查無任何資料" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" DataKeyNames="EFORMSN" DataSourceID="SqlDataSource2" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
            <Columns>
                <asp:BoundField DataField="nAPPTIME" HeaderText="申請日期" SortExpression="nAPPTIME">
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="PAUNIT" HeaderText="部門名稱" SortExpression="PAUNIT" >
                    <HeaderStyle HorizontalAlign="Center" Width="25%" />
                </asp:BoundField>
                <asp:BoundField DataField="nREASON" HeaderText="事由" SortExpression="nREASON" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="nDATE" HeaderText="出入日期" SortExpression="nDATE" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:BoundField DataField="nPLACE" HeaderText="地點" SortExpression="nPLACE" >
                    <HeaderStyle HorizontalAlign="Center" Width="25%" />
                </asp:BoundField>
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
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="">
        </asp:SqlDataSource>
        
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="select '' as ORG_UID,'' as ORG_NAME union SELECT [ORG_NAME], [ORG_NAME] FROM [ADMINGROUP] ORDER BY [ORG_NAME]">
        </asp:SqlDataSource>
        
        
    <div id="Div_grid" runat="server" style="border-right: lightslategray 2px solid;
        border-top: lightslategray 2px solid; display: block; z-index: 3; left: 14px;
        border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
        position: absolute; top: 689px; height: 150pt; background-color: white" visible="false">
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
        border-top: lightslategray 2px solid; display: block; z-index: 3; left: 248px;
        border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
        position: absolute; top: 690px; height: 150pt; background-color: white" visible="false">
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
        
    <div id="Div_grid3" runat="server" style="border-right: lightslategray 2px solid; border-top: lightslategray 2px solid;
        display: block; z-index: 3; left: 480px; border-left: lightslategray 2px solid;
        width: 165pt; border-bottom: lightslategray 2px solid; position: absolute; top: 692px;
        height: 150pt; background-color: white" visible="false">
        <asp:Calendar ID="Calendar3" runat="server" BackColor="White" BorderColor="#3366CC"
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
        <asp:Button ID="btnClose3" runat="server" Text="關閉" Width="220px" /></div>
        
    <div id="Div_grid4" runat="server" style="border-right: lightslategray 2px solid; border-top: lightslategray 2px solid;
        display: block; z-index: 3; left: 714px; border-left: lightslategray 2px solid;
        width: 165pt; border-bottom: lightslategray 2px solid; position: absolute; top: 693px;
        height: 150pt; background-color: white" visible="false">
        <asp:Calendar ID="Calendar4" runat="server" BackColor="White" BorderColor="#3366CC"
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
        <asp:Button ID="btnClose4" runat="server" Text="關閉" Width="220px" /></div>

    </form>
</body>
</html>
