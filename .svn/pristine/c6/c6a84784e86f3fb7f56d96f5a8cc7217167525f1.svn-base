<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04106.aspx.vb" Inherits="M_Source_MOA04106" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>派工查詢</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">     
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="派工查詢" Width="100%"></asp:Label>
            </td></tr>
        </table>
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
            <tr><td noWrap>
                <asp:Label ID="Label3" runat="server" Text="開工時間：" CssClass="form" /></td><td noWrap>
                <asp:TextBox ID="nStartDATE1" runat="server" Width="65px" OnKeyDown="return false" />
                    <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />
                <asp:Label ID="Label4" runat="server" Text="~" CssClass="form"></asp:Label>
                <asp:TextBox ID="nStartDATE2" runat="server" Width="65px" OnKeyDown="return false"></asp:TextBox>
                    <asp:ImageButton ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" /></td><td noWrap>
                <asp:Label ID="Label5" runat="server" Text="申請人：" CssClass="form"></asp:Label>                    
                <asp:TextBox ID="PANAME" runat="server" Width="60px"></asp:TextBox>
            </td><td noWrap>
                <asp:Label ID="Label7" runat="server" Text="聯絡電話：" CssClass="form"></asp:Label>                    
                <asp:TextBox ID="nPHONE" runat="server" Width="80px"></asp:TextBox>
            </td><td noWrap colspan="2">
                <asp:Label ID="Label8" runat="server" Text="請修地點：" CssClass="form"></asp:Label>
                <asp:TextBox ID="nPLACE" runat="server" Width="99px"></asp:TextBox></td></tr> 
            <tr><td noWrap style="height: 26px">
                <asp:Label ID="Label11" runat="server" Text="請修事項：" CssClass="form"></asp:Label></td><td style="height: 26px">
                <asp:TextBox ID="nFIXITEM" runat="server" Width="180px"></asp:TextBox></td>
                <td noWrap style="height: 26px" colspan="2">
                <asp:Label ID="Label6" runat="server" Text="部門名稱：" CssClass="form"></asp:Label>
                    <asp:DropDownList id="PAUNIT"
                        DataSourceID="SqlDataSource1"
                        DataValueField="ORG_NAME"
                        DataTextField="ORG_NAME"
                        runat="server">
                    </asp:DropDownList></td>
                <td style="height: 26px"><asp:Label ID="Label1" runat="server" Text="處理狀態：" CssClass="form" />
                <asp:DropDownList ID="ddl_Status" runat="server">
                    <asp:ListItem Value="0">全部</asp:ListItem>
                    <asp:ListItem Value="1">新送件</asp:ListItem>
                    <asp:ListItem Value="2">處理中</asp:ListItem>
                    <asp:ListItem Value="3">待料中</asp:ListItem>
                    <asp:ListItem Value="4">完工</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td style="height: 26px">
                <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢"/>
                <asp:ImageButton ID="ImgClearAll" runat="server" ImageUrl="~/Image/clearall.gif" ToolTip="清除畫面"/>
		    </td></tr>
		</table>
	    <asp:GridView ID="GridView1" runat="server" EmptyDataText="查無任何資料" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" DataSourceID="SqlDataSource2" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None" DataKeyNames="P_NUM" CssClass="form">
            <Columns>
                <asp:BoundField DataField="P_NUM" HeaderText="流水號" SortExpression="P_NUM" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                </asp:BoundField>
                <asp:BoundField DataField="PAUNIT" HeaderText="部門名稱" SortExpression="PAUNIT" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="PANAME" HeaderText="申請人" SortExpression="PANAME" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="nAPPTIME" HeaderText="申請時間" SortExpression="nAPPTIME" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="nPHONE" HeaderText="聯絡電話" SortExpression="nPHONE" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="nPLACE" HeaderText="請修地點" SortExpression="nPLACE" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="nFIXITEM" HeaderText="請修事項" SortExpression="nFIXITEM" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                 <asp:TemplateField ShowHeader="False" HeaderText="開工時間" >
                    <ItemTemplate>
                      <%# CheckShowDATE(Eval("nStartDATE"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                </asp:TemplateField>

                <asp:TemplateField ShowHeader="False" HeaderText="完工日期" >
                    <ItemTemplate>
                      <%# CheckShowDATE(Eval("nFinalDate"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                </asp:TemplateField>
                <asp:TemplateField ShowHeader="False" HeaderText="處理狀態" >
                    <ItemTemplate>
                      <%# StatusName(Eval("FlowStatus"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                </asp:TemplateField>
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:ImageButton ID="ImgList" runat="server" CausesValidation="False" CommandName="Select"
                            ImageUrl="~/Image/List.gif" OnClick="ImgList_Click" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                </asp:TemplateField>
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
         SelectCommand ="SELECT P_NUM, convert(nvarchar,nAPPTIME,111) as nAPPTIME, isnull(PANAME,'') as PANAME,  isnull(PAUNIT,'') as PAUNIT,isnull(nStartDATE,'') as nStartDATE,
            b.bd_name+'/'+c.fl_name+'/'+d.rnum_name as nPLACE,isnull(nFIXITEM,'') as nFIXITEM,isnull(nPHONE,'') as nPHONE,isnull(nFinalDate,'') as nFinalDate,isnull(FlowStatus,0) as FlowStatus  
            FROM P_0415 a left join P_0404 b on a.nbd_code = b.bd_code
                         left join P_0406 c on a.nfl_code = c.fl_code
                         left join P_0411 d on a.nrnum_code = d.rnum_code
                         Order by nAPPTIME" >
        </asp:SqlDataSource>
        
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="select '' as ORG_UID,'' as ORG_NAME union SELECT [ORG_NAME], [ORG_NAME] FROM [ADMINGROUP] ORDER BY [ORG_NAME]">
        </asp:SqlDataSource>

    &nbsp;
        <div id="Div_grid" runat="server" style="border-right: lightslategray 2px solid;
            border-top: lightslategray 2px solid; display: block; z-index: 3; left: 194px;
            border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
            position: absolute; top: 824px; height: 150pt; background-color: white" visible="false">
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
            position: absolute; top: 825px; height: 150pt; background-color: white" visible="false">
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
            <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" 
ControlToValidate="PANAME" Display="None" ErrorMessage="申請人輸入不可使用特殊符號" 
ValidationExpression="[^<*>&#%']*" ></asp:RegularExpressionValidator> 
    <asp:RegularExpressionValidator ID="RegularExpressionValidator2" runat="server" 
ControlToValidate="nPHONE" Display="None" ErrorMessage="聯絡電話輸入不可使用特殊符號" 
ValidationExpression="[^<*>&#%']*" ></asp:RegularExpressionValidator> 
    <asp:RegularExpressionValidator ID="RegularExpressionValidator3" runat="server" 
ControlToValidate="nPLACE" Display="None" ErrorMessage="請修地點輸入不可使用特殊符號" 
ValidationExpression="[^<*>&#%']*" ></asp:RegularExpressionValidator> 
    <asp:RegularExpressionValidator ID="RegularExpressionValidator4" runat="server" 
ControlToValidate="nFIXITEM" Display="None" ErrorMessage="請修事項輸入不可使用特殊符號" 
ValidationExpression="[^<*>&#%']*" ></asp:RegularExpressionValidator> 
<asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True" ShowSummary="False" />
    </form>
</body>
</html>
