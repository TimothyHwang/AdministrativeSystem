<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA03003.aspx.vb" Inherits="Source_03_MOA03003" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>車輛派遣使用狀況</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server" >     
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="車輛派遣使用狀況" Width="100%"></asp:Label>
            </td></tr>
        </table>
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
            <tr><td noWrap style="height: 24px">
                <asp:Label ID="Label1" runat="server" Text="報到日期：" CssClass="form"></asp:Label>
            </td><td noWrap style="height: 24px">
                <asp:TextBox ID="nARRDATE1" runat="server" Width="60px" OnKeyDown="return false"></asp:TextBox>
                <asp:ImageButton
                    ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />
                <asp:Label ID="Label2" runat="server" Text="~" CssClass="form"></asp:Label>
                <asp:TextBox ID="nARRDATE2" runat="server" Width="60px" OnKeyDown="return false"></asp:TextBox>
                <asp:ImageButton
                    ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" />
            </td><td noWrap style="height: 24px">
                <asp:Label ID="Label5" runat="server" Text="駕駛人姓名：" CssClass="form"></asp:Label>                    
            </td><td style="height: 24px">
                <asp:TextBox ID="DriveName" runat="server" Width="60px"></asp:TextBox>
            </td><td noWrap style="height: 24px">
                <asp:Label ID="Label7" runat="server" Text="駕駛人電話：" CssClass="form"></asp:Label>                    
            </td><td style="height: 24px">
                <asp:TextBox ID="DriveTel" runat="server" Width="60px"></asp:TextBox>
            </td><td noWrap style="height: 24px">
                <asp:Label ID="Label11" runat="server" Text="車輛種類：" CssClass="form"></asp:Label>                    
            </td><td colspan='2' style="height: 24px">
                <asp:DropDownList id="nSTYLE"
                    AutoPostBack="False"
                    DataSourceID="SqlDataSource2"
                    DataValueField="PCK_Name"
                    DataTextField="PCK_Name"
                    runat="server"/>
            </td></tr> 
            <tr><td noWrap>
                <asp:Label ID="Label6" runat="server" Text="報到地點：" CssClass="form"></asp:Label></td><td>
                <asp:TextBox ID="nARRIVEPLACE" runat="server" Width="150px"></asp:TextBox></td><td noWrap>
                <asp:Label ID="Label8" runat="server" Text="車輛號碼：" CssClass="form"></asp:Label></td><td>
                <asp:TextBox ID="CarNumber" runat="server" Width="60px"></asp:TextBox></td><td noWrap>
                <asp:Label ID="Label9" runat="server" Text="狀態：" CssClass="form"></asp:Label></td><td>
                <asp:DropDownList id="CarStatus" runat="server">
                    <asp:ListItem></asp:ListItem>
                    <asp:ListItem Value="未出場">未出場</asp:ListItem>
                    <asp:ListItem Value="已出場">已出場</asp:ListItem>
                    <asp:ListItem Value="回場">回場</asp:ListItem>
                </asp:DropDownList></td><td noWrap>
            </td><td>
                &nbsp;</td><td noWrap align='left'>
                <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" title="查詢" ToolTip="查詢"/>
                <asp:ImageButton ID="ImgClearAll" runat="server" ImageUrl="~/Image/clearall.gif" title="畫面清除"/>
		    </td></tr>
		</table>
	    <asp:GridView ID="GridView1" runat="server" EmptyDataText="查無任何資料" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" DataSourceID="SqlDataSource1" Width="100%" DataKeyNames="EFORMSN" CellPadding="4" ForeColor="#333333" GridLines="None">
            <Columns>
                <asp:BoundField DataField="DriveName" HeaderText="駕駛人姓名" SortExpression="DriveName" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="DriveTel" HeaderText="駕駛人電話" SortExpression="DriveTel" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nSTYLE" HeaderText="車輛種類" SortExpression="nSTYLE" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="CarNumber" HeaderText="車輛號碼" SortExpression="CarNumber" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nARRIVEPLACE" HeaderText="報到地點" SortExpression="nARRIVEPLACE" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="報到日期" SortExpression="nARRDATE">
                    <ItemTemplate>
                        <asp:Label ID="ShowDate" runat="server" Text='<%# ShowDate("nUSEDATE","nSTUSEHOUR","nSTUSEMIN","nEDUSEDATE","nEDUSEHOUR","nEDUSEMIN") %>'></asp:Label>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:BoundField DataField="LeaveTime" HeaderText="出場時間" SortExpression="nARRIVEPLACE" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="ComeTime" HeaderText="回場時間" SortExpression="nARRIVEPLACE" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="狀態" SortExpression="CarStatus">
                    <ItemTemplate>
                        <asp:Label ID="CarStatus" runat="server" Text='<%# ShowCarStatus() %>' ForeColor='<%# CarStatusColor() %>'></asp:Label>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:ImageButton id="ImgPrint" runat="server" title="列印"  CommandName="Select" ImageUrl="../../Image/print.gif" OnClick="ImgPrint_Click"></asp:ImageButton>
                        <asp:Button ID="BtnCarStatus" runat="server" CausesValidation="True" CommandName="Select"  Text="<%# CarStatusText() %>" Visible='<%# CarStatusBtn() %>' OnClientClick="return confirm('確定更新狀態嗎?')"/>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
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
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="" UpdateCommand="UPDATE [P_0301] SET [CarStatus] = @CarStatus WHERE [EFORMSN] = @EFORMSN"
                    ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
            <UpdateParameters>
                <asp:Parameter Name="EFORMSN" Type="String" />
                <asp:Parameter Name="CarStatus" Type="String" />
            </UpdateParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="select '' as PCK_Name union SELECT [PCK_Name] FROM [P_0302] ORDER BY [PCK_Name]"
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
        </asp:SqlDataSource>        
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="">
        </asp:SqlDataSource>
        
        <div id="Div_grid" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:165pt; height:150pt; left:203px; top:781px; display:block;" visible="false">
         
            <asp:Calendar ID="Calendar1" runat="server" BackColor="White" BorderColor="#3366CC"
            BorderWidth="1px" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana"
            Font-Size="8pt" ForeColor="#003399" Height="200px" Width="220px" Caption="" ShowGridLines="True">
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
        
        <div id="Div_grid2" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:165pt; height:150pt; left:431px; top:779px; display:block;" visible="false">
        <asp:Calendar ID="Calendar2" runat="server" BackColor="White" BorderColor="#3366CC"
            BorderWidth="1px" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana"
            Font-Size="8pt" ForeColor="#003399" Height="200px" Width="220px" ShowGridLines="True">
            <SelectedDayStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
            <TodayDayStyle BackColor="#99CCCC" ForeColor="White" />
            <SelectorStyle BackColor="#99CCCC" ForeColor="#336666" />
            <WeekendDayStyle BackColor="#CCCCFF" />
            <OtherMonthDayStyle ForeColor="#999999" />
            <NextPrevStyle Font-Size="8pt" ForeColor="#CCCCFF" />
            <DayHeaderStyle BackColor="#99CCCC" ForeColor="#336666" Height="1px" />
            <TitleStyle BackColor="#003399" BorderColor="#3366CC" BorderWidth="1px" Font-Bold="True"
                Font-Size="10pt" ForeColor="#CCCCFF" Height="25px" />
        </asp:Calendar><asp:Button ID="btnClose2" runat="server" Text="關閉" Width="220px" /></div>   
    </form>
</body>
</html>
