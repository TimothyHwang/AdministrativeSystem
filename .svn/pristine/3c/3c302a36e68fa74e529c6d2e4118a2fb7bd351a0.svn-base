<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA03009.aspx.vb" Inherits="M_Source_03_MOA03009" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>派車批核</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">     
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="派車批核" Width="100%"></asp:Label>
            </td></tr>
        </table>
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
            <tr>
            <td width="10%">
                <asp:Label ID="Label3" runat="server" Text="報到日期：" CssClass="form"></asp:Label></td>
            <td width="30%">
                <asp:TextBox ID="nRECDATE1" runat="server" Width="150px" OnKeyDown="return false"></asp:TextBox>
                <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />
            </td>
            <td width="10%">
                <asp:Label ID="Label2" runat="server" Text="車種：" CssClass="form"></asp:Label></td>
            <td width="30%">
                <asp:DropDownList ID="DDLCar" runat="server" DataSourceID="SqlDataSource1"
                    DataTextField="PCK_Name" DataValueField="PCK_Name">
                </asp:DropDownList></td>
		    <td width="20%">
                <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" title="查詢" ToolTip="查詢"/>
            </td>
		    </tr>
		</table>
	    <asp:GridView ID="GridView1" runat="server" EmptyDataText="查無任何資料" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" DataKeyNames="EFORMSN" DataSourceID="SqlDataSource2" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
            <Columns>
                <asp:BoundField DataField="eformsn" HeaderText="eformsn" SortExpression="eformsn" />
                <asp:BoundField DataField="nARRDATE" HeaderText="報到日期" SortExpression="nARRDATE" >
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="PANAME" HeaderText="申請人" SortExpression="PANAME" >
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                </asp:BoundField>
                <asp:BoundField DataField="PAUNIT" HeaderText="單位" SortExpression="PAUNIT" >
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                </asp:BoundField>
                <asp:BoundField DataField="nARRIVEPLACE" HeaderText="報到地點" SortExpression="nARRIVEPLACE" >
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                </asp:BoundField>
                <asp:BoundField DataField="nREASON" HeaderText="事由" SortExpression="nREASON" >
                    <HeaderStyle HorizontalAlign="Center" Width="25%" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="調度官" SortExpression="stepsid">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("stepsid") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# FunStatus("stepsid") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:CommandField ButtonType="Image" ShowSelectButton="True" SelectImageUrl="~/Image/List.gif" SelectText="詳細資料" >
                    <HeaderStyle Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
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
            SelectCommand="SELECT [PCK_Num], [PCK_Name] FROM [P_0302] ORDER BY [PCK_Num]"></asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT flowctl.appdate, flowctl.eformsn, flowctl.hddate, P_03.nARRDATE, P_03.PANAME, P_03.PAUNIT, P_03.nARRIVEPLACE, P_03.nREASON,flowctl.stepsid FROM flowctl INNER JOIN P_03 ON flowctl.eformsn = P_03.EFORMSN WHERE (flowctl.hddate IS NULL) AND (flowctl.gonogo = '?' OR flowctl.gonogo = 'R') AND (flowctl.eformid = 'j2mvKYe3l9') AND 1=2">
        </asp:SqlDataSource>
        
        <div id="Div_grid" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:165pt; height:150pt; left:278px; top:607px; display:block;" visible="false">
         
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

    </form>
</body>
</html>
