<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00200.aspx.vb" Inherits="Source_00_MOA00200" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>簽到退查詢</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" language="javascript">
        document.oncontextmenu = new Function("return false");
    </script>
</head>
<body language="javascript" onload="return window_onload()">
    <form id="form1" runat="server">   
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="簽到退查詢" Width="100%"></asp:Label>
            </td></tr>
        </table>     
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Lab1" runat="server" Text="單位：" CssClass="form"></asp:Label>                    
                </td>
			    <td noWrap style="height: 25px; width: 35%;">
                    <asp:DropDownList ID="OrgSel" runat="server" DataSourceID="SqlDataSource1" DataTextField="ORG_NAME" DataValueField="ORG_UID" AutoPostBack="True">
                    </asp:DropDownList></td>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Label1" runat="server" Text="姓名：" CssClass="form"></asp:Label>
                </td>
			    <td noWrap style="height: 25px; width: 15%;">
                    <asp:DropDownList ID="UserSel" runat="server" DataSourceID="SqlDataSource3" DataTextField="emp_chinese_name"
                        DataValueField="employee_id">
                    </asp:DropDownList></td>
			    <td noWrap style="height: 25px; width: 5%;">
                          <asp:Label ID="Label3" runat="server" Text="加班：" CssClass="form"></asp:Label></td>
			    <td noWrap style="height: 25px; width: 15%;">
                    <asp:DropDownList ID="DrDn_add" runat="server">
                        <asp:ListItem Value="請選擇">請選擇</asp:ListItem>
                        <asp:ListItem>18:00</asp:ListItem>
                        <asp:ListItem>18:30</asp:ListItem>
                        <asp:ListItem>19:00</asp:ListItem>
                        <asp:ListItem>19:30</asp:ListItem>
                        <asp:ListItem>20:00</asp:ListItem>
                    </asp:DropDownList></td>
                
			    <td noWrap style="height: 25px; width: 20%;">
                    <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />&nbsp;<asp:ImageButton
                        ID="Img_Export" runat="server" ImageUrl="~/Image/ExportFile.gif" />
                    <asp:ImageButton ID="ImagePrint" runat="server" ImageUrl="~/Image/print.gif" ToolTip="列印" /></td>
		    </tr>
		    <tr>
		        <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Label2" runat="server" Text="日期：" CssClass="form"></asp:Label></td>
                <td noWrap style="height: 25px; width: 35%;">
                    <asp:TextBox ID="Sdate" runat="server" Width="80px" OnKeyDown="return false"></asp:TextBox>
                    <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />
                    ~
                    <asp:TextBox ID="Edate" runat="server" Width="80px" OnKeyDown="return false"></asp:TextBox>
                    <asp:ImageButton ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" /></td>
		        <td noWrap style="height: 25px; width: 5%;">
                          <asp:Label ID="Label4" runat="server" Text="簽到狀態：" CssClass="form"></asp:Label></td>
			    <td noWrap style="height: 25px; width: 15%;">
                    <asp:DropDownList ID="DrDn_in" runat="server">
                        <asp:ListItem>請選擇</asp:ListItem>
                        <asp:ListItem>已簽到</asp:ListItem>
                        <asp:ListItem>未簽到</asp:ListItem>                        
                    </asp:DropDownList></td>
		        <td noWrap style="height: 25px; width: 5%;">
                          <asp:Label ID="Label5" runat="server" Text="簽退狀態：" CssClass="form"></asp:Label></td>
			    <td noWrap style="height: 25px; width: 15%;">
                    <asp:DropDownList ID="DrDn_out" runat="server">
                        <asp:ListItem>請選擇</asp:ListItem>
                        <asp:ListItem>已簽退</asp:ListItem>
                        <asp:ListItem>未簽退</asp:ListItem>                        
                    </asp:DropDownList></td>
                <td noWrap style="height: 25px; width: 20%;">
                          </td>
			    
		    </tr>
	    </table>
        <asp:GridView ID="GridView1" runat="server" EmptyDataText="查無任何資料" AllowPaging="True" AllowSorting="True"
            AutoGenerateColumns="False" CellPadding="4" DataSourceID="SqlDataSource2" ForeColor="#333333"
            GridLines="None" Width="100%" PageSize="31">
            <Columns>
                <asp:BoundField DataField="emp_chinese_name" HeaderText="人員姓名" SortExpression="emp_chinese_name">
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="ShowDay" HeaderText="日期" SortExpression="ShowDay" >
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="In_Time_nvc" HeaderText="簽到時間" SortExpression="In_Time_nvc" NullDisplayText="未簽到" >
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="Out_Time_nvc" HeaderText="簽退時間" SortExpression="Out_Time_nvc" >
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="REASON_nvc" HeaderText="遲到理由" SortExpression="REASON_nvc" >
                    <HeaderStyle HorizontalAlign="Center" Width="30%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="Out_REASON" HeaderText="加班理由" SortExpression="Out_REASON" >
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
            SelectCommand="SELECT * FROM [ADMINGROUP] WHERE 1=2"></asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT * FROM T_SIGN_RECORD,EMPLOYEE WHERE T_SIGN_RECORD.LOGONID_nvc = EMPLOYEE.employee_id AND 1=2"></asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [employee_id], [emp_chinese_name], [ORG_UID] FROM [EMPLOYEE] WHERE 1=2 ORDER BY [emp_chinese_name]">
            <SelectParameters>
                <asp:ControlParameter ControlID="OrgSel" Name="ORG_UID" PropertyName="SelectedValue"
                    Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
        
         <div id="Div_grid" runat="server" style="border-right: lightslategray 2px solid;
            border-top: lightslategray 2px solid; display: block; z-index: 3; left: 193px;
            border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
            position: absolute; top: 406px; height: 150pt; background-color: white" visible="false">
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
            border-top: lightslategray 2px solid; display: block; z-index: 3; left: 431px;
            border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
            position: absolute; top: 407px; height: 150pt; background-color: white" visible="false">
            <asp:Calendar ID="Calendar2" runat="server" BackColor="White" BorderColor="#3366CC"
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
            <asp:Button ID="btnClose2" runat="server" Text="關閉" Width="220px" /></div> 
            
    <iframe id="lst" frameborder=0 width=0 height=0 src="/blank.htm"></iframe>   
    <script language=javascript>    
    var errmsg='<%= do_sql.G_errmsg%>';
    var print_file='<%= print_file%>';

    function window_onload() {
        if (errmsg != '') {
            alert(errmsg);
        }
        
        if (print_file != '') {
            lst.navigate(print_file);              
        }
    }

    </script>
    </form>
</body>
</html>
