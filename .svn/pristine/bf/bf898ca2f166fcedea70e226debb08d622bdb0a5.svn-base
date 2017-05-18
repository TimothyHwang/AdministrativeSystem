<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00102.aspx.vb" Inherits="Source_00_MOA00102" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>代理人管理</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />

    <link href="../../css/jquery.datepick.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.10.2.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick-zh-TW.js" type="text/javascript"></script>
    <script type="text/javascript" language="javascript">
        $(function () {
            $("#Agent_SDate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#Agent_SDate").val(), showTrigger: '#calImg' });
            $("#Agent_EDate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#Agent_EDate").val(), showTrigger: '#calImg' });
        }); 
    </script>
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="代理人管理" Width="100%"></asp:Label>
            </td></tr>
        </table>
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
            <tr>
            <td noWrap width="10%">
                <asp:Label ID="Label1" runat="server" Text="部門名稱：" CssClass="form"></asp:Label></td>
            <td width="30%">
                <asp:DropDownList id="Dep"
                    DataSourceID="SqlDataSource1"
                    DataValueField="ORG_UID"
                    DataTextField="ORG_NAME"
                    runat="server" AutoPostBack="True">
                </asp:DropDownList></td>
            <td noWrap width="10%">
                <asp:Label ID="Label2" runat="server" Text="代理人：" CssClass="form"></asp:Label></td>
            <td width="20%">
                <asp:DropDownList AutoPostBack="True" ID="Agent1" runat="server" DataSourceID="SqlDataSource3" DataTextField="emp_chinese_name" DataValueField="employee_id">
                </asp:DropDownList></td>
            <td noWrap style="width: 10%">
                <asp:Label ID="Label5" runat="server" Text="被代理人：" CssClass="form"></asp:Label></td>
            <td width="20%">
                <asp:DropDownList AutoPostBack="True" ID="Agent2" runat="server" DataSourceID="SqlDataSource4" DataTextField="emp_chinese_name" DataValueField="employee_id">
                </asp:DropDownList></td>
            </tr>
            
            <tr>
            <td noWrap>
                <asp:Label ID="Label3" runat="server" Text="代理日期：" CssClass="form"></asp:Label></td>
            <td colspan="2">
                <asp:TextBox ID="Agent_SDate" runat="server" Width="60px" OnKeyDown="return false"></asp:TextBox>
                <div style="display: none;">
                    <img id="calImg" src="../../Image/calendar.gif" alt="選擇日期" />
                </div>
                <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" 
                    Visible="False" />
                <asp:Label ID="Label4" runat="server" Text="~" CssClass="form"></asp:Label>
                <asp:TextBox ID="Agent_EDate" runat="server" Width="60px" OnKeyDown="return false"></asp:TextBox>
                <asp:ImageButton ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" 
                    Visible="False" /></td>
            <td colspan="3">
                <asp:ImageButton ID="ImgInsert" runat="server" ImageUrl="../../Image/add.gif" ToolTip="新增"/>
                <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢"/>
                <asp:ImageButton ID="ImgClearAll" runat="server" ImageUrl="~/Image/clearall.gif" ToolTip="清除畫面"/>
                <asp:Label ID="ErrMsg" runat="server" ForeColor='Red' Text="" ></asp:Label></td>
		    </tr>
		</table>
	    <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" EmptyDataText="查無任何資料" AllowPaging="True" AllowSorting="True" DataKeyNames="Agent_Num" DataSourceID="SqlDataSource2" Width="100%" GridLines="None" CellPadding="4" ForeColor="#333333">
            <Columns>
                <asp:BoundField DataField="DepName" HeaderText="部門名稱" SortExpression="DepName" >
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:BoundField DataField="Agent_Name1" HeaderText="代理人" SortExpression="Agent_Name1" >
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="Agent_Name2" HeaderText="被代理人" SortExpression="Agent_Name2" >
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="Agent_SDate" HeaderText="代理日期(起)" SortExpression="Agent_SDate" >
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="Agent_EDate" HeaderText="代理日期(迄)" SortExpression="Agent_EDate" >
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:ImageButton ID="ImgDel" runat="server" CommandName="Delete" ImageUrl="~/Image/delete.gif"
                            OnClientClick="return confirm('確定刪除代理人嗎?')" ToolTip="刪除" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
            </Columns>
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [AGENT] WHERE [Agent_Num] = @Agent_Num" 
            InsertCommand=""
            SelectCommand=""
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
            <DeleteParameters>
                <asp:Parameter Name="Agent_Num" Type="Int32" />
            </DeleteParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="select '' as ORG_UID,'' as ORG_NAME union SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] ORDER BY [ORG_NAME]">
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [employee_id], [emp_chinese_name], [ORG_UID] FROM [EMPLOYEE] WHERE EMPLOYEE.[ORG_UID]=@ORG_UID ORDER BY [emp_chinese_name]">
            <SelectParameters>
            <asp:ControlParameter ControlID="Dep" Name="ORG_UID" PropertyName="SelectedValue" Type="String" /></SelectParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource4" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [employee_id], [emp_chinese_name], [ORG_UID] FROM [EMPLOYEE] WHERE EMPLOYEE.[ORG_UID]=@ORG_UID ORDER BY [emp_chinese_name]">
            <SelectParameters>
            <asp:ControlParameter ControlID="Dep" Name="ORG_UID" PropertyName="SelectedValue" Type="String" /></SelectParameters>
        </asp:SqlDataSource>
        
        <div id="Div_grid" runat="server" style="border-right: lightslategray 2px solid;
            border-top: lightslategray 2px solid; display: block; z-index: 3; left: 194px;
            border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
            position: absolute; top: 503px; height: 150pt; background-color: white" visible="false">
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
            border-top: lightslategray 2px solid; display: block; z-index: 3; left: 427px;
            border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
            position: absolute; top: 503px; height: 150pt; background-color: white" visible="false">
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
