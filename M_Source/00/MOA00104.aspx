<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00104.aspx.vb" Inherits="Source_00_MOA00104" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>重新分派</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" language="javascript">
        document.oncontextmenu = new Function("return false");
    </script>
</head>
<body>
    <form id="form1" runat="server">   
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="重新分派" Width="100%"></asp:Label>
            </td></tr>
        </table>
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap style="height: 25px; width: 10%;">
                    <asp:Label ID="Lab1" runat="server" Text="申請日期：" CssClass="form"></asp:Label>                    
                </td>
			    <td noWrap style="height: 25px; width: 70%;">			    
                    <asp:TextBox ID="Sdate" runat="server" OnKeyDown="return false" Width="70px"></asp:TextBox>
                    <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />
                    <asp:Label ID="Lab3" runat="server" Text="~" ></asp:Label>
                    <asp:TextBox ID="Edate" runat="server" OnKeyDown="return false" Width="70px"></asp:TextBox>
                    <asp:ImageButton ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" />
                </td>
                
			    <td noWrap style="height: 25px; width: 20%;">
                    <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />&nbsp;
                    <asp:ImageButton ID="ImgAssign" runat="server" ImageUrl="~/Image/assign.gif" ToolTip="重新分派" /></td>
		    </tr>
	    </table>
        <asp:GridView ID="GridView1" runat="server" DataSourceID="SqlDataSource2" AutoGenerateColumns="False" AllowPaging="True" EmptyDataText="查無任何資料" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None" AllowSorting="True">
            <Columns>
                <asp:TemplateField HeaderText="選取">
                    <EditItemTemplate>
                        <asp:CheckBox ID="CheckBox1" runat="server" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:CheckBox ID="selchk" runat="server" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                </asp:TemplateField>
                <asp:BoundField DataField="flowsn" HeaderText="flowsn" InsertVisible="False" ReadOnly="True"
                    SortExpression="flowsn" />
                <asp:BoundField DataField="appdate" HeaderText="申請日期" SortExpression="appdate" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="22%" />
                </asp:BoundField>
                <asp:BoundField DataField="PWNAME" HeaderText="填表者" SortExpression="PWNAME">
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="PANAME" HeaderText="申請者" SortExpression="PANAME">
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="emp_chinese_name" HeaderText="批核者" SortExpression="emp_chinese_name">
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="ShowContent" HeaderText="表單內容" SortExpression="ShowContent" >
                    <HeaderStyle HorizontalAlign="Center" Width="40%" />
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
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap style="height: 25px; width: 20%;">
                    <asp:Label ID="Label3" runat="server" CssClass="form" Text="重新分派單位：" Width="100px"></asp:Label></td>
			    <td noWrap style="height: 25px; width: 80%;">
                    <asp:DropDownList id="OrgSel"
                        DataSourceID="SqlDataSource1"
                        DataValueField="ORG_UID"
                        DataTextField="ORG_NAME"
                        runat="server" AutoPostBack="True">
                    </asp:DropDownList></td>
		    </tr>
		    <tr>
			    <td noWrap style="height: 25px; width: 20%;">
                    <asp:Label ID="Label4" runat="server" CssClass="form" Text="重新分派姓名：" Width="100px"></asp:Label></td>
			    <td noWrap style="height: 25px; width: 80%;">
                    <asp:DropDownList ID="UserSel" runat="server" DataSourceID="SqlDataSource3" DataTextField="emp_chinese_name" DataValueField="employee_id">
                    </asp:DropDownList>
                    <asp:Label ID="ErrName" runat="server" Text="請選擇姓名" Width="100px" ForeColor="Red" Visible="False"></asp:Label></td>
		    </tr>
	    </table>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] ORDER BY [ORG_NAME]">
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT flowctl.flowsn,flowctl.eformid, flowctl.eformrole, flowctl.appdate, flowctl.empuid, flowctl.emp_chinese_name, EFORMS.frm_chinese_name, flowctl.eformsn, flowctl.hddate, V_EformShow.ShowContent,V_EformShow.PWNAME,V_EformShow.PANAME FROM flowctl INNER JOIN EFORMS ON flowctl.eformid = EFORMS.eformid INNER JOIN V_EformShow ON flowctl.eformsn = V_EformShow.EFORMSN WHERE (flowctl.hddate IS NULL) and (flowctl.gonogo='?' OR flowctl.gonogo='R') AND 1=2 ">
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [employee_id], [emp_chinese_name], [ORG_UID] FROM [EMPLOYEE] WHERE ([ORG_UID] = @ORG_UID) ORDER BY [emp_chinese_name]">
            <SelectParameters>
                <asp:ControlParameter ControlID="OrgSel" Name="ORG_UID" PropertyName="SelectedValue"
                    Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
        
        
        <div id="Div_grid" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:165pt; height:150pt; left:197px; top:650px; display:block;" visible="false">
         
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
        
        <div id="Div_grid2" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:165pt; height:150pt; left:428px; top:650px; display:block;" visible="false">
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
