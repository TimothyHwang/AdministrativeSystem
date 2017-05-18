<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00011.aspx.vb" Inherits="Source_00_MOA00011" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>個人表單追蹤</title>

    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <link href="../../css/jquery.datepick.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.10.2.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick-zh-TW.js" type="text/javascript"></script>

    <%-- ReSharper disable Html.IdNotResolved --%>
    <script type="text/javascript">
        $(function () {
            $("#Sdate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#Sdate").val(), showTrigger: '#calImg' });
            $("#Edate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#Edate").val(), showTrigger: '#calImg' });           
        }); 
    </script>
    <%-- ReSharper restore Html.IdNotResolved --%>
    
</head>
<body>
    <form id="form1" runat="server" >
            <table border="0" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;" width="100%">
                <tr><td align="center">
                        <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="個人表單追蹤" Width="100%"></asp:Label>
                </td></tr>
            </table>
	    <table align="center" border="0" cellPadding="0" cellSpacing="0" width="100%">
		    <tr>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Label2" runat="server" Text="表單：" CssClass="form"></asp:Label>                    
                </td>
			    <td noWrap style="height: 25px; width: 35%;">
                    <asp:DropDownList id="FormSel"
                        DataSourceID="SqlDataSource4"
                        DataValueField="eformid"
                        DataTextField="frm_chinese_name"
                        runat="server">
                    </asp:DropDownList></td>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Lab1" runat="server" Text="部門名稱：" CssClass="form"></asp:Label>                    
                </td>
			    <td noWrap style="height: 25px; width: 35%;">
                    <asp:DropDownList id="OrgSel"
                        DataSourceID="SqlDataSource1"
                        DataValueField="ORG_UID"
                        DataTextField="ORG_NAME"
                        runat="server" AutoPostBack="True">
                    </asp:DropDownList></td>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Label1" runat="server" Text="姓名：" CssClass="form"></asp:Label>
                </td>
			    <td noWrap style="height: 25px; width: 25%;">
                    <asp:DropDownList ID="UserSel" runat="server" DataSourceID="SqlDataSource3" DataTextField="emp_chinese_name" DataValueField="employee_id">
                    </asp:DropDownList>
                </td>
		    </tr>		    
		    <tr>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Lab2" runat="server" Text="申請日期：" CssClass="form" ></asp:Label>
                </td>
			    <td noWrap style="height: 25px; width: 35%;">
			        <div style="display: none;">
	                    <img id="calImg" src="../../Image/calendar.gif" alt="Popup">
                    </div>
                    <asp:TextBox ID="Sdate" runat="server" Width="80px" OnKeyDown="return false" 
                        ReadOnly="True" ></asp:TextBox>                   
                    <asp:Label ID="Lab3" runat="server" Text="~" ></asp:Label>
                    <asp:TextBox ID="Edate" runat="server" Width="80px" OnKeyDown="return false" 
                        ReadOnly="True"></asp:TextBox>                   
			    </td>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Label3" runat="server" Text="狀態：" CssClass="form" ></asp:Label>
                </td>
			    <td noWrap style="height: 25px; width: 25%;">
                    <asp:DropDownList ID="StateSel" runat="server">
                        <asp:ListItem>請選擇</asp:ListItem>
                        <asp:ListItem>已批核</asp:ListItem>
                        <asp:ListItem>未批核</asp:ListItem>
                        <asp:ListItem>駁回</asp:ListItem>
                    </asp:DropDownList></td>
			    <td noWrap style="height: 25px; width: 25%;" align="left" colspan="2">
                    <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />&nbsp;<asp:ImageButton
                        ID="ImgRevoke" runat="server" ImageUrl="~/Image/revoke.gif" OnClientClick="return confirm('表單確定撤銷嗎?')"
                        ToolTip="撤銷" />
                    </td>
		    </tr>
	    </table>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True" EmptyDataText="查無任何資料"
            AutoGenerateColumns="False" DataSourceID="SqlDataSource2" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
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
                <asp:BoundField DataField="emp_chinese_name" HeaderText="填表人" SortExpression="emp_chinese_name" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="frm_chinese_name" HeaderText="表單" SortExpression="frm_chinese_name" >
                    <HeaderStyle HorizontalAlign="Center" Width="17%" />
                </asp:BoundField>
                <asp:BoundField DataField="ORG_NAME" HeaderText="部門名稱" SortExpression="ORG_NAME" >
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:BoundField DataField="appdate" HeaderText="申請日期" SortExpression="appdate" >
                    <HeaderStyle HorizontalAlign="Center" Width="25%" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="狀態" SortExpression="status">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("status") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# FunStatus("status") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:CommandField ButtonType="Image" SelectImageUrl="~/Image/List.gif" ShowSelectButton="True" SelectText="詳細資料" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:CommandField>
                <asp:BoundField DataField="eformsn" HeaderText="eformsn" SortExpression="eformsn" />
                <asp:BoundField DataField="eformid" HeaderText="eformid" SortExpression="eformid" />
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
            SelectCommand="SELECT flowctl.flowsn, flowctl.eformsn, flowctl.eformid, flowctl.empuid, flowctl.emp_chinese_name, flowctl.appdate, flowctl.deptcode, EFORMS.frm_chinese_name, ADMINGROUP.ORG_NAME, (SELECT TOP (1) gonogo FROM flowctl AS f WHERE (flowctl.eformsn = eformsn) ORDER BY flowsn DESC) AS status FROM flowctl INNER JOIN EFORMS ON flowctl.eformid = EFORMS.eformid INNER JOIN ADMINGROUP ON flowctl.deptcode = ADMINGROUP.ORG_UID WHERE (flowctl.gonogo = '-') and 1=2">
        </asp:SqlDataSource>
    
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] ORDER BY [ORG_NAME]">
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [employee_id], [emp_chinese_name], [ORG_UID] FROM [EMPLOYEE] WHERE ([ORG_UID] = @ORG_UID) ORDER BY [emp_chinese_name]">
            <SelectParameters>
                <asp:ControlParameter ControlID="OrgSel" Name="ORG_UID" PropertyName="SelectedValue"
                    Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>    
        <asp:SqlDataSource ID="SqlDataSource4" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [eformid], [frm_chinese_name] FROM [EFORMS] ORDER BY [PRIMARY_TABLE]"></asp:SqlDataSource>
            
            
            
            
            
        <div id="Div_grid" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:165pt; height:150pt; left:202px; top:640px; display:block;" visible="false">
         
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
        
        <div id="Div_grid2" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:165pt; height:150pt; left:435px; top:639px; display:block;" visible="false">
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
