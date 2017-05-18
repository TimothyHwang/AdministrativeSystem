<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA01004.aspx.vb" Inherits="Source_01_MOA01004" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>銷假申請</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">      
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="銷假申請" Width="100%"></asp:Label>
            </td></tr>
        </table>           
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Label1" runat="server" Text="假別：" CssClass="form"></asp:Label>
                </td>
			    <td noWrap style="height: 25px; width: 30%;">
                    <asp:DropDownList ID="DDLType" runat="server" DataSourceID="SqlDataSource1" DataTextField="State_Name"
                        DataValueField="State_Name">
                    </asp:DropDownList>
                </td>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Lab2" runat="server" Text="申請日期：" CssClass="form" ></asp:Label>
                </td>
			    <td noWrap style="height: 25px; width: 50%;">
                    <asp:TextBox ID="Sdate" runat="server" Width="100px" OnKeyDown="return false" ></asp:TextBox>
                    <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />
                    <asp:Label ID="Lab3" runat="server" Text="~" ></asp:Label>
                    <asp:TextBox ID="Edate" runat="server" Width="100px" OnKeyDown="return false"></asp:TextBox>&nbsp;<asp:ImageButton
                        ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" /></td>
			    <td noWrap style="height: 25px; width: 10%;" align="center">
                    <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />
                    </td>
		    </tr>
	    </table>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" CellPadding="4"
            DataKeyNames="EFORMSN" DataSourceID="SqlDataSource2" EmptyDataText="查無任何資料" ForeColor="#333333"
            GridLines="None" Width="100%" AllowPaging="True" AllowSorting="True">
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <Columns>
                <asp:BoundField DataField="nAPPTIME" HeaderText="申請時間" SortExpression="nAPPTIME" >
                    <HeaderStyle HorizontalAlign="Center" Width="35%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nTYPE" HeaderText="假別" SortExpression="nTYPE" >
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nSTARTTIME" HeaderText="起始日期" SortExpression="nSTARTTIME" >
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nSTHOUR" HeaderText="時" SortExpression="nSTHOUR" >
                    <HeaderStyle HorizontalAlign="Center" Width="5%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nENDTIME" HeaderText="結束日期" SortExpression="nENDTIME" >
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nETHOUR" HeaderText="時" SortExpression="nETHOUR" >
                    <HeaderStyle HorizontalAlign="Center" Width="5%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:ImageButton ID="btnImgCancel" runat="server" ImageUrl="~/Image/cancel.gif" CommandName="Select" OnClientClick="return confirm('請假單表單確定銷假嗎?')" ToolTip="銷假" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="5%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:BoundField DataField="EFORMSN" HeaderText="EFORMSN" ReadOnly="True" SortExpression="EFORMSN" />
            </Columns>
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT P_NUM,EFORMSN,PAIDNO,nAPPTIME,nTYPE,CONVERT (char(12), nSTARTTIME, 111) AS nSTARTTIME,nSTHOUR,CONVERT (char(12), nENDTIME, 111) AS nENDTIME,nETHOUR,nREASON,PENDFLAG FROM P_01 WHERE PENDFLAG='E' AND (EFORMSN NOT IN (SELECT nEFORMSN FROM P_0101)) AND PAIDNO=@user_id ORDER BY nAPPTIME DESC">
            <SelectParameters>
                <asp:SessionParameter Name="user_id" SessionField="user_id" />
            </SelectParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="select * from SYSKIND where Kind_Num = 2 ORDER BY State_Num"></asp:SqlDataSource>
            
        <div id="Div_grid" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:165pt; height:150pt; left:196px; top:985px; display:block;" visible="false">
         
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
        
        <div id="Div_grid2" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:165pt; height:150pt; left:434px; top:986px; display:block;" visible="false">
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
