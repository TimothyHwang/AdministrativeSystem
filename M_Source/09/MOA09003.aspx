<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA09003.aspx.vb" Inherits="M_Source_09_MOA09003" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>門禁會議管制統計</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.7.2.min.js" type="text/javascript"></script>
    <script type="text/JavaScript">
        function ClearData() {
            $("#<%=txtMeetingDateStart.ClientID %>").val("");
            $("#<%=txtMeetingDateEnd.ClientID %>").val("");
            $("#<%=ddlDepartment.ClientID %>").val("");
            return false;
        }
    </script>
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="門禁會議管制統計" Width="100%"></asp:Label>
            </td></tr>
        </table>
        <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
            <tr>
                <td nowrap="nowrap" width="5%" class="form" align = "left">
                    <asp:Label ID="Lab2" runat="server" Text="開會日期：" />                    
                </td>
                <td nowrap="nowrap" width="20%" class="form"  align = "left">
                     <asp:TextBox ID="txtMeetingDateStart" runat="server" Width="80px" 
                         OnKeyDown="return false" ></asp:TextBox>
                    <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />
                    <asp:Label ID="Lab3" runat="server" Text="~" ></asp:Label>
                    <asp:TextBox ID="txtMeetingDateEnd" runat="server" Width="80px" 
                         OnKeyDown="return false"></asp:TextBox>
                    <asp:ImageButton ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" />
                </td>
                <td nowrap="nowrap" width="8%" class="form"  align = "left">
                    <asp:Label ID="Lab1" runat="server" Text="部門名稱：" />
                </td>
                <td nowrap="nowrap" width="30%" class="form" align = "left">
                    <asp:DropDownList ID="ddlDepartment" runat="server">
                    </asp:DropDownList>
                </td>
                <td nowrap="nowrap" width="8%" class="form"  align = "left">
                    &nbsp;</td>
                <td nowrap="nowrap" width="20%" class="form">
                    <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />
                    &nbsp;<asp:ImageButton ID="ImgClearSearch" runat="server" ImageUrl="~/Image/ClearAll.gif"  ToolTip="清除查詢條件" OnClientClick ="return ClearData();"  />
                </td>
                </tr>
        </table>

        <asp:GridView ID="gvP_09" runat="server" EmptyDataText="查無任何資料" 
            AutoGenerateColumns="False" AllowPaging="True"
         AllowSorting="True" Width="100%" 
            CellPadding="10" ForeColor="#333333" GridLines="None" 
            EnableModelValidation="True">

            <Columns>
                <asp:BoundField DataField="Sponsor" HeaderText="承辦部門名稱">
                <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="NumberOfMeetings" HeaderText="會議次數">
                <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="NumberOfPeople" HeaderText="人員加總" >
                <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:ImageButton ID="ImgDetail" runat="server" CausesValidation="False" CommandArgument ='<%# Eval("Sponsor") %>' CommandName="Detail" ImageUrl="~/Image/List.gif" />
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
    </form>
</body>
</html>
