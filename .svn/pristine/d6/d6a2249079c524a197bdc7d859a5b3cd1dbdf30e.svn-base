<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04022.aspx.vb" Inherits="Source_04_MOA04022" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>倉儲領料功能</title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
    <script type="text/javascript" language="javascript" src="../Inc/inc_common.js"></script>
    <script type="text/JavaScript">
        function Check() {
            var re = /</;
            if (re.test(document.getElementById("tb_eFormSn").value)) {
                document.getElementById("tbnFacilityNo").value = "";
                alert("您輸入的報修單單號有包含危險字元！");
                return false;
            }
            return true;
        }
</script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center">
                    <asp:Label ID="lblTitle" runat="server" CssClass="toptitle" Text="倉儲領料功能" Width="100%"></asp:Label>
                    
                </td>
            </tr>
        </table>
        <asp:Panel ID="plConditionsContainer" runat="server">
            <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
                <tr>
                    <td nowrap class="form" style="height: 25px; width: 10%;">
                        <asp:Label ID="Label1" runat="server" CssClass="form" Text="報修單單號:" Width="77px"></asp:Label>
                    </td>
                    <td nowrap width="15%" class="form" style="height: 25px"> 
                        <asp:TextBox ID="tb_eFormSn" runat="server" MaxLength="16" Width="150px"></asp:TextBox>
                    </td>
                    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Lab2" runat="server" Text="申請領料日期：" CssClass="form" ></asp:Label>
                </td>
			    <td noWrap style="height: 25px; width: 35%;">
                    <asp:TextBox ID="Sdate" runat="server" Width="80px" OnKeyDown="return false" ></asp:TextBox>
                    <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />
                    <asp:Label ID="Lab3" runat="server" Text="~" ></asp:Label>
                    <asp:TextBox ID="Edate" runat="server" Width="80px" OnKeyDown="return false"></asp:TextBox>
                    <asp:ImageButton ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" />
                    </td>
                      <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Label2" runat="server" Text="領料完結狀態：" CssClass="form" ></asp:Label></td> 
                     <td noWrap style="height: 25px; width: 25%;">
                         <asp:DropDownList ID="ddl_nAppStockStatus" runat="server">
                             <asp:ListItem Value="0">全部</asp:ListItem>
                             <asp:ListItem Value="Y">已完結</asp:ListItem>
                             <asp:ListItem Value="Null">未完結</asp:ListItem>
                         </asp:DropDownList>
                   </td>
                     <td nowrap width="1%" class="form" style="padding-left: 50px;">
                      <asp:ImageButton ID="ibtclearall" runat="server" ImageUrl="~/Image/clearall.gif" 
                             ToolTip="清除" Height="22px" />
                        <asp:ImageButton ID="ibtnSearch" runat="server" ImageUrl="~/Image/search.gif" 
                             ToolTip="查詢" OnClientClick = "return Check();" Height="22px" />
                    </td>
                </tr>
                <tr><td colspan="5" style="height: 10px;"></td></tr>
            </table>
        </asp:Panel>
        <asp:GridView ID="gvDataRecords" runat="server" AllowPaging="True" AllowSorting="True" AutoGenerateColumns="False" CellPadding="4" 
        DataKeyNames="EFORMSN" DataSourceID="sdsDataRecords" EmptyDataText="查無任何資料" ForeColor="#333333" GridLines="None" Width="100%">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EditRowStyle BackColor="#999999" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <Columns>
                <asp:BoundField DataField="EFORMSN" SortExpression="EFORMSN" HeaderText="報修單單號" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="left" />
                    <ItemStyle HorizontalAlign="left" Wrap="False" />
                </asp:BoundField>
                <asp:BoundField DataField="location" SortExpression="location" HeaderText="報修地點" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="left" />
                    <ItemStyle HorizontalAlign="Left" Wrap="false" />
                </asp:BoundField>
               <asp:BoundField DataField="nAPPTIME" SortExpression="nAPPTIME" HeaderText="申請日期" ReadOnly="True">
                    <HeaderStyle HorizontalAlign="left" />
                    <ItemStyle HorizontalAlign="Left" Wrap="false" />
                </asp:BoundField>
                 <asp:TemplateField HeaderText="領料完結">
                    <HeaderStyle HorizontalAlign="left" />
                    <ItemStyle HorizontalAlign="left" Wrap="False" />
                    <ItemTemplate>
                    <A HREF="#" onClick="window.open('<%# "MOA04022_1.aspx?EFORMSN=" + Eval("EFORMSN")%>', '領料品項資訊', config='height=800,width=1000,scrollbars=yes,status=yes')"> <asp:Label ID="Label3" runat="server" Text="領料明細" /></A> 
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
        
    </div>
    <asp:SqlDataSource ID="sdsDataRecords" runat="server"
        ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
        SelectCommand="select EFORMSN,b.bd_name+'/'+c.fl_name+'/'+d.rnum_name as location,nAPPTIME,nAppStockStatus
                        from P_0415 a left join P_0404 b on a.nbd_code = b.bd_code
                        left join P_0406 c on a.nfl_code = c.fl_code
                        left join P_0411 d on a.nrnum_code = d.rnum_code
                        where (1=1) Order by nAppStockStatus;">
    </asp:SqlDataSource>

              
        <div id="Div_grid" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:165pt; height:150pt; left:402px; top:640px; display:block;" visible="false">
         
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