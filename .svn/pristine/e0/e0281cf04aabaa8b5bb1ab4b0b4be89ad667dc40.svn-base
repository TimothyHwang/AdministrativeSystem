<%@ Page Language="VB" EnableEventValidation = "false" AutoEventWireup="true" CodeFile="MOA08003.aspx.vb" Inherits="M_Source_08_MOA08003" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>影印紀錄查詢與登載</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <script type="text/JavaScript">
        function Check() {
            var re = /</;
            if (re.test(document.getElementById("Sdate").value)) {
                alert("您輸入的查詢開始日期有包含危險字元！");
                document.getElementById("Sdate").value = "";
                return false;
            }

            if (re.test(document.getElementById("Edate").value)) {
                alert("您輸入的查詢結束日期有包含危險字元！");
                document.getElementById("Edate").value = "";
                return false;
            }

            if (re.test(document.getElementById("tbPrinter_No").value)) {
                alert("您輸入的查詢機器號碼有包含危險字元！");
                document.getElementById("tbPrinter_No").value = "";
                return false;
            }


            if (re.test(document.getElementById("tbFile_Name").value)) {
                alert("您輸入的文件名稱有包含危險字元！");
                document.getElementById("tbFile_Name").value = "";
                return false;
            }

            if (re.test(document.getElementById("tb_Print_Name").value)) {
                alert("您輸入的姓名有包含危險字元！");
                document.getElementById("tb_Print_Name").value = "";
                return false;
            }

            if (re.test(document.getElementById("tb_PID").value)) {
                alert("您輸入的人員帳號有包含危險字元！");
                document.getElementById("tb_PID").value = "";
                return false;
            }

            return true;
        }
</script>
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="影印紀錄查詢與登載" Width="100%"></asp:Label>
            </td></tr>
        </table>
        <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
            <tr>
                <td noWrap width="5%" class="form" align = "left">
                    <asp:Label ID="Lab1" runat="server" Text="列印日期：" />
                </td>
                <td noWrap width="20%" class="form"  align = "left">
                     <asp:TextBox ID="Sdate" runat="server" Width="80px" OnKeyDown="return false" ></asp:TextBox>
                    <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />
                    <asp:Label ID="Lab3" runat="server" Text="~" ></asp:Label>
                    <asp:TextBox ID="Edate" runat="server" Width="80px" OnKeyDown="return false"></asp:TextBox>
                    <asp:ImageButton ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" />
                </td>
                <td noWrap width="8%" class="form"  align = "left">
                    <asp:Label ID="Label4" runat="server" Text="保密區分：" ></asp:Label>
                </td>
                <td noWrap width="30%" class="form" align = "left">
                    <asp:DropDownList ID="ddl_Security_Status" runat="server">
                        <asp:ListItem Value="0">請選擇</asp:ListItem>
                        <asp:ListItem Value="1">普通</asp:ListItem>
                        <asp:ListItem Value="2">密</asp:ListItem>
                        <asp:ListItem Value="3">機密</asp:ListItem>
                        <asp:ListItem Value="4">極機密</asp:ListItem>
                        <asp:ListItem Value="5">絕對機密</asp:ListItem>
                    </asp:DropDownList>
                </td>
              
                <td noWrap width="8%" class="form"  align = "left">
                    <asp:Label ID="Label3" runat="server" Text="文件名稱：" ></asp:Label>
                </td>
                <td noWrap width="20%" class="form">
                    <asp:TextBox ID="tbFile_Name" runat="server" MaxLength='100' Width="200px" />
                </td>
                </tr>
                <tr>
                 <td noWrap width="6%" class="form"  align = "left">
                    <asp:Label ID="lbORG" runat="server" Text="單位：" ></asp:Label>
                </td>
                 <td noWrap width="15%" class="form" align = "left">
                     <asp:DropDownList id="ddl_ORG_NAME" DataSourceID="sds_admingroup" DataValueField="ORG_UID"
                        DataTextField="ORG_NAME" runat="server" AppendDataBoundItems="True" >
                        <asp:ListItem Value="0">請選擇</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td noWrap width="8%" class="form"  align = "left">
                    <asp:Label ID="Lab2" runat="server" Text="機器號碼：" ></asp:Label>
                </td>
                <td noWrap width="6%" class="form"  align = "left">
                    <asp:TextBox ID="tbPrinter_No" runat="server" MaxLength='6' Width="60px" />
                </td> 
                <td noWrap width="8%" class="form"  align = "left">
                    <asp:Label ID="Label5" runat="server" Text="機器名稱：" ></asp:Label>
                </td>
                <td noWrap width="6%" class="form"  align = "left">
                    <asp:TextBox ID="tbPrinter_Name" runat="server" MaxLength='30' Width="200px" />
                </td> 
            </tr>
            <tr>
                <td noWrap width="5%" class="form" align = "left">
                    <asp:Label ID="Label2" runat="server" Text="姓名：" ></asp:Label>
                </td>
                <td noWrap width="10%" class="form" align = "left">
                    <asp:TextBox ID="tb_Print_Name" runat="server" MaxLength='10' Width="100px" />
                </td>
                 <td noWrap width="5%" class="form" align = "left">
                    <asp:Label ID="Label1" runat="server" Text="人員帳號：" ></asp:Label>
                </td>
                <td noWrap width="5%" class="form" align = "left">
                    <asp:TextBox ID="tb_PID" runat="server" MaxLength='10' Width="100px" />
                </td>
               
                <td class="form" align = "right" colspan ="2">
                    <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢"  OnClientClick ="return Check();" />
                    <asp:ImageButton ID="ImgClearSearch" runat="server" ImageUrl="~/Image/ClearAll.gif"  ToolTip="清除查詢條件" OnClientClick ="return Check();"  />
                    <asp:ImageButton ID="ImgExport" runat="server" ImageUrl="~/Image/ExportFile.gif" ToolTip="匯出批核表格"  OnClientClick ="return Check();"  />
                    <asp:ImageButton ID="ImgPrint" runat="server" ImageUrl="~/Image/print.gif" 
                        ToolTip="列印批核表格" OnClientClick ="return Check();" Visible="False"  />
                </td>
            </tr>
            <tr>
                <td colspan ="8" align ="left"><asp:Label ID="ErrMsg" runat="server" ForeColor='Red' Font-Bold ="true" /></td>
            </tr>
        </table>

         <asp:SqlDataSource ID="sds_admingroup" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME">
        </asp:SqlDataSource>

        <asp:GridView ID="GV_NewLog" runat="server" EmptyDataText="查無任何資料" 
            AutoGenerateColumns="False" AllowPaging="True"
         AllowSorting="True" DataSourceID="sqlPrintLog" Width="100%" 
            CellPadding="10" ForeColor="#333333" GridLines="None" 
            DataKeyNames="Log_Guid" EnableModelValidation="True">

            <Columns>
                <asp:BoundField DataField="Log_Guid" HeaderText="No" Visible = "false" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                     <HeaderStyle HorizontalAlign="Center" Width="5%" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False" HeaderText="分類" >
                 <ItemTemplate>
                      <%# ShowPrintType(Eval("Copy_A3C"), Eval("Copy_A4C"), Eval("Copy_A3M"), Eval("Copy_A4M"), Eval("Scan"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                </asp:TemplateField>
                 <asp:TemplateField ShowHeader="False" HeaderText="密等" >
                    <ItemTemplate>
                      <%# ShowSecurity_Status(Eval("Security_status"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="5%" />
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                </asp:TemplateField>
                <asp:BoundField DataField="ORG_NAME" HeaderText="部門名稱"  >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="Print_Name" HeaderText="姓名" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                </asp:BoundField>
                <asp:BoundField DataField="Printer_Name" HeaderText="機器名稱" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False" HeaderText="人員帳號" Visible="false" >
                   <ItemTemplate>
                       <asp:Label ID="lb_employee_id" runat="server" Text='<%# Eval("PAIDNO")%>' />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                </asp:TemplateField>
                <asp:BoundField DataField="Print_Date" HeaderText="列印日期時間" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False" HeaderText="狀態" >
                    <ItemTemplate>
                        <%# ShowStatus(Eval("Status")) %> 
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="批核">
                    <ItemTemplate>
                        <%# ShowApproveStatus(Eval("VerifyRequesterID"), Eval("ApprovedByID"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField ShowHeader="False" HeaderText="影印明細">
                    <ItemTemplate>
                        <asp:HiddenField ID="HF_Status" runat="server" Value = '<%# Eval("Status")%>' />
                        <asp:ImageButton ID="ImgDetail" runat="server" CausesValidation="False" CommandArgument ='<%# Eval("Log_Guid") %>' CommandName="Detail"
                          Enabled ='<%# ShowCheckStatus(Eval("Status")) %>'  ImageUrl="~/Image/List.gif" />
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
        <asp:SqlDataSource ID="sqlPrintLog" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>" />
        
        <asp:GridView ID="GV_Export" runat="server"
            AutoGenerateColumns="False" Width="100%" CssClass="form" Visible = "false" 
            CellPadding="10" ForeColor="#333333" GridLines="None" DataKeyNames="Log_Guid" >
            <Columns>
                <asp:BoundField DataField="File_No" HeaderText="年月時分"  >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                     <HeaderStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="ORG_NAME" HeaderText="使用單位"  >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                  <asp:BoundField DataField="File_Name" HeaderText="複印文件資料名稱" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                 <asp:TemplateField ShowHeader="False" HeaderText="保密區分" >
                    <ItemTemplate>
                      <%# ShowSecurity_Status(Eval("Security_status"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="5%" />
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                </asp:TemplateField>
                 <asp:BoundField DataField="Printer_Name" HeaderText="設備名稱" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" />
                </asp:BoundField>
                  <asp:BoundField DataField="PrintTotalCnt" HeaderText="資料張數" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="5%" />
                </asp:BoundField>
                 <asp:TemplateField ShowHeader="False" HeaderText="紙張型式" >
                    <ItemTemplate>
                      <%# ShowPrint(Eval("Copy_A3C"), Eval("Copy_A4C"), Eval("Copy_A3M"), Eval("Copy_A4M"), Eval("Scan"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                </asp:TemplateField>
                <asp:BoundField DataField="Use_For" HeaderText="用途" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False" HeaderText="級職" >
                    <ItemTemplate>
                      <%# ShowTU_Nam(Eval("TU_ID"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                </asp:TemplateField>
                <asp:BoundField DataField="Print_Name" HeaderText="姓名" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="Useless" HeaderText="作廢張數" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False" HeaderText="狀態" >
                    <ItemTemplate>
                      <%# ShowStatus(Eval("Status"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
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
