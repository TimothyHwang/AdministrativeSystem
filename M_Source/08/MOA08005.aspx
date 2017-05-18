<%@ Page Language="VB" EnableEventValidation = "false" AutoEventWireup="true"  CodeFile="MOA08005.aspx.vb" Inherits="M_Source_08_MOA08005" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>影印紀錄統計</title>
      <link href="../../Styles.css" rel="stylesheet" type="text/css" />
      <script type="text/JavaScript">
          function Check(type) {
              var re = /</;
              if (type == 1) {
                  if (document.getElementById("tbnPrinter_No").value != "") {
                      if (re.test(document.getElementById("tbnFacilityNo").value)) {
                          alert("您輸入的設備編碼有包含危險字元！");
                          document.getElementById("tbnFacilityNo").value = "";
                          return false;
                      }

                      if (document.getElementById("tbnPrinter_No").value.lenght != 6) {
                          alert("查詢輸入的設備編碼與物料名稱請勿2者都空白！");
                          return false;
                      }
                  }
              }
              return true;
          }
</script>
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="影印紀錄統計" Width="100%"></asp:Label>
            </td></tr>
        </table>
        <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
            <tr>
                <td width="45%" class="form">
                 <asp:Label ID="Lab1" runat="server" Text="日期區間：" />
                   <asp:TextBox ID="Sdate" runat="server" Width="80px" OnKeyDown="return false" ></asp:TextBox>
                    <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />
                    <asp:Label ID="Lab3" runat="server" Text="~" ></asp:Label>
                    <asp:TextBox ID="Edate" runat="server" Width="80px" OnKeyDown="return false"></asp:TextBox>
                    <asp:ImageButton ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" />
                </td>
                 <td noWrap width="5%" class="form">
                    <asp:Label ID="Label1" runat="server" Text="查詢分類：" ></asp:Label>
                </td>
                <td noWrap width="10%" class="form">
                    <asp:DropDownList id="ddl_QueryType" AutoPostBack="True" runat="server">
                        <asp:ListItem Value="1"> 影印申請狀態 </asp:ListItem>
                        <asp:ListItem Value="2"> 印表機機型號碼 </asp:ListItem>
                        <asp:ListItem Value="3"> 保密區分 </asp:ListItem>
                        <asp:ListItem Value="4"> 歷程查詢 </asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td align="right">
                <asp:Label ID="ErrMsg" runat="server" ForeColor='Red' Font-Bold="True" Font-Size="Small" />
                </td>
                <td align = "right"><asp:ImageButton ID="Img_Search" runat="server" 
                        ImageUrl="~/Image/search.gif" />
                    <asp:ImageButton ID="Img_Export" runat="server" 
                        ImageUrl="~/Image/ExportFile.gif" Enabled="False" /></td>
            </tr>
        </table>
        <center>
        <asp:GridView ID="gvPrintStatus" runat="server" AutoGenerateColumns="False" ShowFooter = "true"  Visible = "false"
        CellPadding="4" EmptyDataText="查無任何資料" ForeColor="#333333" GridLines="None" Width="500px">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EditRowStyle BackColor="#999999" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <FooterStyle BackColor="#5D1234" Font-Bold="True" ForeColor="yellow"  HorizontalAlign="Center"  />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <Columns>
                 <asp:BoundField  DataField="nStatusName" HeaderText="狀態">
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:BoundField  DataField="Anycnt" HeaderText="筆數" >
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:LinkButton ID="LBPrintStatusDetail" runat="server" CausesValidation="False" CommandArgument=<%#Eval("Status") %> CommandName="AnyDetail" text="明細" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="10%" />
                </asp:TemplateField> 
            </Columns>
        </asp:GridView>	
        <asp:GridView ID="gvPrinterNO" runat="server" AutoGenerateColumns="False" ShowFooter = "true"  Visible = "false"
        CellPadding="4" EmptyDataText="查無任何資料" ForeColor="#333333" GridLines="None" Width="500px">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EditRowStyle BackColor="#999999" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <FooterStyle BackColor="#5D1234" Font-Bold="True" ForeColor="yellow"  HorizontalAlign="Center"  />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <Columns>
                 <asp:BoundField  DataField="Printer_No" HeaderText="影印機機型號碼">
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:BoundField  DataField="Anycnt" HeaderText="筆數" >
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:LinkButton ID="LBPrinter_NoDetail" runat="server" CausesValidation="False" CommandArgument=<%#Eval("Printer_No") %> CommandName="AnyDetail" text="明細" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="10%" />
                </asp:TemplateField> 
            </Columns>
        </asp:GridView>	
        <asp:GridView ID="gvSecurityStatus" runat="server" AutoGenerateColumns="False" ShowFooter = "true"  Visible = "false"
        CellPadding="4" EmptyDataText="查無任何資料" ForeColor="#333333" GridLines="None" Width="500px">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EditRowStyle BackColor="#999999" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <FooterStyle BackColor="#5D1234" Font-Bold="True" ForeColor="yellow"  HorizontalAlign="Center"  />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <Columns>
                 <asp:BoundField  DataField="Security_Status_Name" HeaderText="保密區分">
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:BoundField  DataField="Anycnt" HeaderText="筆數" >
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:LinkButton ID="LBSecurity_StatusDetail" runat="server" CausesValidation="False" CommandArgument=<%#Eval("Security_Status") %> CommandName="AnyDetail" text="明細" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="10%" />
                </asp:TemplateField> 
            </Columns>
        </asp:GridView>
        <asp:GridView ID="gvHistory" runat="server" AutoGenerateColumns="False" ShowFooter = "true"  Visible = "false"
        CellPadding="4" EmptyDataText="查無任何資料" ForeColor="#333333" GridLines="None" Width="500px">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EditRowStyle BackColor="#999999" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <FooterStyle BackColor="#5D1234" Font-Bold="True" ForeColor="yellow"  HorizontalAlign="Center"  />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <Columns>
                 <asp:BoundField  DataField="movementName" HeaderText="操作動作">
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:BoundField  DataField="Anycnt" HeaderText="筆數" >
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:LinkButton ID="LBmovementDetail" runat="server" CausesValidation="False" CommandArgument=<%#Eval("movement") %> CommandName="AnyDetail" text="明細" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="10%" />
                </asp:TemplateField> 
            </Columns>
        </asp:GridView>

        </center>
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
