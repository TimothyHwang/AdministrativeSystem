<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA09002.aspx.vb" Inherits="M_Source_09_MOA09002" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>門禁會議管制查詢</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.7.2.min.js" type="text/javascript"></script>
    <script type="text/JavaScript">
        function ClearData() {
            $("#<%=txtMeetingDateStart.ClientID %>").val("");
            $("#<%=txtMeetingDateEnd.ClientID %>").val("");
            return false;
        }
    </script>
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="門禁會議管制查詢" Width="100%"></asp:Label>
            </td></tr>
        </table>
        <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
            <tr>
                <td nowrap="nowrap" width="5%" class="form" align = "left">
                    <asp:Label ID="Lab1" runat="server" Text="開會日期：" />
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
                    &nbsp;</td>
                <td nowrap="nowrap" width="30%" class="form" align = "left">
                    &nbsp;</td>
              
                <td nowrap="nowrap" width="8%" class="form"  align = "left">
                    &nbsp;</td>
                <td nowrap="nowrap" width="20%" class="form">
                    <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />
                    &nbsp;<asp:ImageButton ID="ImgClearSearch" runat="server" ImageUrl="~/Image/ClearAll.gif"  ToolTip="清除查詢條件" OnClientClick ="return ClearData();"  />
                    <asp:ImageButton ID="imbOutSearch" runat="server" ImageUrl="~/Image/ReadAD.gif" 
                        ToolTip="查詢未匯出資料" />
                    <asp:ImageButton ID="imbExport" runat="server" 
                        ImageUrl="~/Image/mend.gif" ToolTip="匯出資料" Visible="False" onClientClick="ShowProgressBar();"/>
                </td>
                </tr>
        </table>

        <asp:GridView ID="gvP_09" runat="server" EmptyDataText="查無任何資料" 
            AutoGenerateColumns="False" AllowPaging="True"
         AllowSorting="True" Width="100%" 
            CellPadding="10" ForeColor="#333333" GridLines="None" 
            EnableModelValidation="True">

            <Columns>
                <asp:BoundField DataField="eFormSN" HeaderText="eFormSN" ></asp:BoundField>
                <asp:TemplateField HeaderText="開會日期" >
                    <ItemTemplate><%# ShowMeetingDate(Eval("MeetingDate"))%></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="進出營門">
                    <ItemTemplate><%# ShowEnteringGate(Eval("EFORMSN"))%></ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="Subject" HeaderText="開會事由" />
                <asp:BoundField DataField="Location" HeaderText="地點"></asp:BoundField>
                <asp:TemplateField HeaderText="進出人員">
                    <ItemTemplate><%# Eval("EnteringPeopleNumber")%> 員</ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="Moderator" HeaderText="主持人" ></asp:BoundField>
                <asp:TemplateField HeaderText="聯絡人及電話">
                    <ItemTemplate>
                        <table border="0" width="100%">
                            <tr><td width="100%" align="center"><%# Eval("ORG_NAME")%></td></tr>
                            <tr><td width="100%" align="center"><%# Eval("emp_chinese_name")%>&nbsp;<%# Eval("PhoneNumber")%></td></tr>
                        </table>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="DocumentNo" HeaderText="發文字號" ></asp:BoundField>
                <asp:TemplateField HeaderText="狀態" >
                    <ItemTemplate><%# ShowStatus(Eval("Status"))%></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:ImageButton ID="ImgDetail" runat="server" CausesValidation="False" CommandArgument ='<%# Eval("EFORMSN") %>' CommandName="Detail" ImageUrl="~/Image/List.gif" />
                    </ItemTemplate>
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
            <div id="divProgress" style="text-align:center; display: none; position: fixed; top: 50%;  left: 50%;" > 
     <asp:Image ID="imgLoading" runat="server" ImageUrl="~/Image/loading.gif" />          
     <br /> 
     <font color="#1B3563" size="2px">資料處理中</font> 
 </div> 
 <div id="divMaskFrame" style="background-color: #F2F4F7; display: none; left: 0px; 
     position: absolute; top: 0px;"> 
 </div> 
    </form>
    <script type="text/javascript" language="javascript">

        // 顯示讀取遮罩 
        function ShowProgressBar() {
            displayProgress();
            displayMaskFrame();
        }
        // 隱藏讀取遮罩 
        function HideProgressBar() {
            var progress = $('#divProgress');
            var maskFrame = $("#divMaskFrame");
            progress.hide();
            maskFrame.hide();
        }
        // 顯示讀取畫面 
        function displayProgress() {
            var w = $(document).width();
            var h = $(window).height();
            var progress = $('#divProgress');
            progress.css({ "z-index": 999999, "top": (h / 2) - (progress.height() / 2), "left": (w / 2) - (progress.width() / 2) });
            progress.show();
        }
        // 顯示遮罩畫面 
        function displayMaskFrame() {
            var w = $(window).width();
            var h = $(document).height();
            var maskFrame = $("#divMaskFrame");
            maskFrame.css({ "z-index": 999998, "opacity": 0.7, "width": w, "height": h });
            maskFrame.show();
        } 
    </script>
</body>
</html>
