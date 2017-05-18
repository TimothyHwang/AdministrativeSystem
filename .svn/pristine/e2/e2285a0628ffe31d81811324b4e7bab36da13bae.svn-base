<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA08007.aspx.vb" Inherits="M_Source_08_MOA08007" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>機密資訊複(影)印資料申請單新增與查詢</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        .form
        {
            margin-top: 1px;
        }
    </style>
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="機密資訊複(影)印資料申請單新增與查詢" Width="100%"></asp:Label>
            </td></tr>
        </table>
        <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
            <tr>
                <td nowrap="nowrap" width="5%" class="form">
                    <asp:Label ID="Lab1" runat="server" Text="登記申請日期：" />
                </td>
                <td nowrap="nowrap" width="20%" class="form">
                     <asp:TextBox ID="Sdate" runat="server" Width="80px" OnKeyDown="return false" ></asp:TextBox>
                    <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />
                    <asp:Label ID="Lab3" runat="server" Text="~" ></asp:Label>
                    <asp:TextBox ID="Edate" runat="server" Width="80px" OnKeyDown="return false"></asp:TextBox>
                    <asp:ImageButton ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" />
                </td>
                <td nowrap="nowrap" width="6%" class="form">
                    <asp:Label ID="Label1" runat="server" Text=" 登記狀況：" ></asp:Label>
                </td>
                <td nowrap="nowrap" width="6%" class="form">
                    <asp:DropDownList id="ddl_Security_Status" AutoPostBack="True" runat="server">
                        <asp:ListItem Value="-1"> 全部 </asp:ListItem>
                        <asp:ListItem Value="0"> 審核中 </asp:ListItem>
                        <asp:ListItem Value="1"> 審核通過 </asp:ListItem>
                        <asp:ListItem Value="2"> 審核不通過 </asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td nowrap="nowrap" width="5%" class="form"  align ="right">
                    <asp:Label ID="lbUserName" runat="server" Text="姓名：" ></asp:Label>
                </td>
                 <td nowrap="nowrap" width="5%" class="form">
                     <asp:TextBox ID="tbUserName" runat="server" MaxLength ="10" Width ="100px" 
                         Font-Size="Medium" />
                </td>
                <td nowrap="nowrap" width="10%" class="form" align = "right">
                    <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif"
                        ToolTip="查詢" />
                    <asp:ImageButton ID="ImBt_Clear" runat="server" ImageUrl="~/Image/ClearAll.gif"
                        ToolTip="清除查詢條件" />
                        <a href="#" onclick="window.location.href='MOA08008.aspx';"><asp:Image ID="ImgAdd" runat="server" border="0" ImageUrl="~/Image/add.gif" ToolTip="新增申請單" /></A> 
                </td>
            </tr>
            <tr>
                <td colspan ="9" align ="left"><asp:Label ID="ErrMsg" runat="server" ForeColor='Red' Text="" /></td>
            </tr>
        </table>
        
          <asp:GridView ID="GV_Security" runat="server" EmptyDataText="查無任何資料" 
            AutoGenerateColumns="False" AllowPaging="True"
         AllowSorting="True" DataSourceID="sqlSecurityLog" Width="100%" 
            CellPadding="10" ForeColor="#333333" GridLines="None" 
            DataKeyNames="Guid_ID" PageSize="5" EnableModelValidation="True">

            <Columns>                
                <asp:BoundField DataField="Guid_ID" HeaderText="流水號" Visible="false"  >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                     <HeaderStyle HorizontalAlign="Center" Width="5%" />
                </asp:BoundField>
                <asp:BoundField DataField="Top1unitName" HeaderText="申請單位全銜"  >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="emp_chinese_name" HeaderText="申請人姓名"  >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="Subject" HeaderText="主旨/簡由" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                </asp:BoundField>
                 <asp:TemplateField HeaderText="機密等級" >
                    <ItemTemplate>
                      <%# ShowSecurity_Level(Eval("Security_Level"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="機密屬性" >
                    <ItemTemplate>
                      <%# ShowSecurity_Type(Eval("Security_Level"),Eval("Security_Type"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                </asp:TemplateField>
                 <asp:TemplateField HeaderText="用途" >
                   <ItemTemplate>
                        <%# ShowPurpose(Eval("Purpose"), Eval("Purpose_Other"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                </asp:TemplateField>
                 <asp:BoundField DataField="Security_DateTime" HeaderText="填單存檔列印時間" >
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                </asp:BoundField>
                 <asp:TemplateField HeaderText="申請單狀態" >
                    <ItemTemplate>
                      <%# ShowSecurity_Status(Eval("Guid_ID"), Eval("Security_Status"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                </asp:TemplateField>       
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                    <a  href="javascript:void(0);" onclick="javascript:window.open('../08/ConfidentialPhotoCopy.aspx?id=<%#Eval("Guid_ID") %>');"><img alt="" src="../../Image/print.gif"  border="0" /></a>
                    &nbsp<asp:ImageButton ID="LinkBtnDel" runat="server" CausesValidation="false" CommandName="Delete" OnClientClick="return confirm('確認要刪除嗎？');" ImageUrl="~/Image/delete.gif" />
                    </ItemTemplate>
                     <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                </asp:TemplateField>
                <asp:TemplateField Visible="false">
                    <ItemTemplate>
                        <asp:HiddenField ID="hid_guid" runat="server" Value = '<%# Eval("Guid_ID")%>'  />
                        <asp:HiddenField ID="hid_PAIDNO" runat="server" Value = '<%# Eval("PAIDNO")%>'  />
                        <asp:HiddenField ID="hid_Security_Level" runat="server" Value = '<%# Eval("Security_Level")%>'  />
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

        <asp:SqlDataSource ID="sqlSecurityLog" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>" 
            DeleteCommand="DELETE FROM [P_0804] WHERE [Guid_ID] = @Guid_ID"             
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
            <DeleteParameters>
                <asp:Parameter Name="Guid_ID" Type="String" />
            </DeleteParameters>           
        </asp:SqlDataSource>  
        
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
