<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA01009.aspx.vb" Inherits="Source_04_MOA01009" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>文職人員加班統計表</title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
    <script type="text/javascript" language="javascript">
        document.oncontextmenu = new Function("return false");
    </script>
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
    <div>
        <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center">
                    <asp:Label ID="lblHeader" runat="server" CssClass="toptitle" Text="文職人員加班統計表" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>
        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="text-align: center;">
            <tr>
                <td style="width: 1%; white-space: nowrap;">
                    <asp:Label ID="Label1" runat="server" Text="單位:" CssClass="form"></asp:Label>
                </td>
                <td style="width: 33%; height: 25px;">
                    <asp:DropDownList ID="ddl_org_picker" runat="server" AutoPostBack="true" DataSourceID="sdsOrganization" DataTextField="ORG_NAME" DataValueField="ORG_UID" Width="90%"></asp:DropDownList>
                </td>
                <td style="width: 1%; white-space: nowrap; text-align: right;">
                    <asp:Label ID="Label2" runat="server" Text="姓名:" CssClass="form"></asp:Label>
                </td>
                <td style="width: 33%; height: 25px;">
                    <asp:DropDownList ID="ddl_user_picker" runat="server" DataSourceID="sdsUsersList" DataTextField="emp_chinese_name" DataValueField="employee_id" Width="90%"></asp:DropDownList>
                </td>
                <td style="width: 1%; white-space: nowrap; text-align: right;">
                    <asp:Label ID="Label3" runat="server" Text="加班:" CssClass="form"></asp:Label>
                </td>
                <td style="width: 33%; height: 25px;">
                    <asp:DropDownList ID="ddl_overtime" runat="server" Width="90%">
                        <asp:ListItem>17:00</asp:ListItem>
                        <asp:ListItem>18:00</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td style="width: 1%; white-space: nowrap;">
                    <asp:Label ID="Label4" runat="server" Text="日期:" CssClass="form"></asp:Label>
                </td>
                <td align="center" style="width: 33%; height: 25px;">
                    <table border="0" cellpadding="0" cellspacing="0" width="90%">
                        <tr>
                            <td align="left" style="width: 49%;">
                                <asp:TextBox ID="tb_begin_date" runat="server" Width="92%" OnKeyDown="return false"></asp:TextBox>
                            </td>
                            <td style="width: 1%; white-space: nowrap;">
                                <asp:ImageButton ID="ibtn_begin_date" runat="server" ImageUrl="~/Image/calendar.gif" />
                            </td>
                            <td>&nbsp;~&nbsp;</td>
                            <td style="width: 49%; padding-right: 2px;">
                                <asp:TextBox ID="tb_end_date" runat="server" Width="92%" OnKeyDown="return false"></asp:TextBox>
                            </td>
                            <td style="width: 1%; white-space: nowrap;">
                                <asp:ImageButton ID="ibtn_end_date" runat="server" ImageUrl="~/Image/calendar.gif" />
                            </td>
                        </tr>
                    </table>
                </td>
                <td style="width: 1%; white-space: nowrap;">
                    <asp:Label ID="Label5" runat="server" Text="簽到狀態:" CssClass="form"></asp:Label>
                </td>
                <td style="width: 33%; height: 25px;">
                    <asp:DropDownList ID="ddl_sign_in" runat="server" Width="90%">
                        <asp:ListItem>已簽到</asp:ListItem>
                        <asp:ListItem>未簽到</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td style="width: 1%; white-space: nowrap;">
                    <asp:Label ID="Label6" runat="server" Text="簽退狀態:" CssClass="form"></asp:Label>
                </td>
                <td style="width: 33%; height: 25px;">
                    <asp:DropDownList ID="ddl_sign_out" runat="server" Width="90%">
                        <asp:ListItem>已簽退</asp:ListItem>
                        <asp:ListItem>未簽退</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="left">
                    <asp:Button ID="btnNoonSignIn" runat="server" Text="中午加班簽到" />
                    <asp:Button ID="btnNoonSignOut" runat="server" Text="中午加班簽退" />
                </td>
                <td colspan="3">&nbsp;</td>
                <td align="left" style="padding-top: 8px; padding-bottom: 4px;">
                    <table border="0" cellpadding="0" cellspacing="0" width="95%">
                        <tr>
                            <td align="right" style="white-space: nowrap;">
                                <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />
                                <asp:ImageButton ID="ImgExport" runat="server" ImageUrl="~/Image/ExportFile.gif" ToolTip="匯出檔案" />
                                <asp:ImageButton ID="ImagePrint" runat="server" CommandName="Confirm" ImageUrl="~/Image/print.gif" ToolTip="列印" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True" AutoGenerateColumns="False" CellPadding="4" DataSourceID="sdsSignRecords" EmptyDataText="查無任何資料" ForeColor="#333333" GridLines="None" PageSize="12" Width="100%">
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <Columns>
                <asp:TemplateField HeaderText="人員姓名" SortExpression="EMP_CHINESE_NAME">
                    <HeaderStyle HorizontalAlign="Center" Width="1%" />
                    <ItemStyle HorizontalAlign="Left" Wrap="false" />
                    <ItemTemplate>
                        <asp:Label ID="lbl_emp_chinese_name" runat="server" Text='<%# Bind("EMP_CHINESE_NAME") %>'></asp:Label>
                        <asp:HiddenField ID="hid_emp_chinese_name" runat="server" Value='<%# Bind("EMPLOYEE_ID") %>' />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="SIGNDATE_d" HeaderText="日期" DataFormatString="{0:yyyy/M/d}" SortExpression="SIGNDATE_d">
                    <HeaderStyle HorizontalAlign="Center" Width="1%" />
                    <ItemStyle HorizontalAlign="Center" Wrap="false" />
                </asp:BoundField>
                <asp:BoundField DataField="In_Time_nvc" HeaderText="簽到時間" NullDisplayText="未簽到" SortExpression="In_Time_nvc">
                    <HeaderStyle HorizontalAlign="Center" Width="1%" />
                    <ItemStyle HorizontalAlign="Center" Wrap="false" />
                </asp:BoundField>
                <asp:BoundField DataField="Out_Time_nvc" HeaderText="簽退時間" SortExpression="Out_Time_nvc">
                    <HeaderStyle HorizontalAlign="Center" Width="1%" />
                    <ItemStyle HorizontalAlign="Center" Wrap="false" />
                </asp:BoundField>
                <asp:BoundField DataField="Out_REASON" HeaderText="簽退事由" SortExpression="Out_REASON">
                    <HeaderStyle HorizontalAlign="Center" Width="1%" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="NoonHour" HeaderText="中午加班" SortExpression="NoonHour">
                    <HeaderStyle HorizontalAlign="Center" Width="1%" />
                    <ItemStyle HorizontalAlign="Center" Wrap="false" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="總時數">
                    <HeaderStyle HorizontalAlign="Center" Width="1%" />
                    <ItemStyle HorizontalAlign="Center" Wrap="false" />
                    <ItemTemplate>
                        <asp:Label ID="lbl_hours_difference" runat="server"></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="小計">
                    <HeaderStyle HorizontalAlign="Center" Width="1%" />
                    <ItemStyle HorizontalAlign="Right" Wrap="false" />
                    <ItemTemplate>
                        <div style="padding-right: 4%;">
                            <asp:Label ID="lbl_subtotal" runat="server"></asp:Label>
                        </div>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>

        <div id="div_grid1" runat="server" style="border-right: lightslategray 2px solid; border-top: lightslategray 2px solid; display: block; z-index: 3; left: 193px; border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid; position: absolute; top: 406px; height: 150pt; background-color: white" visible="false">
            <asp:Calendar ID="Calendar1" runat="server" BackColor="White" BorderColor="#3366CC" BorderWidth="1px" Caption="" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana" Font-Size="8pt" ForeColor="#003399" Height="200px" ShowGridLines="True" Width="220px">
                <SelectedDayStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
                <TodayDayStyle BackColor="#99CCCC" ForeColor="White" />
                <SelectorStyle BackColor="#99CCCC" ForeColor="#336666" />
                <WeekendDayStyle BackColor="#CCCCFF" />
                <OtherMonthDayStyle ForeColor="#999999" />
                <NextPrevStyle Font-Size="8pt" ForeColor="#CCCCFF" />
                <DayHeaderStyle BackColor="#99CCCC" ForeColor="#336666" Height="1px" />
                <TitleStyle BackColor="#003399" BorderColor="#3366CC" BorderWidth="1px" Font-Bold="True" Font-Size="10pt" ForeColor="#CCCCFF" Height="25px" />
            </asp:Calendar>
            <asp:Button ID="btnClose1" runat="server" Text="關閉" Width="220px" />
        </div>
        <div id="div_grid2" runat="server" style="border-right: lightslategray 2px solid; border-top: lightslategray 2px solid; display: block; z-index: 3; left: 431px; border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid; position: absolute; top: 407px; height: 150pt; background-color: white" visible="false">
            <asp:Calendar ID="Calendar2" runat="server" BackColor="White" BorderColor="#3366CC" BorderWidth="1px" Caption="" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana" Font-Size="8pt" ForeColor="#003399" Height="200px" ShowGridLines="True" Width="220px">
                <SelectedDayStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
                <TodayDayStyle BackColor="#99CCCC" ForeColor="White" />
                <SelectorStyle BackColor="#99CCCC" ForeColor="#336666" />
                <WeekendDayStyle BackColor="#CCCCFF" />
                <OtherMonthDayStyle ForeColor="#999999" />
                <NextPrevStyle Font-Size="8pt" ForeColor="#CCCCFF" />
                <DayHeaderStyle BackColor="#99CCCC" ForeColor="#336666" Height="1px" />
                <TitleStyle BackColor="#003399" BorderColor="#3366CC" BorderWidth="1px" Font-Bold="True" Font-Size="10pt" ForeColor="#CCCCFF" Height="25px" />
            </asp:Calendar>
            <asp:Button ID="btnClose2" runat="server" Text="關閉" Width="220px" />
        </div>

        <asp:SqlDataSource ID="sdsOrganization" runat="server"
            ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="Select ORG_NAME, ORG_UID From [ADMINGROUP] Where (1 = 2)">
        </asp:SqlDataSource>

        <asp:SqlDataSource ID="sdsUsersList" runat="server"

            ConnectionString="<%$ ConnectionStrings:ConnectionString %>"

            SelectCommand="Select emp_chinese_name, employee_id From EMPLOYEE Where ORG_UID = @ORG_UID And 
                           AD_TITLE Like (@AD_TITLE + '%') And (1 = 1) Order by emp_chinese_name">

            <SelectParameters>
                <asp:Parameter Name="AD_TITLE" DbType="String" Size="50" />
                <asp:ControlParameter ControlID="ddl_org_picker" Name="ORG_UID" Type="String" PropertyName="SelectedValue" />
            </SelectParameters>

        </asp:SqlDataSource>

        <asp:SqlDataSource ID="sdsSignRecords" runat="server"

            ConnectionString="<%$ ConnectionStrings:ConnectionString %>"

            SelectCommand=" Declare @REC_EMPLOYEE_ID VarChar(10) Set @REC_EMPLOYEE_ID = Null; 
                            Declare @REC_EMPLOYEE_CNAME NVarChar(10) Set @REC_EMPLOYEE_CNAME = Null;
                            Declare CUR_EMPLOYEE Cursor For (Select employee_id, emp_chinese_name 
                            From EMPLOYEE Where AD_TITLE Like (@AD_TITLE + '%') And ORG_UID In ('') And 
                            employee_id = IsNull(@employee_id, employee_id));
                            Create Table #EMP_SIGN_RECORDS (EMPLOYEE_ID VarChar(10), EMP_CHINESE_NAME NVarChar(10), SIGN_DATE DateTime); 
                            Declare @Index Int Set @Index = 0; 
                            Declare @Limitation Int Set @Limitation = DateDiff(Day, @BEGIN_DATE, @END_DATE);
                            If (@Limitation &lt; 0) Begin Set @Limitation = 0 End;
                            Open CUR_EMPLOYEE;
                            Fetch Next From CUR_EMPLOYEE Into @REC_EMPLOYEE_ID, @REC_EMPLOYEE_CNAME;
                            While @@FETCH_STATUS = 0
                            Begin
                            While (@Index &lt;= @Limitation)
                            Begin
                            Insert Into #EMP_SIGN_RECORDS
                            Select @REC_EMPLOYEE_ID, @REC_EMPLOYEE_CNAME, DateAdd(d, @Index, @BEGIN_DATE);
                            Set @Index = @Index + 1;
                            End;
                            Set @Index = 0;
                            Fetch Next From CUR_EMPLOYEE Into @REC_EMPLOYEE_ID, @REC_EMPLOYEE_CNAME;
                            End;
                            Close CUR_EMPLOYEE;
                            DEALLOCATE CUR_EMPLOYEE;
                            Select E.EMPLOYEE_ID, E.EMP_CHINESE_NAME, IsNull(T.SIGNDATE_d, E.SIGN_DATE) As SIGNDATE_d, T.In_Time_nvc, T.Out_Time_nvc, T.Out_REASON, 
                            P.HourLimit, P.OverTime, NoonHour = Case When (N.NoonIn Is Null) Or (N.NoonOut Is Null) Then Null Else 1 End 
                            From #EMP_SIGN_RECORDS As E Left Join T_SIGN_RECORD As T
                            On E.EMPLOYEE_ID = T.LOGONID_nvc And
                            (T.SIGNDATE_d &gt;= E.SIGN_DATE And T.SIGNDATE_d &lt; DateAdd(d, 1, E.SIGN_DATE))
	                        Left Join P_0103 As P On P.employee_id = E.EMPLOYEE_ID
	                        Left Join P_0104 As N On N.employee_id = E.EMPLOYEE_ID
	                        And (N.Noon &gt;= E.SIGN_DATE And N.Noon &lt; DateAdd(d, 1, E.SIGN_DATE))
                            Where (1 = 1) Order by E.EMP_CHINESE_NAME;
                            Drop Table #EMP_SIGN_RECORDS;">

            <SelectParameters>
                <asp:Parameter Name="AD_TITLE" DbType="String" Size="50" />
                <asp:Parameter Name="employee_id" DbType="String" Size="10" />
                <asp:Parameter Name="BEGIN_DATE" DbType="String" Size="10" />
                <asp:Parameter Name="END_DATE" DbType="String" Size="10" />
            </SelectParameters>

        </asp:SqlDataSource>
    </div>
    </form>
</body>
</html>