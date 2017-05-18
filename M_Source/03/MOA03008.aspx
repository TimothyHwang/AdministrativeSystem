<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA03008.aspx.vb" Inherits="Source_03_MOA03008" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>車輛數量管理</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server" >     
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 103px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="車輛數量管理" Width="100%"></asp:Label>
            </td></tr>
        </table>
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
            <tr><td noWrap width="10%">
                <asp:Label ID="Label3" runat="server" Text="日期：" CssClass="form"></asp:Label></td><td nowrap="noWrap" width="80%">
                <asp:TextBox ID="PCN_Date1" runat="server" Width="120px" OnKeyDown="return false"></asp:TextBox>
                <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />
                <asp:Label ID="Label4" runat="server" Text="~" CssClass="form"></asp:Label>
                <asp:TextBox ID="PCN_Date2" runat="server" Width="120px" OnKeyDown="return false"></asp:TextBox>
                <asp:ImageButton ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" />
                <asp:Label ID="ErrMsg" runat="server" ForeColor='Red' Width="160px" ></asp:Label></td><td nowrap="noWrap" width="10%">
                <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" title="查詢" ToolTip="查詢"/></td></tr>
		</table>
	    <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" DataSourceID="SqlDataSource2" Width="100%" DataKeyNames="PCN_NUM" CellPadding="4" ForeColor="#333333" GridLines="None">
            <Columns>
                <asp:BoundField DataField="PCN_Date" HeaderText="日期" SortExpression="PCN_Date" ReadOnly="True">
                    <ItemStyle HorizontalAlign="Center" Width="30%" />
                </asp:BoundField>
                <asp:BoundField DataField="PCK_Name" HeaderText="車種名稱" SortExpression="PCK_Name" ReadOnly="True">
                    <ItemStyle HorizontalAlign="Center" Width="30%" />
                </asp:BoundField>
                <asp:BoundField DataField="PCN_Use" HeaderText="車輛可使用數量" SortExpression="PCN_Use">
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <EditItemTemplate>
                        <asp:ImageButton ID="btnImgOK" runat="server" CommandName="Update" ImageUrl="~/Image/apply.gif"
                            ToolTip="確認" />&nbsp;<asp:ImageButton ID="btnImgCancel" runat="server" CommandName="Cancel"
                                ImageUrl="~/Image/cancel.gif" ToolTip="取消" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:ImageButton ID="ImgUpd" runat="server" CommandName="Edit" ImageUrl="~/Image/update.gif"
                            ToolTip="修改" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:TemplateField>
            </Columns>
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand=""
            DeleteCommand=""
            InsertCommand=""
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="" UpdateCommand="UPDATE [P_0306] SET [PCN_Use] = @PCN_Use WHERE [PCN_NUM] = @PCN_NUM"
                    ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
            <UpdateParameters>
                <asp:Parameter Name="PCN_NUM" Type="String" />
                <asp:Parameter Name="PCN_Use" Type="Int32" />
            </UpdateParameters>
        </asp:SqlDataSource>
        
        <div id="Div_grid" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:165pt; height:150pt; left:194px; top:495px; display:block;" visible="false">
         
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
        
        <div id="Div_grid2" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:165pt; height:150pt; left:426px; top:497px; display:block;" visible="false">
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
