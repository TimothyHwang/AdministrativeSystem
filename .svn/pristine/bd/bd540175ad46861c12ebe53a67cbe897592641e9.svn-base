<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00012.aspx.vb" Inherits="Source_00_MOA00012" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>表單查詢</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />  
    <link href="../../css/jquery.datepick.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.10.2.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick-zh-TW.js" type="text/javascript"></script>
    <script type="text/javascript" language="javascript">
        $(function () {
            $("#SDate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#SDate").val(), showTrigger: '#calImg' });
            $("#EDate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#EDate").val(), showTrigger: '#calImg' });            
        }); 
    </script>
</head>
<body>
    <form id="form1" runat="server" > 
            <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
                <tr><td align="center">
                        <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="表單查詢" Width="100%"></asp:Label>
                </td></tr>
            </table>
		    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
			    <tr>
				    <td noWrap width="10%" class="form" style="height: 24px">表單：</td>
				    <td noWrap class="form" style="width: 30%; height: 24px">
                        <asp:DropDownList ID="EformSel" runat="server" DataSourceID="SqlDataSource1" DataTextField="frm_chinese_name" DataValueField="eformid">
                        </asp:DropDownList></td>
				    <td noWrap width="10%" class="form" style="height: 24px">日期：</td>
				    <td noWrap class="form" style="width: 30%; height: 24px">
                        <asp:TextBox ID="SDate" runat="server" OnKeyDown="return false" Width="100px"></asp:TextBox>
                        <div style="display: none;">
                            <img id="calImg" src="../../Image/calendar.gif" alt="選擇日期" />
                        </div>
                        <asp:Label ID="Lab3" runat="server" Text="~" ></asp:Label>
                        <asp:TextBox ID="EDate" runat="server" OnKeyDown="return false" Width="100px"></asp:TextBox>                        
                        </td>
                        <td noWrap class="form" style="width: 30%; height: 24px">
                        <asp:Label ID="Label3" runat="server" Text="狀態：" CssClass="form" ></asp:Label>
                    <asp:DropDownList ID="StateSel" runat="server">
                        <asp:ListItem Value="99">請選擇</asp:ListItem>
                        <asp:ListItem Value="E">已批核</asp:ListItem>
                        <asp:ListItem Value="?">未批核</asp:ListItem>                        
                        <asp:ListItem Value="0">駁回</asp:ListItem>                        
                    </asp:DropDownList>
                            
                        </td>
				    <td noWrap class="form" align="center" style="width: 20%; height: 24px">
                        <asp:ImageButton ID="Searchbtn" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" /></td>
			    </tr>
		    </table>
            <asp:GridView ID="GridView1" runat="server" Width="100%" EmptyDataText="查無任何資料" AutoGenerateColumns="False" AllowPaging="True" DataSourceID="SqlDataSource2" CellPadding="4" ForeColor="#333333" GridLines="None" AllowSorting="True"  DataKeyNames="flowsn">
                <Columns>
                    <asp:BoundField DataField="appdate" HeaderText="申請時間" SortExpression="appdate" ReadOnly="True" >
                        <ItemStyle HorizontalAlign="Center" />
                        <HeaderStyle HorizontalAlign="Center" Width="25%" />
                    </asp:BoundField>
                    <asp:BoundField HeaderText="填表人" DataField="emp_chinese_name" SortExpression="emp_chinese_name" ReadOnly="True" >
                        <ItemStyle HorizontalAlign="Center" />
                        <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    </asp:BoundField>
                    <asp:BoundField HeaderText="表單名稱" DataField="frm_chinese_name" SortExpression="frm_chinese_name" ReadOnly="True" >
                        <HeaderStyle HorizontalAlign="Center" Width="20%" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:TemplateField HeaderText="詳細資料" ShowHeader="False">
                        <EditItemTemplate>
                            &nbsp;
                        </EditItemTemplate>
                        <ItemTemplate>
                            &nbsp;<asp:ImageButton ID="ImageButton2" runat="server" CausesValidation="False"
                                CommandName="Select" ImageUrl="~/Image/List.gif" Text="詳細資料" />
                        </ItemTemplate>
                        <HeaderStyle HorizontalAlign="Center" Width="10%" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:TemplateField>
                    <asp:BoundField DataField="eformid" HeaderText="eformid" SortExpression="eformid" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="empuid" HeaderText="empuid" SortExpression="empuid" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="eformsn" HeaderText="eformsn" SortExpression="eformsn" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="eformrole" HeaderText="eformrole" SortExpression="eformrole" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="comment" HeaderText="批核意見" SortExpression="comment">
                        <HeaderStyle HorizontalAlign="Center" Width="25%" />
                    </asp:BoundField>
                    <asp:TemplateField HeaderText="狀態" SortExpression="gonogo">
                        <EditItemTemplate>
                            <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("gonogo") %>'></asp:TextBox>
                        </EditItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label1" runat="server" Text='<%# FunStatus("gonogo") %>'></asp:Label>
                        </ItemTemplate>
                        <HeaderStyle HorizontalAlign="Center" Width="10%" />
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
            <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                                SelectCommand="SELECT [eformid], [frm_chinese_name] FROM [EFORMS] ORDER BY [PRIMARY_TABLE]"></asp:SqlDataSource>
            <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                SelectCommand="SELECT flowctl.flowsn, flowctl.eformid, flowctl.eformrole, flowctl.appdate, flowctl.empuid, V_EformFlow.emp_chinese_name, V_EformFlow.frm_chinese_name, flowctl.eformsn, flowctl.comment,flowctl.gonogo, flowctl.hddate FROM flowctl INNER JOIN V_EformFlow ON flowctl.eformsn = V_EformFlow.eformsn WHERE (flowctl.gonogo <> '-') AND (flowctl.gonogo <> 'G') and flowctl.empuid = @empuid and 1=2 " UpdateCommand="UPDATE flowctl SET comment = @comment WHERE (flowsn = @flowsn)">
                <SelectParameters>
                    <asp:SessionParameter Name="empuid" SessionField="user_id" />
                </SelectParameters>
                <UpdateParameters>
                    <asp:Parameter Name="comment" />
                    <asp:Parameter Name="flowsn" />
                </UpdateParameters>
            </asp:SqlDataSource>
            
        <div id="Div_grid" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:165pt; height:150pt; left:196px; top:616px; display:block;" visible="false">
         
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
        
        <div id="Div_grid2" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:165pt; height:150pt; left:433px; top:617px; display:block;" visible="false">
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
