<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA02002.aspx.vb" Inherits="Source_00_MOA02002" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>會議室使用狀況</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />

    <link href="../../css/jquery.datepick.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.10.2.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick-zh-TW.js" type="text/javascript"></script>
    <script type="text/javascript" language="javascript">
        $(function () {
            $("#Sdate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#SDate").val(), showTrigger: '#calImg' });
            $("#Edate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#EDate").val(), showTrigger: '#calImg' });
        }); 
    </script>
</head>
<body>
    <form id="form1" runat="server" >   
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label1" runat="server" CssClass="toptitle" Text="會議室使用狀況" Width="100%"></asp:Label>
            </td></tr>
        </table>         
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Lab1" runat="server" Text="會議室：" CssClass="form"></asp:Label>
                </td>
			    <td noWrap style="height: 25px; width: 40%;">
                    <asp:DropDownList id="MeetSn"
                        DataSourceID="SqlDataSource1"
                        DataValueField="MeetSn"
                        DataTextField="MeetName"
                        runat="server" AutoPostBack="True">
                    </asp:DropDownList></td>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Lab2" runat="server" Text="日期：" CssClass="form" ></asp:Label>
                </td>
			    <td noWrap style="height: 25px; width: 40%;">
                    <asp:TextBox ID="Sdate" runat="server" Width="100px" OnKeyDown="return false" ></asp:TextBox>
                    <div style="display: none;">
                        <img id="calImg" src="../../Image/calendar.gif" alt="選擇日期" />
                    </div>
                    <asp:Label ID="Lab3" runat="server" Text="~" ></asp:Label>
                    <asp:TextBox ID="Edate" runat="server" Width="100px" OnKeyDown="return false"></asp:TextBox>                    
                </td>
			    <td noWrap style="height: 25px; width: 10%;">
                    <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" /></td>
		    </tr>
	    </table>                            
        
            <asp:PlaceHolder ID="PlaceHolder1" runat="server"></asp:PlaceHolder>
        &nbsp;
        <table border="0" style="z-index: 101; left: 104px; top: 33px" width="100%">
            <tr>
                <td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="會議室資訊" Width="100%"></asp:Label>
                </td>
            </tr>
            <tr>
                <td align="center">
                
                <asp:GridView ID="GridView1" runat="server" DataSourceID="SqlDataSource2" Width="100%" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" EmptyDataText="查無任何資料">
                    <Columns>
                        <asp:BoundField DataField="emp_chinese_name" HeaderText="管理者" SortExpression="emp_chinese_name">
                            <HeaderStyle HorizontalAlign="Center" Width="30%" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Tel" HeaderText="管理者電話" SortExpression="Tel">
                            <HeaderStyle HorizontalAlign="Center" Width="30%" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ContainNum" HeaderText="容納人數" SortExpression="ContainNum">
                            <HeaderStyle HorizontalAlign="Center" Width="20%" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="會議室設備">
                            <ItemTemplate>
                                <asp:BulletedList ID="BulletedList1" runat="server" BulletStyle="Numbered" DataSourceID="SqlDataSource4"
                                    DataTextField="DeviceName" DataValueField="DeviceName">
                                </asp:BulletedList>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" Width="20%" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>
                    </Columns>
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <EditRowStyle BackColor="#999999" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EmptyDataRowStyle ForeColor="DarkBlue" />
                </asp:GridView>
                
                
                </td>
            </tr>
        </table>
        
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                SelectCommand="SELECT [MeetSn], MeetName+'('+(select emp_chinese_name from EMPLOYEE where Owner = employee_id)+')' as MeetName FROM [P_0201] WHERE (share = @share OR Org_Uid=@Org_Uid) AND (Enabled=1) ORDER BY [MeetName]"
                ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
            <SelectParameters>
                <asp:Parameter DefaultValue="1" Name="share" />
                <asp:SessionParameter DefaultValue="" Name="Org_Uid" SessionField="org_uid" />
            </SelectParameters>
        </asp:SqlDataSource>        
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>" SelectCommand="SELECT P_0201.Tel, P_0201.ContainNum, P_0201.MeetSn, EMPLOYEE.emp_chinese_name FROM P_0201 INNER JOIN EMPLOYEE ON P_0201.Owner = EMPLOYEE.employee_id AND MeetSn = @MeetSn ORDER BY MeetName">
            <SelectParameters>
                <asp:ControlParameter ControlID="MeetSn" Name="MeetSn" PropertyName="SelectedValue"
                    Type="Decimal" />
            </SelectParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource4" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT P_0202.DeviceName FROM P_0202 INNER JOIN P_0203 ON P_0202.DeviceSn = P_0203.DeviceSn WHERE (P_0203.MeetSn = @PLACE) ORDER BY P_0203.MeetSn">
            <SelectParameters>
                <asp:ControlParameter ControlID="MeetSn" Name="PLACE" PropertyName="SelectedValue" />
            </SelectParameters>
        </asp:SqlDataSource>
        
            
    </form>
</body>
</html>
