<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA02003.aspx.vb" Inherits="Source_02_MOA02003" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>會議室使用統計功能</title>
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
<body onload="return window_onload()">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="會議室使用統計" Width="100%"></asp:Label>
            </td></tr>
        </table>     
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap style="height: 25px; width: 10%;">
                    <asp:Label ID="Lab1" runat="server" Text="會議室：" CssClass="form"></asp:Label></td>
			    <td noWrap style="height: 25px; width: 30%;">
                    <asp:DropDownList id="MeetSn"
                        AutoPostBack="False"
                        DataSourceID="SqlDataSource1"
                        DataValueField="MeetSn"
                        DataTextField="MeetName"
                        runat="server" Width="150px">
                    </asp:DropDownList>
                </td>
			    <td noWrap style="height: 25px; width: 10%;">
                    <asp:Label ID="Label1" runat="server" Text="借用部門：" CssClass="form"></asp:Label></td>
			    <td noWrap style="height: 25px; width: 25%;">
                    <asp:TextBox ID="MeetOrg" runat="server" Width="200px"></asp:TextBox></td>
                 <td noWrap style="height: 25px; width: 25%;" align="center">
                    </td>
		    </tr>
		    <tr>
		        <td noWrap style="height: 25px; width: 10%;">			    
			        <asp:Label ID="Label3" runat="server" Text="申請日期：" CssClass="form"></asp:Label></td>
			    <td noWrap style="height: 25px; width: 30%;">
                    <asp:TextBox ID="SDate" runat="server" Width="80px" OnKeyDown="return false"></asp:TextBox>
                    <div style="display: none;">
                        <img id="calImg" src="../../Image/calendar.gif" alt="選擇日期" />
                    </div>                    
                    <asp:Label ID="Lab3" runat="server" Text="~" CssClass="form" ></asp:Label>
                    <asp:TextBox ID="EDate" runat="server" Width="80px" OnKeyDown="return false"></asp:TextBox>                    
                </td>
                <td noWrap style="height: 25px; width: 10%;">
                    <asp:Label ID="Label2" runat="server" Text="借用人：" CssClass="form"></asp:Label></td>
			    <td noWrap style="height: 25px; width: 25%;">
			        <asp:TextBox ID="BorrowPer" runat="server" Width="200px" ></asp:TextBox>
			    </td>
		        <td noWrap style="height: 25px; width: 25%;" align="center">
                    <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />&nbsp;<asp:ImageButton
                        ID="Img_Export" runat="server" ImageUrl="~/Image/ExportFile.gif" />
                    <asp:ImageButton ID="ImagePrint" runat="server" ImageUrl="~/Image/print.gif" ToolTip="列印" /></td>
		    </tr> 
	    </table>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
            EmptyDataText="查無任何資料" AutoGenerateColumns="False" DataSourceID="SqlDataSource2" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
            <Columns>
                <asp:BoundField DataField="MeetName" HeaderText="會議室名稱" SortExpression="MeetName" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="PAUNIT" HeaderText="借用部門" SortExpression="PAUNIT" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="PANAME" HeaderText="借用人" SortExpression="PANAME" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="UseCount" HeaderText="借用次數" SortExpression="UseCount" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:CommandField ButtonType="Image" SelectImageUrl="~/Image/List.gif" SelectText="詳細資料"
                    ShowSelectButton="True" />
                <asp:BoundField DataField="MeetSn" HeaderText="MeetSn" InsertVisible="False" ReadOnly="True"
                    SortExpression="MeetSn" />
                <asp:BoundField DataField="PAIDNO" SortExpression="PAIDNO" />
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
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT P_0201.MeetSn, P_0201.Org_Uid, P_0201.MeetName, P_0201.Owner, P_0201.Share, COUNT(P_0204.EFORMSN) AS UseCount, P_02.PAUNIT, P_02.PANAME FROM P_0204 INNER JOIN P_02 ON P_0204.EFORMSN = P_02.EFORMSN RIGHT OUTER JOIN P_0201 ON P_0204.MeetSn = P_0201.MeetSn WHERE (P_02.nAPPLYTIME BETWEEN GETDATE() - 30 AND GETDATE()) GROUP BY P_0201.MeetSn, P_0201.Org_Uid, P_0201.MeetName, P_0201.Owner, P_0201.Share, P_02.PAUNIT, P_02.PANAME HAVING (P_0201.Share = 1) AND P_0201.Org_Uid = @org_uid AND 1=2 ">
            <SelectParameters>
                <asp:SessionParameter Name="org_uid" SessionField="org_uid" />
            </SelectParameters>
        </asp:SqlDataSource>
        
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                SelectCommand="SELECT [MeetSn], [MeetName] FROM [P_0201] WHERE org_uid = @org_uid AND Enabled=1 ORDER BY [MeetName]"
                ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
            <SelectParameters>
                <asp:SessionParameter DefaultValue="" Name="org_uid" SessionField="org_uid" />
            </SelectParameters>
        </asp:SqlDataSource>
        &nbsp;
                    
            <iframe id="lst" frameborder=0 width=0 height=0 src="/blank.htm"></iframe>   
        
            <script language=javascript>    
            var errmsg='<%= do_sql.G_errmsg %>';
            var print_file='<%= print_file%>';

            function window_onload() {
                if (errmsg != '') {
                    alert(errmsg);
                }
                
                if (print_file != '') {
                    lst.navigate(print_file);              
                }
            }

            </script>
    </form>    
</body>
</html>
