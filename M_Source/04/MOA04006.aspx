<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04006.aspx.vb" Inherits="M_Source_04_MOA04006" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>房屋水電修繕統計表</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />    
    <link href="../../css/jquery.datepick.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.10.2.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick-zh-TW.js" type="text/javascript"></script>
    <script type="text/javascript" language="javascript">
        $(function () {
            $("#Sdate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#Sdate").val(), showTrigger: '#calImg' });
            $("#Edate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#Edate").val(), showTrigger: '#calImg' });
        }); 
    </script>
</head>
<body language="javascript" onload="return window_onload()">
    <form id="form1" runat="server">
    
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="房屋水電修繕統計表" Width="100%"></asp:Label>
            </td></tr>
        </table>   
          
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Lab1" runat="server" Text="單位：" CssClass="form"></asp:Label>                    
                </td>
			    <td noWrap style="height: 25px; width: 35%;">
                    <asp:DropDownList ID="OrgSel" runat="server" DataSourceID="SqlDataSource1" DataTextField="ORG_NAME" DataValueField="ORG_NAME" AutoPostBack="True">
                    </asp:DropDownList></td>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Label1" runat="server" Text="日期：" CssClass="form"></asp:Label></td>
                <td noWrap style="height: 25px; width: 35%;">
                    <asp:TextBox ID="Sdate" runat="server" Width="80px" OnKeyDown="return false"></asp:TextBox>
                    <div style="display: none;">
                    <img id="calImg" src="../../Image/calendar.gif" alt="選擇日期" />
                </div>
                    ~
                    <asp:TextBox ID="Edate" runat="server" Width="80px" OnKeyDown="return false"></asp:TextBox>
                    </td>			    
                
			    <td noWrap style="height: 25px; width: 20%;">
                    <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />&nbsp;<asp:ImageButton
                        ID="Img_Export" runat="server" ImageUrl="~/Image/ExportFile.gif" />
                    <asp:ImageButton ID="ImagePrint" runat="server" ImageUrl="~/Image/print.gif" ToolTip="列印" /></td>
		    </tr>		    
	    </table>
	    
        <asp:GridView ID="GridView1" runat="server" EmptyDataText="查無任何資料" AllowPaging="True" AllowSorting="True"
            AutoGenerateColumns="False" CellPadding="4" DataSourceID="SqlDataSource2" ForeColor="#333333"
            GridLines="None" Width="100%" DataKeyNames="PAUNIT">
            <Columns>
                <asp:BoundField DataField="PAUNIT" HeaderText="部門名稱" SortExpression="PAUNIT">
                    <HeaderStyle HorizontalAlign="Center" Width="60%" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="P01" HeaderText="核派" SortExpression="P01" >
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="P02" HeaderText="完成" SortExpression="P02" >
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:CommandField ButtonType="Image" SelectImageUrl="~/Image/List.gif" ShowSelectButton="True" SelectText="詳細資料" />
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
            SelectCommand="SELECT * FROM [ADMINGROUP] WHERE 1=2"></asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT * FROM T_SIGN_RECORD,EMPLOYEE WHERE T_SIGN_RECORD.LOGONID_nvc = EMPLOYEE.employee_id AND 1=2"></asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [employee_id], [emp_chinese_name], [ORG_UID] FROM [EMPLOYEE] WHERE 1=2 ORDER BY [emp_chinese_name]">
            <SelectParameters>
                <asp:ControlParameter ControlID="OrgSel" Name="ORG_UID" PropertyName="SelectedValue"
                    Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>                        
            
        <iframe id="lst" frameborder=0 width=0 height=0 src="/blank.htm"></iframe>   
        
        <script language=javascript>    
            var errmsg='<%= do_sql.G_errmsg%>';
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
