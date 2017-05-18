<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA01014.aspx.vb" Inherits="M_Source_01_MOA01014" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>加班查詢</title>
    <link href="../../css/jquery.datepick.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.10.2.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick-zh-TW.js" type="text/javascript"></script>

    <script type="text/javascript" language="javascript">
        document.oncontextmenu = new Function("return false");

        $(function () {
            $("#txtSdate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#txtSdate").val(), showTrigger: '#calImg' });
            $("#txtEdate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#txtEdate").val(), showTrigger: '#calImg' });
        }); 
    </script>
    <style type="text/css">



.toptitle{
	font:bold 15px "Verdana, Arial, Helvetica, sans-serif";
	background-color:#6499CD;
	border:1px solid #6499CD;
	text-align: center;
	color: #FFFFFF;
	padding-top:2px;
}
.form
{
	font: 13px Verdana, Arial, Helvetica, sans-serif;
	color: #666666;
	text-decoration: none;
}

    </style>
</head>
<body>
    <form id="form1" runat="server">   
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="個人加班查詢" Width="100%"></asp:Label>
            </td></tr>
        </table>     
	    <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
		    <tr>
                <td nowrap width="10%" class="form">
				        <asp:CheckBox ID="AllChk" runat="server" Text="全選" AutoPostBack="True" 
                            Enabled="False" /></td>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Lab1" runat="server" Text="部門名稱：" CssClass="form"></asp:Label>                    
                </td>
			    <td noWrap style="height: 25px; width: 20%;">
                    &nbsp;
                    <asp:Label ID="lblOrg" runat="server" Text="Label"></asp:Label>
                </td>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Label1" runat="server" Text="姓名：" CssClass="form"></asp:Label>
                </td>
			    <td noWrap style="height: 25px; width: 15%;">
                    <asp:Label ID="lblUser" runat="server" Text="Label"></asp:Label>
                </td>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Lab2" runat="server" Text="申請日期：" CssClass="form" ></asp:Label>
                </td>
			    <td noWrap style="height: 25px; width: 30%;">
                    <asp:TextBox ID="txtSdate" runat="server" Width="80px" 
                        OnKeyDown="return false" ></asp:TextBox>
                    <div style="display: none;">
	                    <img id="calImg" src="../../Image/calendar.gif" alt="選擇日期" />
                    </div>
                    <asp:Label ID="Lab3" runat="server" Text="~" ></asp:Label>
                    <asp:TextBox ID="txtEdate" runat="server" Width="80px" OnKeyDown="return false"></asp:TextBox>

                </td>
			    <td noWrap style="height: 25px; width: 20%;" align="center">
                    <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />
                    <asp:ImageButton ID="ImgPrint" runat="server" ImageUrl="~/Image/print.gif" 
                        ToolTip="印表" />
                    </td>
		    </tr>
	    </table>
        <asp:GridView ID="gvList" runat="server" EmptyDataText="查無任何資料" Width="100%" 
            AutoGenerateColumns="False" DataSourceID="SqlDataSource2" AllowPaging="True" 
            AllowSorting="True" CellPadding="4" ForeColor="#333333" GridLines="None" 
            EnableModelValidation="True">
            <Columns>
                <asp:TemplateField HeaderText="選取">                    
                    <ItemTemplate>
                        <asp:CheckBox ID="selchk" runat="server" />
                        <asp:HiddenField ID="hdnP_NUM" runat="server" Value='<%# Eval("P_NUM") %>' />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="PAUNIT" HeaderText="部門名稱" SortExpression="PAUNIT" >
                </asp:BoundField>
                <asp:BoundField DataField="PANAME" HeaderText="申請人" SortExpression="PANAME" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="STARTTIME" HeaderText="加班日期" 
                    SortExpression="STARTTIME" DataFormatString="{0:yyyy/MM/dd}" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="STHOUR" HeaderText="起時" SortExpression="STHOUR" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="ETHOUR" HeaderText="迄時" SortExpression="ETHOUR" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField InsertVisible="False" ShowHeader="False">
                    <EditItemTemplate>
                        <asp:Button ID="Button1" runat="server" CausesValidation="True" 
                            CommandName="Update" Text="更新" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Button ID="btnEdit" runat="server" CausesValidation="False" 
                            CommandName="Edit" Text="編輯" />
                        &nbsp;<asp:Button ID="btnDelete" runat="server" CausesValidation="False" 
                            CommandName="Delete" Text="撤銷" 
                            onclientclick="return confirm('確定刪除嗎!!!')" ToolTip="刪除" />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <PagerTemplate>
                <asp:DropDownList ID="ddlPageNo" runat="server" AutoPostBack="True" 
                    onselectedindexchanged="ddlPageNo_SelectedIndexChanged">
                </asp:DropDownList>
                <asp:Label ID="lblPageCount" runat="server" Text="/"></asp:Label>
                每頁<asp:DropDownList ID="ddlPageSize" runat="server" AutoPostBack="True" 
                    onselectedindexchanged="ddlPageSize_SelectedIndexChanged">
                    <asp:ListItem>5</asp:ListItem>
                    <asp:ListItem>10</asp:ListItem>
                    <asp:ListItem>30</asp:ListItem>
                    <asp:ListItem>50</asp:ListItem>
                    <asp:ListItem>100</asp:ListItem>
                </asp:DropDownList>
                筆
            </PagerTemplate>
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT * FROM [P_0107] WHERE 1=2" 
            InsertCommand="DELETE FROM P_0107 WHERE 1=2">
        </asp:SqlDataSource>
        
    </form>
</body>
</html>
