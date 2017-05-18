<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA01002.aspx.vb" Inherits="Source_01_MOA01002" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>差假查詢</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" language="javascript">
        document.oncontextmenu = new Function("return false");
    </script>
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
<body>
    <form id="form1" runat="server">   
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="差假查詢" Width="100%"></asp:Label>
            </td></tr>
        </table>     
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Lab1" runat="server" Text="部門名稱：" CssClass="form"></asp:Label>                    
                </td>
			    <td noWrap style="height: 25px; width: 20%;">
                    <asp:DropDownList id="OrgSel"
                        DataSourceID="SqlDataSource1"
                        DataValueField="ORG_UID"
                        DataTextField="ORG_NAME"
                        runat="server" AutoPostBack="True">
                    </asp:DropDownList>&nbsp;
                </td>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Label1" runat="server" Text="姓名：" CssClass="form"></asp:Label>
                </td>
			    <td noWrap style="height: 25px; width: 15%;">
                    <asp:DropDownList ID="UserSel" runat="server" DataSourceID="SqlDataSource3" DataTextField="emp_chinese_name" DataValueField="employee_id">
                    </asp:DropDownList>
                </td>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Lab2" runat="server" Text="申請日期：" CssClass="form" ></asp:Label>
                </td>
			    <td noWrap style="height: 25px; width: 30%;">
                    <asp:TextBox ID="Sdate" runat="server" Width="80px" OnKeyDown="return false" ></asp:TextBox>
                    <div style="display: none;">
                        <img id="calImg" src="../../Image/calendar.gif" alt="選擇日期" />
                    </div>
                    <asp:Label ID="Lab3" runat="server" Text="~" ></asp:Label>
                    <asp:TextBox ID="Edate" runat="server" Width="80px" OnKeyDown="return false"></asp:TextBox>
                </td>
			    <td noWrap style="height: 25px; width: 20%;" align="center">
                    <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />
                    <asp:ImageButton ID="ImgRevoke" runat="server" ImageUrl="~/Image/revoke.gif" OnClientClick="return confirm('表單確定撤銷嗎?')"
                        ToolTip="撤銷" /></td>
		    </tr>
	    </table>
        <asp:GridView ID="GridView1" runat="server" EmptyDataText="查無任何資料" Width="100%" 
            AutoGenerateColumns="False" DataSourceID="SqlDataSource2" AllowPaging="True" 
            AllowSorting="True" CellPadding="4" ForeColor="#333333" GridLines="None" 
            EnableModelValidation="True">
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:CheckBox ID="selchk" runat="server" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="P_NUM" HeaderText="流水號" InsertVisible="False" SortExpression="P_NUM" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="PAUNIT" HeaderText="部門名稱" SortExpression="PAUNIT" >
                </asp:BoundField>
                <asp:BoundField DataField="nTYPE" HeaderText="假別" SortExpression="nTYPE" >
                </asp:BoundField>
                <asp:BoundField DataField="PANAME" HeaderText="申請人" SortExpression="PANAME" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nAPPTIME" HeaderText="申請日期" SortExpression="nAPPTIME" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nSTARTTIME" HeaderText="起日" SortExpression="nSTARTTIME" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nSTHOUR" HeaderText="時" SortExpression="nSTHOUR" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nENDTIME" HeaderText="迄日" SortExpression="nENDTIME" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nETHOUR" HeaderText="時" SortExpression="nETHOUR" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="狀態" SortExpression="PENDFLAG">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("PENDFLAG") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# FunStatus("PENDFLAG") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:CommandField ButtonType="Image" SelectImageUrl="~/Image/List.gif" ShowSelectButton="True" SelectText="詳細資料" />
                <asp:BoundField DataField="eformsn" HeaderText="eformsn" />
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
            SelectCommand="SELECT EFORMSN, PWUNIT, PWTITLE, PWNAME, PWIDNO, PAUNIT, PANAME, PATITLE, PAIDNO, nSTATUS, nAPPTIME, nTYPE, nPROVEMENT, CONVERT (char(12), nSTARTTIME, 111) AS nSTARTTIME, nSTHOUR, CONVERT (char(12), nENDTIME, 111) AS nENDTIME, nETHOUR, nDAY, nHOUR, nAGENT1, nAGENT2, nAGENT3, nPLACE, nPHONE, nREASON, P_NUM,PENDFLAG FROM P_01 WHERE (1 = 2) ORDER BY nAPPTIME DESC">
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT * FROM [ADMINGROUP] WHERE 1=2"></asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [employee_id], [emp_chinese_name], [ORG_UID] FROM [EMPLOYEE] WHERE 1=2 ORDER BY [emp_chinese_name]">
            <SelectParameters>
                <asp:ControlParameter ControlID="OrgSel" Name="ORG_UID" PropertyName="SelectedValue"
                    Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>                
                    
    </form>
</body>
</html>
