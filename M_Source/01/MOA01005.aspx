<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA01005.aspx.vb" Inherits="Source_01_MOA01005" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>年度休假累計查詢</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" language="javascript">
        document.oncontextmenu = new Function("return false");
    </script>
</head>
<body language="javascript" onload="return window_onload()">
    <form id="form1" runat="server">   
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="年度休假累計查詢" Width="100%"></asp:Label>
            </td></tr>
        </table>     
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Lab1" runat="server" Text="部門名稱：" CssClass="form"></asp:Label>                    
                </td>
			    <td noWrap style="height: 25px; width: 30%;">
                    <asp:DropDownList ID="OrgSel" runat="server" DataSourceID="SqlDataSource1" DataTextField="ORG_NAME" DataValueField="ORG_UID" AutoPostBack="True">
                    </asp:DropDownList></td>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Label1" runat="server" Text="姓名：" CssClass="form"></asp:Label>
                </td>
			    <td noWrap style="height: 25px; width: 10%;">
                    <asp:DropDownList ID="UserSel" runat="server" DataSourceID="SqlDataSource3" DataTextField="emp_chinese_name"
                        DataValueField="employee_id">
                    </asp:DropDownList></td>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Label2" runat="server" Text="年度：" CssClass="form"></asp:Label>                    
                </td>
			    <td noWrap style="height: 25px; width: 10%;">
                    <asp:DropDownList ID="YearSel" runat="server">
                    </asp:DropDownList>
                </td>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Label13" runat="server" Text="月：" CssClass="form"></asp:Label>                    
                </td>
			    <td noWrap style="height: 25px; width: 10%;">
                    <asp:DropDownList ID="MonthSel" runat="server">
                        <asp:ListItem>01</asp:ListItem>
                        <asp:ListItem>02</asp:ListItem>
                        <asp:ListItem>03</asp:ListItem>
                        <asp:ListItem>04</asp:ListItem>
                        <asp:ListItem>05</asp:ListItem>
                        <asp:ListItem>06</asp:ListItem>
                        <asp:ListItem>07</asp:ListItem>
                        <asp:ListItem>08</asp:ListItem>
                        <asp:ListItem>09</asp:ListItem>
                        <asp:ListItem>10</asp:ListItem>
                        <asp:ListItem>11</asp:ListItem>
                        <asp:ListItem>12</asp:ListItem>
                    </asp:DropDownList></td>
			    <td noWrap style="height: 25px; width: 20%;" align="center">
                    <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />
                    <asp:ImageButton ID="Img_Export" runat="server" ImageUrl="~/Image/ExportFile.gif" />
                    <asp:ImageButton ID="ImagePrint"
                        runat="server" ImageUrl="~/Image/print.gif" ToolTip="列印" /></td>
		    </tr>
	    </table>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
            AutoGenerateColumns="False" CellPadding="4" DataSourceID="SqlDataSource2" ForeColor="#333333"
            GridLines="None" Width="100%">
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <Columns>
                <asp:BoundField DataField="emp_chinese_name" HeaderText="人員名稱" SortExpression="emp_chinese_name" >
                    <HeaderStyle HorizontalAlign="Center" Width="12%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="慰勞假" SortExpression="s2">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("s2") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# FunHours("s2") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="值日補休" SortExpression="s4">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox4" runat="server" Text='<%# Bind("s4") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label4" runat="server" Text='<%# FunHours("s4") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="事假" SortExpression="s1">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox2" runat="server" Text='<%# Bind("s1") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# FunHours("s1") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="病假" SortExpression="s3">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox3" runat="server" Text='<%# Bind("s3") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label3" runat="server" Text='<%# FunHours("s3") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="公假(差)" SortExpression="s9">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox9" runat="server" Text='<%# Bind("s9") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label9" runat="server" Text='<%# FunHours("s9") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="婚假" SortExpression="s7">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox7" runat="server" Text='<%# Bind("s7") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label7" runat="server" Text='<%# FunHours("s7") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="喪假" SortExpression="s8">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox8" runat="server" Text='<%# Bind("s8") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label8" runat="server" Text='<%# FunHours("s8") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="陪產假" SortExpression="s6">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox6" runat="server" Text='<%# Bind("s6") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label6" runat="server" Text='<%# FunHours("s6") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="產假" SortExpression="s5">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox5" runat="server" Text='<%# Bind("s5") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label5" runat="server" Text='<%# FunHours("s5") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="產前假" SortExpression="s10">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox10" runat="server" Text='<%# Bind("s10") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label10" runat="server" Text='<%# FunHours("s10") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="流產假" SortExpression="s11">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox11" runat="server" Text='<%# Bind("s11") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label11" runat="server" Text='<%# FunHours("s11") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
            </Columns>
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT * FROM [ADMINGROUP] WHERE 1=2"></asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT * FROM [V_P01_ABSENCE] WHERE 1=2"></asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [employee_id], [emp_chinese_name], [ORG_UID] FROM [EMPLOYEE] WHERE 1=2 ORDER BY [emp_chinese_name]">
            <SelectParameters>
                <asp:ControlParameter ControlID="OrgSel" Name="ORG_UID" PropertyName="SelectedValue"
                    Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
                    
    </form>
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
</body>
</html>
