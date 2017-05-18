<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA01012.aspx.vb" Inherits="M_Source_01_MOA01012" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>年度休假天數管理</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.7.2.min.js" type="text/javascript"></script>
    <script type="text/javascript">
        function checkText(obj) {
            var re = /^[0-9]+$/;
            if (!re.test($(obj).val()) && $(obj).val() != "") {
                alert("請輸入正整數");
                $(obj).val('');
            }
        }

        function checkData() {
            var checkResult = true;
            $("input[type=text]").each(function () {
                if ($(this).val() == "") {
                    if (!confirm("您有未輸入慰勞假天數之欄位,確定仍要儲存資料嗎?")) {
                        checkResult = false; //使用者取消儲存資料
                    }
                    return false; //跳離each迴圈
                }
            });
            return checkResult;
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">   
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="年度休假天數管理" Width="100%"></asp:Label>
            </td></tr>
        </table>     
	    <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td nowrap="nowrap" align="center" style="height: 25px; width: 35%;">
                    <asp:Label ID="Lab1" runat="server" Text="年度：" CssClass="form"></asp:Label>                    
                    <asp:DropDownList ID="ddlYear" runat="server" 
                        DataTextField="Year" DataValueField="Year" AutoPostBack="True">
                    </asp:DropDownList>
                </td>
                <td nowrap="nowrap" align="center" style="height: 25px; width: 35%;">
                
                    <asp:Label ID="Lab2" runat="server" Text="部門：" CssClass="form"></asp:Label>                    
                <asp:DropDownList ID="ddlOrgSel" DataSourceID="SqlDataSource1" DataValueField="ORG_UID"
                    DataTextField="ORG_NAME" runat="server" AutoPostBack="True">
                </asp:DropDownList>
    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
        SelectCommand="SELECT * FROM [ADMINGROUP] WHERE 1=2"></asp:SqlDataSource>
                
                </td>
                <td nowrap="nowrap" align="center" style="height: 25px; width: 30%;">
                    <asp:ImageButton ID="ImgSubmit" runat="server" ImageUrl="~/Image/apply.gif" 
                        ToolTip="確認" onclientclick="return checkData();" />
                </td>
		    </tr>
	    </table>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
            AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333"
            GridLines="None" Width="100%" EnableModelValidation="True">
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <Columns>
                <asp:BoundField DataField="emp_chinese_name" HeaderText="人員名稱" >
                <HeaderStyle HorizontalAlign="Center" />
                <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="慰勞假天數">
                    <ItemTemplate>
                        <asp:TextBox ID="txtHolidays" Text="<%# Bind('Holidays') %>" runat="server" MaxLength="3" Width="25px" onblur="javascript:checkText(this);"></asp:TextBox>
                        <asp:Label ID="lbHolidays" Text="<%# Bind('Holidays') %>" runat="server"></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
            </Columns>
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
        </asp:GridView>
                    
    </form>
</body>
</html>
