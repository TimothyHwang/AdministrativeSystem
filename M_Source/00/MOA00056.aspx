<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00056.aspx.vb" Inherits="Source_00_MOA00056" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>個人權限控管</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" language="javascript">
        document.oncontextmenu = new Function("return false");
    </script>
</head>
<body>
    <form id="form1" runat="server">
    
    
        <table width="100%">
            <tr>
                <td align="center">
                    <asp:Label ID="Label4" runat="server" CssClass="toptitle" Text="個人權限控管" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>              
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap style="height: 25px; width: 20%;">
                    <asp:Label ID="Label5" runat="server" CssClass="form" Text="部門名稱：" Width="100px"></asp:Label></td>
			    <td noWrap style="height: 25px; width: 80%;">
                    <asp:DropDownList id="OrgSel"
                        DataSourceID="SqlDataSource1"
                        DataValueField="ORG_UID"
                        DataTextField="ORG_NAME"
                        runat="server" AutoPostBack="True">
                    </asp:DropDownList></td>
		    </tr>
		    <tr>
			    <td noWrap style="height: 25px; width: 20%;">
                    <asp:Label ID="Label6" runat="server" CssClass="form" Text="姓名：" Width="100px"></asp:Label></td>
			    <td noWrap style="height: 25px; width: 80%;">
                    <asp:DropDownList ID="UserSel" runat="server" DataSourceID="SqlDataSource3" DataTextField="emp_chinese_name" DataValueField="employee_id" AutoPostBack="True">
                    </asp:DropDownList>
                    <asp:Label ID="ErrName" runat="server" Text="請選擇姓名" Width="100px" ForeColor="Red" Visible="False"></asp:Label></td>
		    </tr>
	    </table>        
        <table width="100%">
            <tr>
                <td width="100%" style="height: 49px">
                <asp:CheckBoxList ID="ChkLimit" runat="server" DataSourceID="SqlDataSource2"
                    DataTextField="ProgName" DataValueField="LinkNum" RepeatColumns="5" RepeatDirection="Horizontal" Width="100%">
                </asp:CheckBoxList>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <asp:ImageButton ID="ImgOK" runat="server" ImageUrl="~/Image/apply.gif" />
                    <asp:Label ID="lbMsg" runat="server" ForeColor="#FF3300" ></asp:Label>
                </td>
                
            </tr>            
        </table>
        
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] ORDER BY [ORG_NAME]">
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [LinkNum], [ProgName] FROM [ROLEMAPS] ORDER BY [GroupID], [DisplayOrder]"></asp:SqlDataSource>                
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [employee_id], [emp_chinese_name], [ORG_UID] FROM [EMPLOYEE] WHERE ([ORG_UID] = @ORG_UID) ORDER BY [emp_chinese_name]">
            <SelectParameters>
                <asp:ControlParameter ControlID="OrgSel" Name="ORG_UID" PropertyName="SelectedValue"
                    Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
    
    </form>
</body>
</html>
