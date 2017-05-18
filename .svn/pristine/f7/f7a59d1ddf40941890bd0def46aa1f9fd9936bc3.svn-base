<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00010.aspx.vb" Inherits="Source_00_MOA00010" %>

<%@ Register assembly="Microsoft.ReportViewer.WebForms, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a" namespace="Microsoft.Reporting.WebForms" tagprefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
	<title>批核作業</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />    
</head>
<body >
    <form id="form1" runat="server">     
            <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
                <tr><td align="center">
                        <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="表單批核" Width="100%"></asp:Label>
                </td></tr>
            </table>
		    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
			    <tr>
				    <td noWrap width="10%" class="form">
				        <asp:CheckBox ID="AllChk" runat="server" Text="全選" AutoPostBack="True" 
                            Enabled="False" /></td>
				    <td noWrap width="10%" class="form">
                        <asp:Label ID="Label1" runat="server" Text="表單："></asp:Label></td>
				    <td noWrap width="50%" class="form">
                        <asp:DropDownList ID="DropDownList1" runat="server" DataSourceID="SqlDataSource1" DataTextField="frm_chinese_name" DataValueField="eformid" AutoPostBack="True">
                        </asp:DropDownList></td>
                    <td noWrap width="50%" class="form">
                        <asp:Label ID="lbMsg" runat="server" ForeColor = "Red" />
                    </td>
				    <td noWrap width="15%" class="form">
				        <asp:Button ID="AppBtn" runat="server" Text="核准" ToolTip="選取待批核表單" /></td>
				    <td noWrap width="15%" class="form">
				        <asp:Button ID="BackBtn" runat="server" Text="駁回" OnClientClick="return confirm('確定駁回嗎?')" /></td>
			    </tr>
		    </table>
            <asp:GridView ID="GridView1" runat="server" Width="100%" EmptyDataText="查無任何資料" AutoGenerateColumns="False" AllowPaging="True" DataSourceID="SqlDataSource2" CellPadding="4" ForeColor="#333333" GridLines="None" AllowSorting="True">
                <Columns>
                    <asp:TemplateField HeaderText="選取">
                        <EditItemTemplate>
                            <asp:CheckBox ID="CheckBox1" runat="server" />                            
                        </EditItemTemplate>
                        <ItemTemplate>
                            <asp:CheckBox ID="selchk" runat="server" />
                            <asp:HiddenField ID="HiddenField1" runat="server"  Value ='<%# Bind("eformid") %>'/>
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Center" />
                        <HeaderStyle HorizontalAlign="Center" Width="8%" />
                    </asp:TemplateField>
                    <asp:BoundField DataField="appdate" HeaderText="申請時間" SortExpression="appdate" >
                        <ItemStyle HorizontalAlign="Center" />
                        <HeaderStyle HorizontalAlign="Center" Width="25%" />
                    </asp:BoundField>
                    <asp:BoundField HeaderText="填表人" DataField="emp_chinese_name" SortExpression="emp_chinese_name" >
                        <ItemStyle HorizontalAlign="Center" />
                        <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    </asp:BoundField>
                    <asp:BoundField DataField="ShowContent" HeaderText="表單內容" SortExpression="ShowContent" >
                        <HeaderStyle HorizontalAlign="Center" Width="32%" />
                    </asp:BoundField>
                    <asp:BoundField HeaderText="表單名稱" DataField="frm_chinese_name" SortExpression="frm_chinese_name" >
                        <HeaderStyle HorizontalAlign="Center" Width="15%" />
                    </asp:BoundField>
                    <asp:CommandField ButtonType="Image" SelectImageUrl="~/Image/List.gif" SelectText="詳細資料"
                        ShowSelectButton="True" >
                        <ItemStyle HorizontalAlign="Center" />
                        <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    </asp:CommandField>
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
                SelectCommand="SELECT flowctl.eformid, flowctl.eformrole, flowctl.appdate, flowctl.empuid,flowctl.emp_chinese_name, EFORMS.frm_chinese_name, flowctl.eformsn, flowctl.hddate, V_EformShow.ShowContent FROM flowctl INNER JOIN EFORMS ON flowctl.eformid = EFORMS.eformid INNER JOIN V_EformShow ON flowctl.eformsn = V_EformShow.EFORMSN WHERE (flowctl.hddate IS NULL) and flowctl.gonogo != '1' and flowctl.empuid = @empuid ORDER BY appdate desc">
                <SelectParameters>
                    <asp:SessionParameter Name="empuid" SessionField="user_id" />
                </SelectParameters>
            </asp:SqlDataSource>
    </form>
</body>
</html>
