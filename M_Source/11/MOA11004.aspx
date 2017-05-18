<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA11004.aspx.vb" Inherits="M_Source._11.M_Source_11_MOA11004" %>

<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="cc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">

        .toptitle
        {
            font: bold 15px "Verdana, Arial, Helvetica, sans-serif";
            background-color: #6499CD;
            border: 1px solid #6499CD;
            text-align: center;
            color: #FFFFFF;
            padding-top: 2px;
        }
                
        .style1
        {
            font: 13px Verdana, Arial, Helvetica, sans-serif;
            color: #666666;
            text-decoration: none;
            height: 24px;
            width: 7%;
        }
        .form
        {
            font: 13px Verdana, Arial, Helvetica, sans-serif;
            color: #666666;
            text-decoration: none;
        }
        
        .style4
        {
            font: 13px Verdana, Arial, Helvetica, sans-serif;
            color: #666666;
            text-decoration: none;
            height: 24px;
            width: 26%;
        }
        .modalBox {
	background-color : #f5f5f5;
	border-width: 3px;
	border-style: solid;
	border-color: Blue;
	padding: 3px;
}
        .style2
        {
            font: 13px Verdana, Arial, Helvetica, sans-serif;
            color: #666666;
            text-decoration: none;
            height: 24px;
            width: 24%;
        }
        .style5
        {
            font: 13px Verdana, Arial, Helvetica, sans-serif;
            color: #666666;
            text-decoration: none;
            height: 24px;
            width: 5%;
        }
        </style>
</head>
<body>
    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
        <tr>
            <td align="center">
                <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="表單查詢" Width="100%"></asp:Label>
            </td>
        </tr>
    </table>
    <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
        <tr>
            <td nowrap class="style1">
                <asp:Label ID="Lab1" runat="server" Text="部門名稱：" CssClass="form"></asp:Label>
            </td>
            <td nowrap class="style4">
                <asp:Label ID="lblOrgSel" runat="server" Text="請選擇部門" Visible="False"></asp:Label>
                <asp:Label ID="lblOrgSelID" runat="server" Text="Label" Visible="False"></asp:Label>
                <asp:Button ID="btnSelOrg" runat="server" Text="選擇部門..." Visible="False" />
                <cc1:ModalPopupExtender ID="btnSelOrg_ModalPopupExtender" runat="server" DynamicServicePath=""
                    Enabled="True" TargetControlID="btnSelOrg" BackgroundCssClass="modalBackground"
                    PopupControlID="PanelOrgSel">
                </cc1:ModalPopupExtender>
                <asp:Panel ID="PanelOrgSel" runat="server" CssClass="modalBox" Height="400px" ScrollBars="Auto"
                    Style="display: none;" Width="400px">
                    選擇部門：
                    <asp:TreeView ID="tvOrg" runat="server" ExpandDepth="3" Target="_self">
                    </asp:TreeView>
                    <asp:Button ID="btnORGSelectOK" runat="server" OnClick="btnORGSelectOK_Click" Text="確定" />
                    <asp:Button ID="btnORGSelectCancel" runat="server" OnClick="btnORGSelectCancel_Click"
                        Text="取消" />
                </asp:Panel>
                    <asp:Label ID="Lab12" runat="server" Text="部門名稱：" CssClass="form"></asp:Label>                    
                    <asp:DropDownList id="ddlOrgSel"
                        DataSourceID="SqlDataSource1"
                        DataValueField="ORG_UID"
                        DataTextField="ORG_NAME"
                        runat="server" AutoPostBack="True">
                    </asp:DropDownList>
                    <asp:Label ID="Label1" runat="server" Text="姓名：" CssClass="form"></asp:Label>
                    <asp:DropDownList ID="ddlUserSel" runat="server" 
                    DataSourceID="SqlDataSource3" DataTextField="emp_chinese_name" 
                    DataValueField="employee_id">
                    </asp:DropDownList>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT * FROM [ADMINGROUP] WHERE 1=2"></asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            
                    SelectCommand="SELECT [employee_id], [emp_chinese_name], [ORG_UID] FROM [EMPLOYEE] WHERE ([ORG_UID] = @ORG_UID) ORDER BY [emp_chinese_name]">
            <SelectParameters>
                <asp:ControlParameter ControlID="ddlOrgSel" Name="ORG_UID" PropertyName="SelectedValue"
                    Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
        
            </td>
            <td nowrap class="style1">
                <asp:Label ID="Lab4" runat="server" Text="申請時間：" CssClass="form"></asp:Label>
                <div style="display: none;">
                    <img id="Img1" src="../../Image/calendar.gif" alt="選擇日期" />
                </div>
            </td>
            <td nowrap class="style2">
                <asp:TextBox ID="txtAPPSDate" runat="server" OnKeyDown="return false" Width="100px"></asp:TextBox>
                <div style="display: none;">
                    <img id="calImg" src="../../Image/calendar.gif" alt="選擇日期" />
                </div>
                <asp:Label ID="Lab3" runat="server" Text="~"></asp:Label>
                <asp:TextBox ID="txtAPPEDate" runat="server" OnKeyDown="return false" Width="100px"></asp:TextBox>
            </td>
            <td nowrap class="style5">
                <asp:Label ID="Lab11" runat="server" Text="種類：" CssClass="form"></asp:Label>
            </td>
            <td nowrap class="form" align="left" style="width: 20%; height: 24px">
                <asp:DropDownList ID="ddlRepairMainKind" DataSourceID="SqlDataSource4" DataValueField="Kind_Num"
                    DataTextField="Kind_Name" runat="server" AutoPostBack="True">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSource4" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                    SelectCommand="SELECT DISTINCT KIND_NUM,KIND_NAME FROM [SYSKIND] WHERE ([Kind_Num] IN ('7','8','9'))">
                </asp:SqlDataSource>
                <asp:DropDownList ID="ddlProblemKind" runat="server" DataSourceID="SqlDataSource5"
                    DataTextField="State_Name" DataValueField="Kind_SysID">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSource5" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                    SelectCommand="SELECT * FROM [SYSKIND] WHERE ([Kind_Num] = @Kind_Num)">
                    <SelectParameters>
                        <asp:ControlParameter ControlID="ddlRepairMainKind" Name="Kind_Num" PropertyName="SelectedValue"
                            Type="Int32" />
                    </SelectParameters>
                </asp:SqlDataSource>
            </td>
        </tr>
        <tr>
            <td nowrap class="style1">
                <asp:Label ID="Label3" runat="server" Text="狀態：" CssClass="form"></asp:Label>
            </td>
            <td nowrap class="style4">
                <asp:DropDownList ID="ddlStatusSel" runat="server">
                    <asp:ListItem Value="99">全部</asp:ListItem>
                    <asp:ListItem Value="0">未處理</asp:ListItem>
                    <asp:ListItem Value="1">處理中</asp:ListItem>
                    <asp:ListItem Value="2">處理完成</asp:ListItem>
                </asp:DropDownList>
            </td>
            <td nowrap class="style1">
                <asp:Label ID="Lab5" runat="server" Text="完修時間：" CssClass="form"></asp:Label>
            </td>
            <td nowrap class="style2">
                <asp:TextBox ID="txtFixedSDate" runat="server" OnKeyDown="return false" Width="100px"
                    Style="margin-top: 0px"></asp:TextBox>
                <asp:Label ID="Lab6" runat="server" Text="~"></asp:Label>
                <asp:TextBox ID="txtFixedEDate" runat="server" OnKeyDown="return false" Width="100px"></asp:TextBox>
            </td>
            <td nowrap class="style5">
                &nbsp;
            </td>
            <td nowrap class="form" align="center" style="width: 20%; height: 24px">
                <asp:ImageButton ID="Searchbtn" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />
            </td>
        </tr>
        <tr>
            <td nowrap class="style1">
                <asp:Label ID="Lab7" runat="server" Text="叫修時間：" CssClass="form"></asp:Label>
            </td>
            <td nowrap class="style4">
                <asp:TextBox ID="txtCallSDate" runat="server" OnKeyDown="return false" Width="100px"></asp:TextBox>
                <asp:Label ID="Lab8" runat="server" Text="~"></asp:Label>
                <asp:TextBox ID="txtCallEDate" runat="server" OnKeyDown="return false" Width="100px"></asp:TextBox>
            </td>
            <td nowrap class="style1">
                <asp:Label ID="Lab9" runat="server" Text="到修時間：" CssClass="form"></asp:Label>
            </td>
            <td nowrap class="style2">
                <asp:TextBox ID="txtARRSDate" runat="server" OnKeyDown="return false" Width="100px"></asp:TextBox>
                <asp:Label ID="Lab10" runat="server" Text="~"></asp:Label>
                <asp:TextBox ID="txtARREDate" runat="server" OnKeyDown="return false" Width="100px"></asp:TextBox>
            </td>
            <td nowrap class="style5">
                &nbsp;
            </td>
            <td nowrap class="form" align="center" style="width: 20%; height: 24px">
                &nbsp;
            </td>
        </tr>
    </table>
    <asp:GridView ID="gvList" runat="server" Width="100%" EmptyDataText="查無任何資料" AutoGenerateColumns="False"
        AllowPaging="True" DataSourceID="SqlDataSource2" CellPadding="4" ForeColor="#333333"
        GridLines="None" AllowSorting="True" EnableModelValidation="True">
        <Columns>
            <asp:BoundField DataField="EFORMSN" HeaderText="EFORMSN" SortExpression="EFORMSN" />
            <asp:BoundField DataField="PAUNIT" HeaderText="申請單位" SortExpression="PAUNIT"></asp:BoundField>
            <asp:BoundField HeaderText="申請人" DataField="PANAME" SortExpression="PANAME"></asp:BoundField>
            <asp:BoundField HeaderText="申請時間" DataField="APPTIME" SortExpression="APPTIME"></asp:BoundField>
            <asp:BoundField DataField="BROKENTYPE" HeaderText="報修種類" SortExpression="BROKENTYPE">
            </asp:BoundField>
            <asp:BoundField DataField="FIXSTATUS" HeaderText="維修進度" SortExpression="FIXSTATUS">
            </asp:BoundField>
            <asp:BoundField DataField="FINALDATE" HeaderText="完修時間" SortExpression="FINALDATE">
            </asp:BoundField>
            <asp:CommandField ButtonType="Button" InsertVisible="False" SelectText="詳細資料" ShowCancelButton="False"
                ShowSelectButton="True" />
            <asp:BoundField DataField="pendflag" />
        </Columns>
        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
        <PagerTemplate>
            <asp:DropDownList ID="ddlPageNo" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlPageNo_SelectedIndexChanged">
            </asp:DropDownList>
            <asp:Label ID="lblPageCount" runat="server" Text="/"></asp:Label>
            每頁<asp:DropDownList ID="ddlPageSize" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlPageSize_SelectedIndexChanged">
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
        SelectCommand="SELECT * FROM [P_11] WHERE 1=2"></asp:SqlDataSource>
    </form>
</body>
</html>
