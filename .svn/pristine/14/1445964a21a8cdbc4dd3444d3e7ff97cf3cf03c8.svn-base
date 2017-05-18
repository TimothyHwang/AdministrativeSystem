<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00043.aspx.vb" Inherits="M_Source_00_MOA00043" %>
<%@ Assembly Name="system.DirectoryServices,Version=2.0.0.0,Culture=neutral,PublickeyToken=B03f5f7f11d50a3a"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>人員新增</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">    
        <table border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;" align="center">
            <tr>
            <td align="center" style="height: 24px">
                    <asp:Label ID="LabCaption" runat="server" CssClass="toptitle" Text="人員新增" Width="100%"></asp:Label>
            </td>
            </tr>
        </table>
        
        <table width="100%" border="3" bordercolor="#ccddee" align="center">
            <tr>
            <td style="width: 20%">
                <asp:Label ID="Label1" runat="server" Text="人員帳號："></asp:Label></td>
            <td style="width: 80%">
                <asp:TextBox ID="txtUserID" runat="server" Width="100px"></asp:TextBox>
                <asp:ImageButton ID="btnImgRead" runat="server" ImageUrl="~/Image/ReadAD.gif" />
                <asp:ImageButton ID="ImgClearAll" runat="server" ImageUrl="~/Image/clearall.gif" ToolTip="清除" />
                <asp:Label ID="LabErr" runat="server" ForeColor="Red" Width="270px"></asp:Label></td>
            </tr>
            
            <tr>
            <td style="width: 20%; height: 23px;">
                <asp:Label ID="Label2" runat="server" Text="單位："></asp:Label></td>
            <td style="width: 80%; height: 23px;">
                <asp:Label ID="LabUnit" runat="server" Width="450px" ></asp:Label>
            </td>
            </tr>
            
            <tr>
            <td style="width: 20%">
                <asp:Label ID="Label3" runat="server" Text="職稱："></asp:Label></td>
            <td style="width: 80%">
                <asp:Label ID="LabTitle" runat="server" Width="300px" ></asp:Label>
            </td>
            </tr>
            
            <tr>
            <td style="width: 20%">
                <asp:Label ID="Label4" runat="server" Text="姓名："></asp:Label></td>
            <td style="width: 80%">
                <asp:Label ID="LabName" runat="server" Width="150px" ></asp:Label></td>
            </tr>
            
            <tr>
            <td style="width: 20%">
                <asp:Label ID="Label5" runat="server" Text="公務郵件帳號："></asp:Label></td>
            <td style="width: 80%">
                <asp:Label ID="LabMail" runat="server" Width="200px" ></asp:Label>
            </td>
            </tr>
            
            <tr>
            <td style="width: 20%">
                <asp:Label ID="Label6" runat="server" Text="組織圖角色："></asp:Label></td>
            <td style="width: 80%">
                <asp:DropDownList ID="DDLORG" runat="server" DataSourceID="SqlDataSource4" DataTextField="ORG_NAME" DataValueField="ORG_UID">
                </asp:DropDownList></td>
            </tr>
            
            <% If Request.QueryString("empuid") <> "" Then%>
            <tr>
            <td style="width: 20%">
                <asp:Label ID="Label8" runat="server" Text="在職："></asp:Label></td>
            <td style="width: 80%">
                <asp:DropDownList ID="DDLYN" runat="server">
                    <asp:ListItem Value="Y">在職</asp:ListItem>
                    <asp:ListItem Value="N">離職</asp:ListItem>
                </asp:DropDownList></td>
            </tr>
            <%End If%>
            
            <tr>
            <td colspan="2" align="center">
                <asp:ImageButton ID="btnImgUpd" runat="server" ImageUrl="~/Image/apply.gif" OnClientClick="return confirm('確定修改人員資料嗎?')" ToolTip="確認" />
                <asp:ImageButton ID="btnImgDel" runat="server" ImageUrl="~/Image/movewait.gif" OnClientClick="return confirm('確定人員移至待派區嗎?')"
                    Visible="False" />
                <asp:ImageButton ID="btnImgBack" runat="server" ImageUrl="~/Image/backtop.gif" ToolTip="回上頁" /></td>
            </tr>
        </table>
        <asp:SqlDataSource ID="SqlDataSource4" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] ORDER BY [ORG_NAME]">
        </asp:SqlDataSource>
        
    </form>
</body>
</html>
