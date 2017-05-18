<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA02007.aspx.vb" Inherits="Source_02_MOA02007" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>會議室設備關聯</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" style="Z-INDEX: 101; LEFT: 104px; WIDTH: 508px; POSITION: absolute; TOP: 33px; HEIGHT: 201px">
            <tr>
                <td align="center">
                    <asp:Label ID="Label1" runat="server" CssClass="toptitle" Text="會議室設備" Width="100%"></asp:Label>
                </td>
            </tr>
            <tr>
                <td width="100%">               
                
                <asp:CheckBoxList ID="ChkBox1" runat="server" DataSourceID="SqlDataSource1"
                    DataTextField="DeviceName" DataValueField="DeviceSn" RepeatColumns="5" RepeatDirection="Horizontal" Width="100%">
                </asp:CheckBoxList>
                <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                    SelectCommand="SELECT [DeviceSn], [DeviceName] FROM [P_0202]"></asp:SqlDataSource>                
                </td>
            </tr>
            <tr>
                <td align="center">
                    <asp:ImageButton ID="ImgOK" runat="server" ImageUrl="~/Image/apply.gif" />
                    <asp:ImageButton ID="BackBtn" runat="server" AlternateText="回上頁" ImageUrl="~/Image/backtop.gif" /></td>
            </tr>
            
        </table>
        
    </form>
</body>
</html>
