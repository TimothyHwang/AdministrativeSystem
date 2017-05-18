<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00106.aspx.vb" Inherits="Source_00_MOA00106" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>歷史資料轉移</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr>
            <td align="center">
                <asp:Label ID="Label11" runat="server" CssClass="toptitle" Text="歷史資料轉移" Width="100%"></asp:Label>
            </td>
            </tr>
        </table>
        <table border="0" width="100%" >
            <tr>
                <td width="50%" align="center">
                    <asp:Label ID="Label1" runat="server" CssClass="form" Text="轉移年度:"></asp:Label><asp:DropDownList ID="DrDown_year" runat="server">
                    </asp:DropDownList><asp:Label ID="Label9" runat="server" Text="年" CssClass="form" ></asp:Label>
                </td>
                <td width="50%">
                    <asp:Button ID="BtnSel" runat="server" Text="查詢" />
                </td>                
            </tr>
        </table>
        <table border="0" width="100%" > 
            <tr>
                <td width="50%" align="center">
                    <asp:Label ID="Label2" runat="server" Text="現有資料總筆數" CssClass="form"></asp:Label>
                    <asp:Label ID="LabCNow" runat="server" Text="" CssClass="form"></asp:Label>
                </td>
                <td width="50%">
                    <asp:Label ID="Label4" runat="server" Text="歷史資料總筆數" CssClass="form"></asp:Label>
                    <asp:Label ID="LabCOld" runat="server" Text="" CssClass="form"></asp:Label>
                </td>       
            </tr>
        </table>
        <table border="0" width="100%" >
            <tr>
                <td width="100%" align="center">
                    <asp:Label ID="Label7" runat="server" Text="歷史資料待轉移筆數" CssClass="form" ForeColor="#0000C0"></asp:Label>
                </td>
            </tr>
            <tr>
                <td width="100%" align="center">
                    <asp:Label ID="Label3" runat="server" Text="歷史資料批核完成筆數" CssClass="form" ForeColor="#0000C0"></asp:Label>
                    <asp:Label ID="LabStand" runat="server" Text="0" CssClass="form"></asp:Label>
                    <asp:Label ID="Label5" runat="server" Text="歷史資料未批核筆數" CssClass="form" ForeColor="#0000C0"></asp:Label>
                    <asp:Label ID="LabStandErr" runat="server" Text="0" CssClass="form"></asp:Label>
                </td>
            </tr>
            <tr>
                <td width="100%" align="center">
                    <asp:Table ID="TabOld" runat="server" CssClass="form" Width="100%" GridLines="Both"></asp:Table>
                </td>
            </tr>
            <tr>
                <td width="100%" align="center">
                    &nbsp;<asp:Button ID="but_exe" runat="server" Text="轉移" OnClientClick="return confirm('確定要轉移嗎?')" /></td>
            </tr> 
        </table>
    </form>
</body>
</html>
