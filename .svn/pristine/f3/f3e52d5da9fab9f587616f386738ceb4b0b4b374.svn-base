<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00211.aspx.vb" Inherits="M_Source_00_MOA00211" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>公告新增</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <link href="../../css/jquery.datepick.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.10.2.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick-zh-TW.js" type="text/javascript"></script>
    <script type="text/javascript" language="javascript">
        $(function () {            
            $("#EDate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#EDate").val(), showTrigger: '#calImg' });
        });        
    </script>
</head>
<body>
    <form id="form1" runat="server">    
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr>
                <td align="center">
                    <asp:Label ID="Label6" runat="server" CssClass="toptitle" Text="公告新增" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>        
        
        <table width="100%" border="3" bordercolor="#ccddee" align="center">
            
            <tr>
            <td style="width: 20%">
                <asp:Label ID="Label1" runat="server" Text="公告單位：" Width="100px"></asp:Label></td>
            
            <td style="width: 100%">
                <asp:DropDownList ID="DDLUnicode" runat="server" SelectedValue='<%# Bind("uicode") %>'>
                    <asp:ListItem Selected="True" Value="0">單位內部</asp:ListItem>
                    <asp:ListItem Value="1">國防部全體</asp:ListItem>
                    <asp:ListItem Value="2">系統公告</asp:ListItem>
                </asp:DropDownList></td>
            </tr>
            
            <tr>
            <td style="width: 20%">
                <asp:Label ID="Label2" runat="server" Text="標題：" Width="100px"></asp:Label></td>
            
            <td style="width: 100%">
                <asp:TextBox ID="txtTitle" runat="server" MaxLength="100" Width="230px"></asp:TextBox></td>
            </tr>
            
            <tr>
            <td style="width: 20%">
                <asp:Label ID="Label3" runat="server" Text="姓名：" Width="100px"></asp:Label></td>
            
            <td style="width: 100%">                
                <asp:Label ID="txtName" runat="server" Text="Label"></asp:Label>
                </td>
            </tr>
            
            <tr>
            <td style="width: 20%">
                <asp:Label ID="Label7" runat="server" Text="電話：" Width="100px"></asp:Label></td>
            
            <td style="width: 100%">
                <asp:TextBox ID="txtTel" runat="server" MaxLength="20" Width="130px"></asp:TextBox></td>
            </tr>
            
            <tr>
            <td style="width: 20%">
                <asp:Label ID="Label4" runat="server" Text="截止日期：" Width="100px"></asp:Label></td>
            
            <td style="width: 100%">
                <asp:TextBox ID="EDate" runat="server" onkeydown="return false" Width="110px"></asp:TextBox><div style="display: none;">
                        <img id="calImg" src="../../Image/calendar.gif" alt="選擇日期" />
                    </div></td>
            </tr>
            
            <tr>
            <td style="width: 20%">
                <asp:Label ID="Label5" runat="server" Text="內容：" Width="100px"></asp:Label></td>
            
            <td style="width: 100%">
                <asp:TextBox ID="txtContent" runat="server" MaxLength="1000" Rows="10" TextMode="MultiLine"
                    Width="500px"></asp:TextBox></td>
            </tr>
            
            <tr>
            <td colspan="2" align="center">
                <asp:ImageButton ID="ImgOK" runat="server" ImageUrl="~/Image/apply.gif" />
                <asp:ImageButton ID="BackBtn" runat="server" AlternateText="回上頁" ImageUrl="~/Image/backtop.gif" />
            </td>
            </tr>
            
        </table>                        
        
    </form>
</body>
</html>
