<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00103.aspx.vb" Inherits="Source_00_MOA00103" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>密碼變更</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />

<script language="javascript" type="text/javascript">
<!--
document.oncontextmenu = new Function("return false");
function window_onload() {
   <%if do_sql.G_errmsg <>"" then %>  
    alert('<%= do_sql.G_errmsg.Replace("'", "") %>');
  <%end if  %> 
}

// -->
</script>
</head>
<body language="javascript" onload="return window_onload()">
    <form id="form1" runat="server">
    <div>
         <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label2" runat="server" CssClass="toptitle" Text="密碼變更" Width="100%"></asp:Label>
            </td></tr>
        </table>
         <table border="0" style="width: 800px; height: 57px; color:Red" >
             <tr>
              <td style="width: 800px">
                        <asp:Label ID="Label1" runat="server" Text="單位:" Width="58px" ForeColor="Black" CssClass="form"></asp:Label>&nbsp;
                        <asp:DropDownList ID="DrDown_PAUNIT" runat="server" Width="194px" AutoPostBack="True">
                        </asp:DropDownList>&nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;&nbsp;
                 <td>
            </tr>
             <tr>
              <td style="width: 800px">
                        <asp:Label ID="Label5" runat="server" Text="姓名:" Width="58px" ForeColor="Black" CssClass="form"></asp:Label>&nbsp;
                        <asp:DropDownList ID="DrDown_emp_chinese_name" runat="server" Width="143px" AutoPostBack="True">
                        </asp:DropDownList>
              <td>
            </tr>
            <tr>
                    <td style="height: 26px; width: 800px;" >
                        <asp:Label ID="Label12" runat="server" ForeColor="Black" Text="密碼:" CssClass="form"></asp:Label>
                        &nbsp; &nbsp; &nbsp; &nbsp;<asp:TextBox ID="TXT_password" runat="server"  MaxLength="10" TextMode="Password"></asp:TextBox></td>
                </tr> 
         </table>  
         <table border="0" style="width: 100%; height: 40px">
                <tr>
                   <td style="width: 100%; height: 23px;">
                       <asp:ImageButton ID="btnImgIns" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" /></td>
                </tr> 
              </table>
    </div>
    </form>
</body>
</html>
