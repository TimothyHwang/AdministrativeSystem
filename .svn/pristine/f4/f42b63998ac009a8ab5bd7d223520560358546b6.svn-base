<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00101.aspx.vb" Inherits="Source_00_MOA00101" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>年度行事曆設定</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
<script language="javascript">
window.onload = function() {
<%
    Dim cnt As Integer = dv.Count()
    Dim i As Integer = 0
    Do While i < cnt
 %>
    document.all.D<%=CDate(dv.Table.Rows(i)(0)).ToString("yyyyMMdd")  %>.style.background="<%= butColor2 %>";
<%
    i+=1
    Loop
 %>
}
function  CalendarChage(v,d){
    MOA00101.Date.value=d;
    MOA00101.action="MOA00101.aspx";
    if(v.style.background=="<%= butColor1 %>")
        MOA00101.AddDate.value=1;
}

</script>
</head>
<body background="../../Image/main_bg.jpg">
<form id="MOA00101" runat="server" action="MOA00101.aspx" method="get">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
        <tr>
            <td style="width:3%" >
                <a href="javascript:window.history.go('-1');">
                </a><asp:ImageButton
                    ID="btnImgBack" runat="server" ImageUrl="~/Image/backtop.gif" ToolTip="回上頁" />
            </td>
            <td align='center' Width="97%" class="toptitle">
                <asp:ImageButton ID="btnImgLeft" runat="server" CssClass="toptitle" ImageUrl="~/Image/arrowleft_w.gif" ImageAlign="TextTop" ToolTip="上一年" />
                <asp:Label ID="Label1" runat="server" CssClass="toptitle" Text="年度行事曆設定" Width="40%"></asp:Label>
                <asp:ImageButton ID="btnImgRight" runat="server" CssClass="toptitle" ImageAlign="TextTop" ImageUrl="~/Image/arrowright_w.gif" ToolTip="下一年" />
            </td>
        </tr>
        </table>
        
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px; HEIGHT: 183px">
        <tr align="center" style="height:220;"><td>
        <asp:PlaceHolder ID="PlaceHolder1" runat="server" ></asp:PlaceHolder>
        </td><td>
        <asp:PlaceHolder ID="PlaceHolder2" runat="server" ></asp:PlaceHolder>
        </td><td>
        <asp:PlaceHolder ID="PlaceHolder3" runat="server" ></asp:PlaceHolder>
        </td></tr>
        <tr><td>
        <asp:PlaceHolder ID="PlaceHolder4" runat="server" ></asp:PlaceHolder>
        </td><td>
        <asp:PlaceHolder ID="PlaceHolder5" runat="server" ></asp:PlaceHolder>
        </td><td>
        <asp:PlaceHolder ID="PlaceHolder6" runat="server" ></asp:PlaceHolder>
        </td></tr>
        <tr><td>
        <asp:PlaceHolder ID="PlaceHolder7" runat="server" ></asp:PlaceHolder>
        </td><td>
        <asp:PlaceHolder ID="PlaceHolder8" runat="server" ></asp:PlaceHolder>
        </td><td>
        <asp:PlaceHolder ID="PlaceHolder9" runat="server" ></asp:PlaceHolder>
        </td></tr>
        <tr><td>
        <asp:PlaceHolder ID="PlaceHolder10" runat="server" ></asp:PlaceHolder>
        </td><td>
        <asp:PlaceHolder ID="PlaceHolder11" runat="server" ></asp:PlaceHolder>
        </td><td>
        <asp:PlaceHolder ID="PlaceHolder12" runat="server" ></asp:PlaceHolder>
        </td></tr>
        </table>
        
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand=""
            DeleteCommand=""
            InsertCommand=""
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
        </asp:SqlDataSource>
<input type="hidden" name="Date" value="" />
<input type="hidden" name="AddDate" value="0" />
    <asp:TextBox ID="txtSelYear" runat="server" Visible="False"></asp:TextBox>
</form>
</body>
</html>

