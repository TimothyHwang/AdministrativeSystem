<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04019_1.aspx.vb" Inherits="M_Source_04_MOA04019_1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>物料詳細資料</title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
    <style type="text/css">
        .style1
        {
            width: 147px;
        }
    </style>
</head>
<body background="../../Image/main_bg.jpg">
<body>
    <form id="form1" runat="server">
    <div>
          <table border="0" width="400" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center" colspan = "2">
                    <asp:Label ID="lblTitle" runat="server" CssClass="toptitle" Text="倉儲資料-物料詳細資料" Width="100%"></asp:Label>
                </td>
            </tr>
            <tr>
            <td  colspan = "2"><asp:Label ID="lb_it_code1" runat="server" Visible = "false" /></td>
            </tr>
            <asp:Repeater ID="RptitemList" runat="server" DataSourceID="sdsDataRecords" ><ItemTemplate>
            <tr>
                <td width = "25%" >物料圖片1：</td>
                <td><asp:Image ID="Image1" runat="server" ImageUrl ='<%# sPicPath + Eval("it_code")+ "A_" + Eval("file_a") %>' Visible='<%# showPic(Eval("file_a"))%>' />
                <asp:Label ID="lbImage1" runat="server" Text ="此物料無上傳圖片一" Visible='<%# showPicLB(Eval("file_a"))%>' /></td>
            </tr>
            <tr>
                <td>物料圖片2：</td>
                <td><asp:Image ID="Image2" runat="server" ImageUrl ='<%# sPicPath + Eval("it_code")+ "B_" + Eval("file_b") %>' Visible='<%# showPic(Eval("file_b"))%>' />
                <asp:Label ID="lbImage2" runat="server" Text ="此物料無上傳圖片二" Visible='<%# showPicLB(Eval("file_a"))%>'/></td>
            </tr>
             <tr>
                <td>物料代號：</td>
                <td><asp:Label ID="lb_it_code" runat="server" Text ='<%# Bind("it_code") %>' /></td>
            </tr>
            <tr>
                <td>物料名稱：</td>
                <td><asp:Label ID="lb_it_name" runat="server" Text ='<%# Bind("it_name") %>' /></td>
            </tr>
             <tr>
                <td>物料規格：</td>
                <td><asp:Label ID="lb_it_spec" runat="server"  Text ='<%# Bind("it_spec") %>' /></td>
            </tr>
             <tr>
                <td>物料價格：</td>
                <td><asp:Label ID="lb_it_cost" runat="server" Text ='<%# Bind("it_cost") %>' /></td>
            </tr>
             <tr>
                <td>物料公司：</td>
                <td><asp:Label ID="lb_manufacturer" runat="server" Text ='<%# Bind("manufacturer") %>'  /></td>
            </tr>
             <tr>
                <td>安全數量：</td>
                <td><asp:Label ID="lb_snum" runat="server" Text ='<%# Bind("snum") %>' /></td>
            </tr>
             <tr>
                <td>有效期：</td>
                <td><asp:Label ID="lb_expired_y" runat="server" Text ='<%# Bind("expired_y") %>' /></td>
            </tr>
             <tr>
                <td>物料類別：</td>
                <td><asp:Label ID="lb_it_sort" runat="server" Text ='<%# Bind("it_sort") %>' /></td>
            </tr>
             <tr>
                <td>物料單位：</td>
                <td><asp:Label ID="lb_it_unit" runat="server" Text ='<%# Bind("it_unit") %>' /></td>
            </tr>
             </ItemTemplate></asp:Repeater>
        </table>
        <br />
        <center><a herf ="#" onclick = "window.close();"><asp:Button ID="btClose" runat="server" Text="關閉視窗" /></a></center>
    </div>
        <asp:SqlDataSource ID="sdsDataRecords" runat="server"
            ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="select it_code,it_name,it_cost,manufacturer,
            snum,expired_y,it_sort,it_spec,it_unit,
            isnull(file_a,'') as file_a,isnull(file_b,'') as file_b
            from P_0407 with(nolock) where it_code=@it_code;">
        <SelectParameters>
            <asp:ControlParameter ControlID="lb_it_code1" Name="it_code" PropertyName="Text" type="String" />
        </SelectParameters>
        </asp:SqlDataSource>
    </form>
</body>
</html>
