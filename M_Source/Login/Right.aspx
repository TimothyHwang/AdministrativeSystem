<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Right.aspx.vb" Inherits="OA_Right" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>個人表單狀況</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">    
        <table border="0" width="76%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;" align="center">
            <tr>
            <td align="center" style="height: 24px">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="個人表單狀況" Width="100%"></asp:Label>
            </td>
            </tr>
        </table>
        
        <% If strApprote <> "" Then%>
        <table width="75%" border="3" bordercolor="#ccddee" align="center">
            <tr>
            <td align="center" class="CellClass" style="width: 35%" rowspan="8">
                <asp:Label ID="Label1" runat="server" Text="未批核表單" ></asp:Label>
            </td>
            <td style="width: 42%">
                <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="~/M_Source/00/MOA00010.aspx?strHyper=YAqBTxRP8P">差假作業</asp:HyperLink>
            </td>
            <td style="width: 15%" align="right">
                <asp:Label ID="LabApp1" runat="server" ></asp:Label>
            </td>
            <td style="width: 8%" align="center">
                <asp:Label ID="Label2" runat="server" Text="筆" ></asp:Label>
            </td>
            </tr>
            
            <tr>
            <td style="width: 42%">
                <asp:HyperLink ID="HyperLink2" runat="server" NavigateUrl="~/M_Source/00/MOA00010.aspx?strHyper=4rM2YFP73N">會議室作業</asp:HyperLink>
            </td>
            <td style="width: 15%" align="right">
                <asp:Label ID="LabApp2" runat="server" ></asp:Label>
            </td>
            <td style="width: 8%" align="center">
                <asp:Label ID="Label4" runat="server" Text="筆" ></asp:Label>
            </td>
            </tr>
            
            <tr>
            <td style="width: 42%">
                <asp:HyperLink ID="HyperLink3" runat="server" NavigateUrl="~/M_Source/00/MOA00010.aspx?strHyper=j2mvKYe3l9">派車作業</asp:HyperLink>
            </td>
            <td style="width: 15%" align="right">
                <asp:Label ID="LabApp3" runat="server" ></asp:Label>
            </td>
            <td style="width: 8%" align="center">
                <asp:Label ID="Label18" runat="server" Text="筆" ></asp:Label>
            </td>
            </tr>
            
            <tr>
            <td style="width: 42%">
                <asp:HyperLink ID="HyperLink4" runat="server" NavigateUrl="~/M_Source/00/MOA00010.aspx?strHyper=61TY3LELYT">房舍水電修繕作業</asp:HyperLink>
            </td>
            <td style="width: 15%" align="right">
                <asp:Label ID="LabApp4" runat="server" ></asp:Label>
            </td>
            <td style="width: 8%" align="center">
                <asp:Label ID="Label23" runat="server" Text="筆" ></asp:Label>
            </td>
            </tr>
            
            <tr>
            <td style="width: 42%">
                <asp:HyperLink ID="HyperLink5" runat="server" NavigateUrl="~/M_Source/00/MOA00010.aspx?strHyper=U28r13D6EA">會客洽公作業</asp:HyperLink>
            </td>
            <td style="width: 15%" align="right">
                <asp:Label ID="LabApp5" runat="server" ></asp:Label>
            </td>
            <td style="width: 8%" align="center">
                <asp:Label ID="Label30" runat="server" Text="筆" ></asp:Label>
            </td>
            </tr>

            <tr>
            <td style="width: 42%">
                <asp:HyperLink ID="HyperLink11" runat="server" NavigateUrl="~/M_Source/00/MOA00010.aspx?strHyper=S9QR2W8X6U">門禁會議管制作業</asp:HyperLink>
            </td>
            <td style="width: 15%" align="right">
                <asp:Label ID="LabApp6" runat="server" ></asp:Label>
            </td>
            <td style="width: 8%" align="center">
                <asp:Label ID="Label6" runat="server" Text="筆" ></asp:Label>
            </td>
            </tr>
            
            <tr>
            <td style="width: 42%">
                <asp:HyperLink ID="HyperLink13" runat="server" 
                    NavigateUrl="~/M_Source/00/MOA00010.aspx?strHyper=BL7U2QP3IG">資訊設備維修作業</asp:HyperLink>
            </td>
            <td style="width: 15%" align="right">
                <asp:Label ID="LabApp7" runat="server" ></asp:Label>
            </td>
            <td style="width: 8%" align="center">
                <asp:Label ID="Label38" runat="server" Text="筆" ></asp:Label>
            </td>
            </tr>
            
            <tr>
            <td style="width: 42%">
                <asp:Label ID="Label25" runat="server" Text="總計" ></asp:Label>
            </td>
            <td style="width: 15%" align="right">
                <asp:Label ID="LabAppALL" runat="server" ></asp:Label>
            </td>
            <td style="width: 8%" align="center">
                <asp:Label ID="Label27" runat="server" Text="筆" ></asp:Label>
            </td>
            </tr>
        </table>
        <% end if %>
        
        <table width="75%" border="3" bordercolor="#ccddee" align="center">
            <tr>
            <td align="center" class="CellClass" style="width: 35%" rowspan="8">
                <asp:Label ID="Label3" runat="server" Text="申請未完成表單" ></asp:Label>
            </td>
            <td style="width: 42%">
                <asp:HyperLink ID="HyperLink6" runat="server" NavigateUrl="~/M_Source/00/MOA00011.aspx?strHyper=YAqBTxRP8P">差假作業</asp:HyperLink>
            </td>
            <td style="width: 15%" align="right">
                <asp:Label ID="LabWait1" runat="server" Text="" ></asp:Label>
            </td>
            <td style="width: 8%" align="center">
                <asp:Label ID="Label7" runat="server" Text="筆" ></asp:Label>
            </td>
            </tr>
            
            <tr>
            <td style="width: 42%">
                <asp:HyperLink ID="HyperLink7" runat="server" NavigateUrl="~/M_Source/00/MOA00011.aspx?strHyper=4rM2YFP73N">會議室作業</asp:HyperLink>
            </td>
            <td style="width: 15%" align="right">
                <asp:Label ID="LabWait2" runat="server" Text="" ></asp:Label>
            </td>
            <td style="width: 8%" align="center">
                <asp:Label ID="Label20" runat="server" Text="筆" ></asp:Label>
            </td>
            </tr>
            
            <tr>
            <td style="width: 42%">
                <asp:HyperLink ID="HyperLink8" runat="server" NavigateUrl="~/M_Source/00/MOA00011.aspx?strHyper=j2mvKYe3l9">派車作業</asp:HyperLink>
            </td>
            <td style="width: 15%" align="right">
                <asp:Label ID="LabWait3" runat="server" Text="" ></asp:Label>
            </td>
            <td style="width: 8%" align="center">
                <asp:Label ID="Label31" runat="server" Text="筆" ></asp:Label>
            </td>
            </tr>
            
            <tr>
            <td style="width: 42%">
                <asp:HyperLink ID="HyperLink9" runat="server" NavigateUrl="~/M_Source/00/MOA00011.aspx?strHyper=61TY3LELYT">房舍水電修繕作業</asp:HyperLink>
            </td>
            <td style="width: 15%" align="right">
                <asp:Label ID="LabWait4" runat="server" Text="" ></asp:Label>
            </td>
            <td style="width: 8%" align="center">
                <asp:Label ID="Label34" runat="server" Text="筆" ></asp:Label>
            </td>
            </tr>
            
            <tr>
            <td style="width: 42%">
                <asp:HyperLink ID="HyperLink10" runat="server" NavigateUrl="~/M_Source/00/MOA00011.aspx?strHyper=U28r13D6EA">會客洽公作業</asp:HyperLink>
            </td>
            <td style="width: 15%" align="right">
                <asp:Label ID="LabWait5" runat="server" Text="" ></asp:Label>
            </td>
            <td style="width: 8%" align="center">
                <asp:Label ID="Label37" runat="server" Text="筆" ></asp:Label>
            </td>
            </tr>

            <tr>
            <td style="width: 42%">
                <asp:HyperLink ID="HyperLink12" runat="server" NavigateUrl="~/M_Source/00/MOA00011.aspx?strHyper=S9QR2W8X6U">門禁會議管制作業</asp:HyperLink>
            </td>
            <td style="width: 15%" align="right">
                <asp:Label ID="LabWait6" runat="server" Text="" ></asp:Label>
            </td>
            <td style="width: 8%" align="center">
                <asp:Label ID="Label8" runat="server" Text="筆" ></asp:Label>
            </td>
            </tr>
            
            <tr>
            <td style="width: 42%">
                <asp:HyperLink ID="HyperLink14" runat="server" 
                    NavigateUrl="~/M_Source/00/MOA00011.aspx?strHyper=BL7U2QP3IG">資訊設備維修作業</asp:HyperLink>
            </td>
            <td style="width: 15%" align="right">
                <asp:Label ID="LabWait7" runat="server" Text="" ></asp:Label>
            </td>
            <td style="width: 8%" align="center">
                <asp:Label ID="Label39" runat="server" Text="筆" ></asp:Label>
            </td>
            </tr>
            
            <tr>
            <td style="width: 42%">
                <asp:Label ID="Label10" runat="server" Text="總計" ></asp:Label>
            </td>
            <td style="width: 15%" align="right">
                <asp:Label ID="LabWaitALL" runat="server" Text="" ></asp:Label>
            </td>
            <td style="width: 8%" align="center">
                <asp:Label ID="Label15" runat="server" Text="筆" ></asp:Label>
            </td>
            </tr>
        </table>
        
    </form>
</body>
</html>
