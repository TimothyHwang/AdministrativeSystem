<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA08014.aspx.vb" Inherits="M_Source_08_MOA08014" %>
<%@ Register Src="../90/FlowRoute.ascx" TagName="FlowRoute" TagPrefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>影印紀錄呈核單明細</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.7.2.min.js" type="text/javascript"></script>
    <script type="text/javascript">
    function checkOpinion(obj){
        var op = $("#" + "<%=txtOpinion.ClientID %>").val();
        if(obj.id == "<%=btnReject.ClientID %>" && op == ""){
            alert("請輸入批核意見");
            return false;
        }
        return true;
    }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
        <tr><td align="center" class = "toptitle">
        影印紀錄呈核單明細</td></tr>
    </table>
    <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
            <tr>
                <td nowrap="nowrap" width="10%" class="form" align = "left">
                    批核意見：
                </td>
                <td nowrap="nowrap" class="form" align = "left" width="50%">
                    <asp:TextBox ID="txtOpinion" runat="server"></asp:TextBox>
                    <asp:Label ID="lbOpinion" runat="server" Visible="False"></asp:Label>
                </td>
                <td class="form" colspan ="2" align="right" width="40%">
                    <asp:Button ID="btnApprove" runat="server" Text="核准" OnClientClick="return checkOpinion(this);" />&nbsp;&nbsp;
                    <asp:Button ID="btnReject" runat="server" Text="駁回" OnClientClick="return checkOpinion(this);" />&nbsp;&nbsp;
                </td>
            </tr>
        </table>
    <asp:GridView ID="gvPrintReports" runat="server" EmptyDataText="查無任何資料" 
            AutoGenerateColumns="False" AllowPaging="True"
         AllowSorting="True" Width="100%" 
            CellPadding="10" ForeColor="#333333" GridLines="None" 
        EnableModelValidation="True" HorizontalAlign="Center">

            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" 
                HorizontalAlign="Center" VerticalAlign="Middle" />
            <Columns>
                <asp:BoundField DataField="File_Name" HeaderText="複印資料名稱" >
                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="9%"/>
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="Print_Date" HeaderText="複印時間" >
                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="9%"/>
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="ORG_NAME" HeaderText="使用單位" >
                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="9%"/>
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="Print_Name" HeaderText="姓名" >
                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="9%"/>
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="密等">
                    <ItemTemplate>
                        <%# displaySecurityStatus(Eval("Security_Status"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="9%"/>
                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:TemplateField>
                <asp:BoundField DataField="Total_sheet" HeaderText="申請張數">
                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="9%"/>
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="實際張數">
                    <ItemTemplate>
                        <%# displayCopyDetail(Eval("Copy_A3M"),Eval("Copy_A4M"),Eval("Copy_A3C"),Eval("Copy_A4C"),Eval("Scan"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="10%"/>
                </asp:TemplateField>
                <asp:BoundField DataField="Use_For" HeaderText="用途" >
                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="9%"/>
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="Print_Num" HeaderText="流水號" >
                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="9%"/>
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="批核">
                    <ItemTemplate>
                        <%# ShowApproveStatus(Eval("VerifyRequesterID"), Eval("ApprovedByID"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="9%"/>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="影印明細">
                    <ItemTemplate>
                        <asp:HiddenField ID="HF_Status" runat="server" Value = '<%# Eval("Status")%>' />
                        <asp:ImageButton ID="ImgDetail" runat="server" CausesValidation="False" CommandArgument ='<%# Eval("Log_Guid") %>' CommandName="Detail" ImageUrl="~/Image/List.gif" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="9%"/>
                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:TemplateField>
            </Columns>
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
        </asp:GridView>
        <center><uc1:FlowRoute ID="FlowRoute1" runat="server" /></center>
    </form>
</body>
</html>
