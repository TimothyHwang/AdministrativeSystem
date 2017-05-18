<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA09004.aspx.vb" Inherits="M_Source_09_MOA09004" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>門禁會議管制部門統計列表</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="lbDeptStatistics" runat="server" CssClass="toptitle" Width="100%"></asp:Label>
            </td></tr>
        </table>
        <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
            <tr>
                <td nowrap="nowrap" width="5%" class="form" align = "right">
                <asp:ImageButton ID="ImgBtn1" runat="server" ImageUrl="../../Image/backtop.gif" ToolTip="返回上一頁" />
                </td>
            </tr>
        </table>
        -
        <asp:GridView ID="gvP_09" runat="server" EmptyDataText="查無任何資料" 
            AutoGenerateColumns="False" AllowPaging="True"
         AllowSorting="True" Width="100%" 
            CellPadding="10" ForeColor="#333333" GridLines="None" 
            EnableModelValidation="True">

            <Columns>
                <asp:TemplateField HeaderText="開會日期" >
                    <ItemTemplate><%# ShowMeetingDate(Eval("MeetingDate"))%></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="進出營門">
                    <ItemTemplate><%# ShowEnteringGate(Eval("EFORMSN"))%></ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="Subject" HeaderText="開會事由"></asp:BoundField>
                <asp:BoundField DataField="Location" HeaderText="地點"></asp:BoundField>
                <asp:TemplateField HeaderText="進出人員">
                    <ItemTemplate><%# Eval("EnteringPeopleNumber")%> 員</ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="Moderator" HeaderText="主持人" ></asp:BoundField>
                <asp:TemplateField HeaderText="聯絡人及電話">
                    <ItemTemplate>
                        <table border="0" width="100%">
                            <tr><td width="100%" align="center"><%# Eval("ORG_NAME")%></td></tr>
                            <tr><td width="100%" align="center"><%# Eval("emp_chinese_name")%>&nbsp;<%# Eval("PhoneNumber")%></td></tr>
                        </table>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="DocumentNo" HeaderText="發文字號" ></asp:BoundField>
                <asp:TemplateField HeaderText="狀態" >
                    <ItemTemplate><%# ShowStatus(Eval("Status"))%></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:ImageButton ID="ImgDetail" runat="server" CausesValidation="False" CommandArgument ='<%# Eval("EFORMSN") %>' CommandName="Detail" ImageUrl="~/Image/List.gif" />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
        </asp:GridView>
    </form>
</body>
</html>
