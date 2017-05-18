<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA02008.aspx.vb" Inherits="Source_02_MOA02008" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>會議室借用詳細資料</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr>
            <td align="center">
                <asp:Label ID="LabTitle" runat="server" CssClass="toptitle" Text="會議室借用詳細資料" Width="100%"></asp:Label>
            </td>
            </tr>
        </table>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" DataSourceID="SqlDataSource1"
            Width="100%" DataKeyNames="MTT_Num" CellPadding="4" ForeColor="#333333" GridLines="None" AllowPaging="True" AllowSorting="True" >
            <Columns>
                <asp:BoundField DataField="PAUNIT" HeaderText="借用單位" SortExpression="PAUNIT">
                    <HeaderStyle Width="15%" HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="PANAME" HeaderText="借用人" SortExpression="PANAME">
                    <HeaderStyle Width="10%" HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nPHONE" HeaderText="電話" SortExpression="nPHONE">
                    <HeaderStyle Width="8%" HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nMEETNAME" HeaderText="會議名稱" SortExpression="nMEETNAME">
                    <HeaderStyle Width="13%" HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="MeetTime" HeaderText="借用日期" SortExpression="MeetTime">
                    <HeaderStyle Width="13%" HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="MeetHour" HeaderText="時段" SortExpression="MeetHour">
                    <HeaderStyle Width="8%" HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="MTT_Num" HeaderText="MTT_Num" SortExpression="MTT_Num" >
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="註銷理由">
                    <ItemTemplate>
                        <asp:TextBox ID="txtComment" runat="server" Width="90px" MaxLength="100"></asp:TextBox>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                </asp:TemplateField>
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:ImageButton ID="DelBtn" runat="server" CausesValidation="False" CommandName="Delete"
                            ImageUrl="~/Image/writeoff.gif" OnClientClick="return confirm('確定註銷會議室嗎?')" AlternateText="註銷" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="狀態" SortExpression="PENDFLAG">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("PENDFLAG") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# FunStatus("PENDFLAG") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
            </Columns>
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
        </asp:GridView>        
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr>
            <td align="center">
                <asp:ImageButton ID="BackBtn" runat="server" ImageUrl="~/Image/backtop.gif" AlternateText="回上頁" />
            </td>
            </tr>
        </table>
        
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [P_0204] WHERE [MTT_Num] = @MTT_Num" SelectCommand="SELECT P_02.PAUNIT, P_02.PANAME, P_02.nPHONE, P_02.nMEETNAME, P_02.nPLACE, P_0204.MTT_Num, P_0204.EFORMSN, P_0204.MeetSn, CONVERT (char(12), P_0204.MeetTime, 111) AS MeetTime, P_0204.MeetHour FROM P_02,P_0204 WHERE P_02.EFORMSN = P_0204.EFORMSN AND P_0204.MeetSn = @Meetsn and 1=2 order by P_0204.MeetTime DESC">
            <SelectParameters>
                <asp:QueryStringParameter Name="Meetsn" QueryStringField="meetsn" />
            </SelectParameters>
            <DeleteParameters>
                <asp:Parameter Name="MTT_Num" Type="Decimal" />
            </DeleteParameters>
        </asp:SqlDataSource>
    </form>
</body>
</html>
