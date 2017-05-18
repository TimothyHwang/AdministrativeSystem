<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA05004.aspx.vb" Inherits="Source_05_MOA05004" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>憲兵換證作業</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="left" style="width:3%" >
                <a href="javascript:window.history.go(<%= HistoryGo %>);">
                <img  align='absmiddle' title="回上頁" border="0" src="../../Image/backtop.gif" class="btn_image"/></a>
            </td><td align="center" style="width:97%" >
                    <asp:Label ID="Label2" runat="server" CssClass="toptitle" Text="憲兵換證作業" Width="100%"></asp:Label>
            </td></tr>
        </table>
        <table width="100%">
            <tr><td>
                    <asp:DetailsView ID="DetailsView1" runat="server" DataSourceID="SqlDataSource1"  DataKeyNames="EFORMSN" 
                        Height="50px" Width="100%" AutoGenerateRows="False">
                        <Fields>
                            <asp:BoundField DataField="PAUNIT" HeaderText="申請人單位" ReadOnly="True" SortExpression="PAUNIT" >
                                <HeaderStyle Width="30%" />
                            </asp:BoundField>
                            <asp:BoundField DataField="PATITLE" HeaderText="申請人級職" ReadOnly="True" SortExpression="PATITLE" >
                                <HeaderStyle Width="30%" />
                            </asp:BoundField>
                            <asp:BoundField DataField="PANAME" HeaderText="申請人姓名" ReadOnly="True" SortExpression="PANAME" >
                                <HeaderStyle Width="30%" />
                            </asp:BoundField>
                            <asp:BoundField DataField="nRECROOM" HeaderText="會客室名稱" ReadOnly="True" SortExpression="nRECROOM" >
                                <HeaderStyle Width="30%" />
                            </asp:BoundField>
                            <asp:BoundField DataField="nRECEXIT" HeaderText="會客入口名稱" ReadOnly="True" SortExpression="nRECEXIT" >
                                <HeaderStyle Width="30%" />
                            </asp:BoundField>
                            <asp:BoundField DataField="nPLACE" HeaderText="會客地點" ReadOnly="True" SortExpression="nPLACE" >
                                <HeaderStyle Width="30%" />
                            </asp:BoundField>
                            <asp:BoundField DataField="nRECDATE" HeaderText="會客時間" ReadOnly="True" SortExpression="nRECDATE" >
                                <HeaderStyle Width="30%" />
                            </asp:BoundField>
                            <asp:BoundField DataField="nSTARTTIME" HeaderText="會客時間(起始 - 時)" ReadOnly="True" SortExpression="nSTARTTIME" >
                                <HeaderStyle Width="30%" />
                            </asp:BoundField>
                            <asp:BoundField DataField="nENDTIME" HeaderText="會客時間(結束 - 時)" ReadOnly="True" SortExpression="nENDTIME" >
                                <HeaderStyle Width="30%" />
                            </asp:BoundField>
                            <asp:BoundField DataField="nREASON" HeaderText="事由" ReadOnly="True" SortExpression="nREASON" >
                                <HeaderStyle Width="30%" />
                            </asp:BoundField>
                        </Fields>
                    </asp:DetailsView>
                    <asp:GridView ID="GridView1" runat="server" AllowPaging="True" 
                        AllowSorting="True" AutoGenerateColumns="False" DataKeyNames="Receive_Num,EFORMSN" 
                        DataSourceID="SqlDataSource2" Width="100%" CellPadding="4" 
                        ForeColor="#333333" GridLines="None" PageSize="10" EnableModelValidation="True">
                        <Columns>
                            <asp:BoundField DataField="nName" HeaderText="姓名" ReadOnly="True" SortExpression="nName" >
                                <ItemStyle HorizontalAlign="Center" />
                                <HeaderStyle HorizontalAlign="Center" Width="10%" />
                            </asp:BoundField>
                            <asp:BoundField DataField="nSex" HeaderText="性別" ReadOnly="True" SortExpression="nSex" >
                                <ItemStyle HorizontalAlign="Center" />
                                <HeaderStyle HorizontalAlign="Center" Width="8%" />
                            </asp:BoundField>
                            <asp:BoundField DataField="nService" HeaderText="服務單位" ReadOnly="True" SortExpression="nService" >
                                <ItemStyle HorizontalAlign="Center" />
                                <HeaderStyle HorizontalAlign="Center" Width="18%" />
                            </asp:BoundField>
                            <asp:BoundField DataField="nID" HeaderText="身分證字號" ReadOnly="True" SortExpression="nID" >
                                <ItemStyle HorizontalAlign="Center" />
                                <HeaderStyle HorizontalAlign="Center" Width="15%" />
                            </asp:BoundField>
                            <asp:TemplateField HeaderText="車號">
                                <EditItemTemplate>
                                    <asp:Label ID="lblCarNo" runat="server" Text='<%# Bind("nCarNo") %>'></asp:Label>
                                </EditItemTemplate>
                                <ItemTemplate>
                                    <asp:Label ID="lblCarNo" runat="server" Text='<%# Bind("nCarNo") %>'></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" Width="100px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="證號" SortExpression="nPassID">
                                <EditItemTemplate>
                                    <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("nPassID") %>' Width="30px" MaxLength="3"></asp:TextBox>
                                </EditItemTemplate>
                                <ItemTemplate>
                                    <asp:Label ID="Label1" runat="server" Text='<%# Bind("nPassID") %>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" Width="8%" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="進入時間">
                                <EditItemTemplate>
                                    <asp:Label ID="nInDate" runat="server" Text='<%# Bind("nInDate", "{0:yyyy/MM/dd hh:mm:ss}") %>'
                                        Width="77px"></asp:Label>
                                    <asp:Button ID="btn_nInDate" runat="server" OnClick="btn_nInDate_Click" Text="進入" />
                                </EditItemTemplate>
                                <ItemTemplate>
                                    <asp:Label ID="nInDate" runat="server" Text='<%# Bind("nInDate", "{0:yyyy/MM/dd HH:mm:ss}") %>' Width="79px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" />
                                <HeaderStyle HorizontalAlign="Center" Width="18%" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="離開時間">
                                <EditItemTemplate>
                                    <asp:Label ID="nLeaveDate" runat="server" Text='<%# Bind("nLeaveDate", "{0:yyyy/MM/dd HH:mm:ss}") %>'
                                        Width="77px"></asp:Label>
                                    <asp:Button ID="btn_nLeaveDate" runat="server" OnClick="btn_nLeaveDate_Click" Text="離開" />
                                </EditItemTemplate>
                                <ItemTemplate>
                                    <asp:Label ID="nLeaveDate" runat="server" Text='<%# Bind("nLeaveDate", "{0:yyyy/MM/dd HH:mm:ss}") %>' Width="79px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" />
                                <HeaderStyle HorizontalAlign="Center" Width="18%" />
                            </asp:TemplateField>
                    <asp:TemplateField ShowHeader="False">
                        <EditItemTemplate>
                            <asp:ImageButton ID="btnImgOK" runat="server" CommandName="Update" ImageUrl="~/Image/apply.gif" 
                                ToolTip="確認" />&nbsp;<asp:ImageButton ID="btnImgCancel" runat="server" CommandName="Cancel"
                                    ImageUrl="~/Image/cancel.gif" ToolTip="取消" />
                        </EditItemTemplate>
                        <ItemTemplate>
                            <asp:ImageButton ID="ImgUpd" runat="server" CommandName="Edit" ImageUrl="~/Image/update.gif"
                                ToolTip="修改" />
                        </ItemTemplate>
                        <HeaderStyle Width="10%" />
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
            </td></tr>
            
        </table>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand=""
            UpdateCommand=""
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand=""
            UpdateCommand=""
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
                <UpdateParameters>
                    <asp:Parameter Name="nPassID" Type="String" />
                    <asp:Parameter Name="nInDate" Type="DateTime" />
                    <asp:Parameter Name="nLeaveDate"  Type="DateTime" />
                    <asp:Parameter Name="Receive_Num" Type="String" />
                    <asp:Parameter Name="EFORMSN" Type="String" />
                </UpdateParameters>
        </asp:SqlDataSource>
    </form>
</body>
</html>
