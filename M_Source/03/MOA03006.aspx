<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA03006.aspx.vb" Inherits="Source_03_MOA03006" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>車輛新增</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="車輛新增" Width="100%"></asp:Label>
            </td></tr>
        </table>
        <asp:DetailsView ID="DetailsView1" runat="server" AutoGenerateRows="False" CssClass="form"
            DataKeyNames="PCI_Num" DataSourceID="SqlDataSource1" DefaultMode="Insert" Height="50px"
            Width="100%">
            <Fields>
                <asp:TemplateField ShowHeader="False" SortExpression="PCI_Num">
                    <InsertItemTemplate>
                        <table align="center" border="3" bordercolor="#ccddee" width="100%">
                            <tr>
                                <td style="width: 20%">
                        <asp:Label ID="Lab1" runat="server" Text="車種："></asp:Label></td>
                                <td style="width: 80%">
                        <asp:DropDownList ID="PCK_Num" runat="server" DataSourceID="SqlDataSource2" DataTextField="PCK_Name"
                            DataValueField="PCK_Num" SelectedValue='<%# Bind("PCK_Num") %>'>
                            <asp:ListItem></asp:ListItem>
                        </asp:DropDownList></td>
                            </tr>
                            <tr>
                                <td style="width: 20%">
                        <asp:Label ID="Lab2" runat="server" Text="車號："></asp:Label></td>
                                <td style="width: 80%">
                        <asp:TextBox ID="PCI_CarNumber" runat="server" MaxLength="10" Text='<%# Bind("PCI_CarNumber") %>'
                            Width="80px"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td style="width: 20%">
                        <asp:Label ID="Lab3" runat="server" Text="車輛類型："></asp:Label></td>
                                <td style="width: 80%">
                        <asp:DropDownList ID="PCI_Status" runat="server" SelectedValue='<%# Bind("PCI_Status") %>'>
                            <asp:ListItem></asp:ListItem>
                            <asp:ListItem Value="1"> 一般車 </asp:ListItem>
                            <asp:ListItem Value="2"> 經常性支援 </asp:ListItem>
                            <asp:ListItem Value="3"> 長官車 </asp:ListItem>
                        </asp:DropDownList></td>
                            </tr>
                            <tr>
                                <td style="width: 20%">
                        <asp:Label ID="Label4" runat="server" Text="專用車長官："></asp:Label></td>
                                <td style="width: 80%">
                        <asp:TextBox ID="PCI_OnlySir" runat="server" MaxLength="50" Text='<%# Bind("PCI_OnlySir") %>'
                            Width="80px"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td style="width: 20%">
                        <asp:Label ID="Label1" runat="server" Text="使用狀況："></asp:Label></td>
                                <td style="width: 80%">
                        <asp:DropDownList ID="PCI_Use" runat="server" SelectedValue='<%# Bind("PCI_Use") %>'>
                            <asp:ListItem></asp:ListItem>
                            <asp:ListItem Value="1"> 待命 </asp:ListItem>
                            <asp:ListItem Value="2"> 派遣 </asp:ListItem>
                            <asp:ListItem Value="3"> 維修 </asp:ListItem>
                        </asp:DropDownList></td>
                            </tr>                            
                            <tr>
                                <td colspan ="2" align="center">
                                <asp:Button ID="btnInsert" runat="server" OnClick="btnInsert_Click" Text="新增" />
                                <asp:Label ID="ErrMsg" runat="server" ForeColor="Red" Text=""></asp:Label>
                                </td>
                            </tr>
                        </table>
                    </InsertItemTemplate>
                </asp:TemplateField>
                <asp:CommandField />
            </Fields>
        </asp:DetailsView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [P_0303] WHERE [PCI_Num] = @PCI_Num" InsertCommand="INSERT INTO [P_0303] ([PCK_Num], [PCI_CarNumber],[PCI_Status],[PCI_OnlySir],[PCI_Use]) VALUES (@PCK_Num, @PCI_CarNumber, @PCI_Status,@PCI_OnlySir,@PCI_Use)"
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>" SelectCommand="SELECT [PCI_Num], [P_0303].[PCK_Num], [PCI_CarNumber], [PCI_Status], [PCI_OnlySir],[PCK_Name],[PCI_Use] FROM [P_0303],[P_0302] where [P_0303].[PCK_Num]=[P_0302].[PCK_Num] ORDER BY [PCI_Num]"
            UpdateCommand="UPDATE [P_0303] SET [PCK_Num] = @PCK_Num, [PCI_CarNumber] = @PCI_CarNumber,[PCI_Status]=@PCI_Status,[PCI_OnlySir]=@PCI_OnlySir,[PCI_Use]=@PCI_Use WHERE [PCI_Num] = @PCI_Num">
            <DeleteParameters>
                <asp:Parameter Name="PCI_Num" Type="Int32" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="PCK_Num" Type="Int32" />
                <asp:Parameter Name="PCI_CarNumber" Type="String" />
                <asp:Parameter Name="PCI_Status" Type="Int32" />
                <asp:Parameter Name="PCI_OnlySir" Type="String" />
                <asp:Parameter Name="PCI_Use" Type="String" />
                <asp:Parameter Name="PCI_Num" Type="Int32" />
            </UpdateParameters>
            <InsertParameters>
                <asp:Parameter Name="PCK_Num" Type="Int32" />
                <asp:Parameter Name="PCI_CarNumber" Type="String" />
                <asp:Parameter Name="PCI_Status" Type="Int32" />
                <asp:Parameter Name="PCI_OnlySir" Type="String" />
                <asp:Parameter Name="PCI_Use" Type="String" />
            </InsertParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>" SelectCommand="select -1 as PCK_Num,'' as PCK_Name union SELECT [PCK_Num], [PCK_Name] FROM [P_0302] ORDER BY [PCK_Num]">
        </asp:SqlDataSource>
    </form>
</body>
</html>
