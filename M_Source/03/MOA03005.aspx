<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA03005.aspx.vb" Inherits="Source_03_MOA03005" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>車輛登錄</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="車輛登錄" Width="100%"></asp:Label>
            </td></tr>
        </table>
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap width="10%" class="form" style="height: 25px">
                    車種：</td>
			    <td noWrap width="10%" class="form" style="height: 25px">
                    <asp:DropDownList ID="PCK_Num" runat="server" DataSourceID="SqlDataSource3" DataTextField="PCK_Name" DataValueField="PCK_Num">
                    </asp:DropDownList>
			    </td>
			    <td noWrap width="10%" class="form" style="height: 25px">
                    車號：</td>
			    <td noWrap width="10%" class="form" style="height: 25px">
                    <asp:TextBox ID="PCI_CarNumber" runat="server" Width="80px"></asp:TextBox>
                </td>
                <td noWrap width="10%" class="form" style="height: 25px">
                    車輛類型：</td>
			    <td noWrap width="10%" class="form" style="height: 25px">
                    <asp:DropDownList ID="PCI_Status" runat="server">
                        <asp:ListItem></asp:ListItem>
                        <asp:ListItem Value="1"> 一般車 </asp:ListItem>
                        <asp:ListItem Value="2"> 經常性支援 </asp:ListItem>
                        <asp:ListItem Value="3"> 長官車 </asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td noWrap width="10%" class="form" style="height: 25px">
                    專用車長官：</td>
			    <td noWrap width="10%" class="form" style="height: 25px">
                    <asp:TextBox ID="PCI_OnlySir" runat="server" Width="80px"></asp:TextBox>
                </td>
                <td noWrap width="10%" class="form" style="height: 25px">
                    使用狀況：</td>
			    <td noWrap width="10%" class="form" style="height: 25px">
                    <asp:DropDownList ID="PCI_Use" runat="server">
                        <asp:ListItem></asp:ListItem>
                            <asp:ListItem Value="1"> 待命 </asp:ListItem>
                            <asp:ListItem Value="2"> 派遣 </asp:ListItem>
                            <asp:ListItem Value="3"> 維修 </asp:ListItem>
                    </asp:DropDownList>
                </td>
			    <td noWrap width="10%" class="form" style="height: 25px">
			    <asp:ImageButton ID="ImgBtn1" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" />
                <asp:ImageButton ID="btnSelect" runat="server" ImageUrl="~/Image/Search.gif" ToolTip="查詢" /></td>                
		    </tr>
	    </table>
		<asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" DataKeyNames="PCI_Num" DataSourceID="SqlDataSource1" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
            <Columns>
                <asp:templatefield HeaderText="車種"  SortExpression="PCK_Num">
                    <itemtemplate>
                        <asp:label id="LastNameLabel"
                        text= '<%# Eval("PCK_Name") %>'
                        runat="server"/>
                    </itemtemplate>
                    <EditItemTemplate>
                        <asp:DropDownList id="PCK_Num"  SelectedValue='<%# Bind("PCK_Num") %>'
                            AutoPostBack="True"
                            DataSourceID="SqlDataSource2"
                            DataValueField="PCK_Num"
                            DataTextField="PCK_Name"
                            runat="server">
                        </asp:DropDownList>
                    </EditItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:templatefield>
                <asp:BoundField  DataField="PCI_CarNumber" HeaderText="車號" SortExpression="PCI_CarNumber">
                    <ItemStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:templatefield HeaderText="車輛類型"  SortExpression="PCI_Status">
                    <itemtemplate>
                        <asp:label id="LastNameLabe2"
                        text= '<%# Chg_Status("PCI_Status") %>'
                        runat="server"/>
                    </itemtemplate>
                    <EditItemTemplate>
                        <asp:DropDownList id="PCI_Status2" SelectedValue='<%# Bind("PCI_Status") %>'
                            AutoPostBack="True"
                            runat="server" >
                            <asp:ListItem Value=""></asp:ListItem>
                            <asp:ListItem Value="1"> 一般車 </asp:ListItem>
                            <asp:ListItem Value="2"> 經常性支援 </asp:ListItem>
                            <asp:ListItem Value="3"> 長官車 </asp:ListItem>
                        </asp:DropDownList>
                    </EditItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:templatefield>
                <asp:BoundField  DataField="PCI_OnlySir" HeaderText="專用車長官" SortExpression="PCI_OnlySir">
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:templatefield HeaderText="使用狀況"  SortExpression="PCI_Use">
                    <itemtemplate>
                        <asp:label id="LastNameLabe3"
                        text= '<%# Chg_Use("PCI_Use") %>'
                        runat="server"/>
                    </itemtemplate>
                    <EditItemTemplate>
                        <asp:DropDownList id="PCI_Use2" SelectedValue='<%# Bind("PCI_Use") %>'
                            AutoPostBack="True"
                            runat="server" >
                            <asp:ListItem Value=""></asp:ListItem>
                            <asp:ListItem Value="1"> 待命 </asp:ListItem>
                            <asp:ListItem Value="2"> 派遣 </asp:ListItem>
                            <asp:ListItem Value="3"> 維修 </asp:ListItem>
                        </asp:DropDownList>
                    </EditItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="10%" />
                </asp:templatefield>
                <asp:TemplateField ShowHeader="False">
                    <EditItemTemplate>
                        <asp:ImageButton ID="btnImgOK" runat="server" CommandName="Update" ImageUrl="~/Image/apply.gif"
                            ToolTip="確認" />&nbsp;<asp:ImageButton ID="btnImgCancel" runat="server" CommandName="Cancel"
                                ImageUrl="~/Image/cancel.gif" ToolTip="取消" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:ImageButton ID="ImgUpd" runat="server" CommandName="Edit" ImageUrl="~/Image/update.gif"
                            ToolTip="修改" />&nbsp;<asp:ImageButton ID="ImgDel" runat="server" CommandName="Delete"
                                ImageUrl="~/Image/delete.gif" OnClientClick="return confirm('確定刪除嗎?')" ToolTip="刪除" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
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
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [P_0303] WHERE [PCI_Num] = @PCI_Num" 
            InsertCommand="INSERT INTO [P_0303] ([PCK_Num], [PCI_CarNumber],[PCI_Status],[PCI_OnlySir],[PCI_Use]) VALUES (@PCK_Num, @PCI_CarNumber, @PCI_Status,@PCI_OnlySir,@PCI_Use)"
            SelectCommand="SELECT [PCI_Num], [P_0303].[PCK_Num], [PCI_CarNumber], [PCI_Status], [PCI_OnlySir],[PCK_Name],[PCI_Use] FROM [P_0303],[P_0302] where [P_0303].[PCK_Num]=[P_0302].[PCK_Num] ORDER BY [PCI_Num]"
            UpdateCommand="UPDATE [P_0303] SET [PCK_Num] = @PCK_Num, [PCI_CarNumber] = @PCI_CarNumber,[PCI_Status]=@PCI_Status,[PCI_OnlySir]=@PCI_OnlySir,[PCI_Use]=@PCI_Use WHERE [PCI_Num] = @PCI_Num"
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
            <DeleteParameters>
                <asp:Parameter Name="PCI_Num" Type="Int32" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="PCK_Num" Type="Int32" />
                <asp:Parameter Name="PCI_CarNumber" Type="String" />
                <asp:Parameter Name="PCI_Status"  Type="Int32" />
                <asp:Parameter Name="PCI_OnlySir"  Type="String" />
                <asp:Parameter Name="PCI_Use" Type="String" />
                <asp:Parameter Name="PCI_Num" Type="Int32" />
            </UpdateParameters>
            <InsertParameters>
                <asp:Parameter Name="PCK_Num" Type="Int32" />
                <asp:Parameter Name="PCI_CarNumber" Type="String" />
                <asp:Parameter Name="PCI_Status"  Type="Int32" />
                <asp:Parameter Name="PCI_OnlySir"  Type="String" />
                <asp:Parameter Name="PCI_Use" Type="String" />
            </InsertParameters>
        </asp:SqlDataSource>        
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="select -1 as PCK_Num,'' as PCK_Name union SELECT [PCK_Num], [PCK_Name] FROM [P_0302] ORDER BY [PCK_Num]"
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
        </asp:SqlDataSource>        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="select -1 as PCK_Num,'' as PCK_Name union SELECT [PCK_Num], [PCK_Name] FROM [P_0302] ORDER BY [PCK_Num]"
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>"></asp:SqlDataSource>
    </form>
</body>
</html>
