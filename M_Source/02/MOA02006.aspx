<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA02006.aspx.vb" Inherits="Source_02_MOA02006" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>會議室設備管理</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />  
</head>
<body>
    <form id="form1" runat="server">
    <div>  
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label6" runat="server" CssClass="toptitle" Text="會議室設備管理" Width="100%"></asp:Label>
            </td></tr>
        </table>         
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap class="form" style="height: 25px; width: 10%;">
                    <asp:Label ID="Label1" runat="server" CssClass="form" Text="設備名稱:" Width="77px"></asp:Label></td>
			    <td noWrap width="70%" class="form" style="height: 25px">			    
                        <asp:TextBox ID="DevName" runat="server" Width="212px"></asp:TextBox></td>
                <td noWrap width="20%" class="form" style="height: 25px">
			    <asp:ImageButton ID="ImgBtn1" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" /></td>
		    </tr>
	    </table>
    
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
            AutoGenerateColumns="False" CellPadding="4" DataKeyNames="DeviceSn" DataSourceID="SqlDataSource1"
            ForeColor="#333333" GridLines="None" HorizontalAlign="Left" Width="100%">
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <Columns>
                <asp:BoundField DataField="DeviceSn" HeaderText="DeviceSn" InsertVisible="False"
                    ReadOnly="True" SortExpression="DeviceSn" Visible="False" />
                <asp:TemplateField HeaderText="設備名稱" SortExpression="DeviceName">
                    <EditItemTemplate>
                        <asp:TextBox ID="DevName" runat="server" Text='<%# Bind("DeviceName") %>' Width="185px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="DevName"
                            ErrorMessage="請輸入設備名稱" Width="119px"></asp:RequiredFieldValidator>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("DeviceName") %>'></asp:Label>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="80%" />
                </asp:TemplateField>
                <asp:TemplateField ShowHeader="False">
                    <EditItemTemplate>
                        <asp:ImageButton ID="btnImgOK" runat="server" CommandName="Update" ImageUrl="~/Image/apply.gif"
                            ToolTip="確認" />&nbsp;<asp:ImageButton ID="btnImgCancel" runat="server" CommandName="Cancel"
                                ImageUrl="~/Image/cancel.gif" ToolTip="取消" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        &nbsp;<asp:ImageButton ID="ImgUpd" runat="server" CommandName="Edit" ImageUrl="~/Image/update.gif" ToolTip="修改" />
                        <asp:ImageButton ID="ImgDel" runat="server" CommandName="Delete" ImageUrl="~/Image/delete.gif"
                            OnClientClick="return confirm('確定刪除嗎?')" ToolTip="刪除" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:TemplateField>
            </Columns>
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [P_0202] WHERE [DeviceSn] = @DeviceSn" InsertCommand="INSERT INTO [P_0202] ([DeviceName]) VALUES (@DeviceName)"
            SelectCommand="SELECT [DeviceSn], [DeviceName] FROM [P_0202] ORDER BY [DeviceName]"
            UpdateCommand="UPDATE [P_0202] SET [DeviceName] = @DeviceName WHERE [DeviceSn] = @DeviceSn">
            <DeleteParameters>
                <asp:Parameter Name="DeviceSn" Type="Decimal" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="DeviceName" Type="String" />
                <asp:Parameter Name="DeviceSn" Type="Decimal" />
            </UpdateParameters>
            <InsertParameters>
                <asp:Parameter Name="DeviceName" Type="String" />
            </InsertParameters>
        </asp:SqlDataSource>
    
    </div>
    </form>
</body>
</html>
