<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA02004.aspx.vb" Inherits="Source_02_MOA02004" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>會議室管理</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>     
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label1" runat="server" CssClass="toptitle" Text="會議室管理" Width="100%"></asp:Label>
            </td></tr>
        </table>         
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap width="10%" class="form" style="height: 25px">
                    會議室名稱：</td>
			    <td noWrap width="40%" class="form" style="height: 25px"><asp:TextBox ID="meetname" runat="server" Width="120px"></asp:TextBox></td>
			    <td noWrap width="10%" class="form" style="height: 25px">
                    管理者：</td>
			    <td noWrap width="20%" class="form" style="height: 25px">
                    <asp:TextBox ID="Owner" runat="server" Width="120px"></asp:TextBox>
                </td>
			    <td noWrap width="20%" class="form" style="height: 25px">
			    <asp:ImageButton ID="ImgBtn1" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" />
                <asp:ImageButton ID="ImgBtn2" runat="server" ImageUrl="~/Image/Search.gif" ToolTip="查詢" /></td>                
		    </tr>
	    </table>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
            AutoGenerateColumns="False" CellPadding="4" DataKeyNames="MeetSn" 
            DataSourceID="SqlDataSource1" EmptyDataText="查無任何資料"
            ForeColor="#333333" GridLines="None" Width="100%" 
            EnableModelValidation="True">
            <Columns>
                <asp:BoundField DataField="MeetSn" HeaderText="MeetSn" InsertVisible="False" ReadOnly="True"
                    ShowHeader="False" SortExpression="MeetSn" />
                <asp:TemplateField HeaderText="會議室名稱" SortExpression="MeetName">
                    <EditItemTemplate>
                        <asp:TextBox ID="MeetName" runat="server" Text='<%# Bind("MeetName") %>' Width="150px"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label5" runat="server" Text='<%# Bind("MeetName") %>'></asp:Label>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="管理者" SortExpression="Owner">
                    <EditItemTemplate>
                        <asp:DropDownList ID="DropDownList2" runat="server" DataSourceID="SqlDataSource2"
                            DataTextField="emp_chinese_name" DataValueField="employee_id" SelectedValue='<%# Bind("Owner") %>'>
                        </asp:DropDownList>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# Bind("emp_chinese_name") %>'></asp:Label>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="電話" SortExpression="Tel">
                    <EditItemTemplate>
                        <asp:TextBox ID="Tel" runat="server" Text='<%# Bind("Tel") %>' Width="100px"></asp:TextBox><br />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label4" runat="server" Text='<%# Bind("Tel") %>'></asp:Label>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="容納人數" SortExpression="ContainNum">
                    <EditItemTemplate>
                        <asp:TextBox ID="UsePer" runat="server" Text='<%# Bind("ContainNum") %>' Width="58px"></asp:TextBox>
                        <asp:RangeValidator ID="RangeValidator1" runat="server" ErrorMessage="請輸入數字" ControlToValidate="UsePer" MaximumValue="9999" MinimumValue="0" Width="80px"></asp:RangeValidator>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label3" runat="server" Text='<%# Bind("ContainNum") %>'></asp:Label>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="類型" SortExpression="Share">
                    <EditItemTemplate>
                        <asp:DropDownList ID="DropDownList1" runat="server" SelectedValue='<%# Bind("share") %>'>
                            <asp:ListItem Selected="True" Value="1">共用</asp:ListItem>
                            <asp:ListItem Value="2">不共用</asp:ListItem>
                        </asp:DropDownList>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# FunShare("Share") %>'></asp:Label>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="是否啟用" SortExpression="Enabled">
                    <EditItemTemplate>
                        <asp:CheckBox ID="chkEnabled" runat="server" Checked='<%# Bind("Enabled") %>' />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label6" runat="server" Text='<%# Bind("Enabled") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:CommandField ButtonType="Image" SelectText="詳細資料" ShowSelectButton="True" SelectImageUrl="~/Image/List.gif" HeaderText="設備" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                </asp:CommandField>
                <asp:TemplateField ShowHeader="False">
                    <EditItemTemplate>
                        <asp:ImageButton ID="btnImgOK" runat="server" CommandName="Update" ImageUrl="~/Image/apply.gif"
                            ToolTip="確認" />&nbsp;<asp:ImageButton ID="btnImgCancel" runat="server" CommandName="Cancel"
                                ImageUrl="~/Image/cancel.gif" ToolTip="取消" /><br />
                        <asp:Label ID="ErrMsg" runat="server" ForeColor="Red" Text=""></asp:Label>
                    </EditItemTemplate>
                    <ItemTemplate>
                        &nbsp;<asp:ImageButton ID="ImgUpd" runat="server" CommandName="Edit" ImageUrl="~/Image/update.gif" ToolTip="修改" />
                        <asp:ImageButton ID="ImgDel" runat="server" CommandName="Delete" 
                            ImageUrl="~/Image/delete.gif" OnClientClick="return confirm('確定刪除嗎?')" 
                            ToolTip="刪除" Visible="False" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
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
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [P_0201] WHERE [MeetSn] = @MeetSn" InsertCommand="INSERT INTO [P_0201] ([Org_Uid], [MeetName], [Owner], [Tel], [ContainNum], [Share]) VALUES (@Org_Uid, @MeetName, @Owner, @Tel, @ContainNum, @Share)"
            SelectCommand="SELECT MeetSn, P_0201.Org_Uid, MeetName, Owner, Tel, ContainNum, Share,emp_chinese_name,Enabled FROM P_0201,EMPLOYEE WHERE employee_id = P_0201.Owner AND (P_0201.Org_Uid = @Org_Uid) ORDER BY [MeetName]"
            
            UpdateCommand="UPDATE P_0201 SET MeetName = @MeetName, Owner = @Owner, Tel = @Tel, ContainNum = @ContainNum, Share = @Share, Enabled = @Enabled WHERE (MeetSn = @MeetSn)">
            <DeleteParameters>
                <asp:Parameter Name="MeetSn" Type="Decimal" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="MeetName" Type="String" />
                <asp:Parameter Name="Owner" Type="String" />
                <asp:Parameter Name="Tel" Type="String" />
                <asp:Parameter Name="ContainNum" Type="Int32" />
                <asp:Parameter Name="Share" Type="Int32" />
                <asp:Parameter Name="Enabled" />
                <asp:Parameter Name="MeetSn" Type="Decimal" />
            </UpdateParameters>
            <InsertParameters>
                <asp:Parameter Name="Org_Uid" Type="String" />
                <asp:Parameter Name="MeetName" Type="String" />
                <asp:Parameter Name="Owner" Type="String" />
                <asp:Parameter Name="Tel" Type="String" />
                <asp:Parameter Name="ContainNum" Type="Int32" />
                <asp:Parameter Name="Share" Type="Int32" />
            </InsertParameters>
            <SelectParameters>
                <asp:SessionParameter Name="Org_Uid" SessionField="ORG_UID" Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
    
    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                            SelectCommand="SELECT [employee_id], [emp_chinese_name] FROM [V_EmpInfo] WHERE ([ORG_UID] = @ORG_UID) ORDER BY [emp_chinese_name]">
                            <SelectParameters>
                                <asp:SessionParameter Name="ORG_UID" SessionField="org_uid" Type="String" />
                            </SelectParameters>
                        </asp:SqlDataSource>
    
    </div>
    </form>
</body>
</html>
