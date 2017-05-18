<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00111.aspx.vb" Inherits="M_Source_00_MOA00111" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>批核片語管理</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>  
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label6" runat="server" CssClass="toptitle" Text="批核片語管理" Width="100%"></asp:Label>
            </td></tr>
        </table>         
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap class="form" style="height: 25px; width: 10%;">
                    <asp:Label ID="Label1" runat="server" CssClass="form" Text="批核片語:" Width="77px"></asp:Label></td>
			    <td noWrap width="70%" class="form" style="height: 25px">			    
                        <asp:TextBox ID="comment" runat="server" Width="527px" MaxLength="255"></asp:TextBox></td>
                <td noWrap width="20%" class="form" style="height: 25px">
			    <asp:ImageButton ID="ImgBtn1" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" /></td>
		    </tr>
	    </table>
            
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
            AutoGenerateColumns="False" CellPadding="4" DataKeyNames="Phrase_num" DataSourceID="SqlDataSource1"
            ForeColor="#333333" GridLines="None" HorizontalAlign="Left" Width="100%">
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <Columns>
                <asp:BoundField DataField="Phrase_num" HeaderText="Phrase_num" InsertVisible="False"
                    ReadOnly="True" SortExpression="Phrase_num" Visible="False" />
                <asp:TemplateField HeaderText="批核片語" SortExpression="comment">
                    <EditItemTemplate>
                        <asp:TextBox ID="comment" runat="server" Text='<%# Bind("comment") %>' Width="550px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="comment"
                            ErrorMessage="請輸入批核片語" Width="119px"></asp:RequiredFieldValidator>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("comment") %>'></asp:Label>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Left" Width="80%" />
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
                    <ItemStyle HorizontalAlign="Center" Width="80%" />
                </asp:TemplateField>
            </Columns>
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
        </asp:GridView>
        &nbsp;
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [PHRASE] WHERE [Phrase_num] = @Phrase_num" InsertCommand="INSERT INTO PHRASE(employee_id, comment) VALUES (@employee_id, @comment)"
            SelectCommand="SELECT Phrase_num, comment FROM PHRASE WHERE [employee_id] = @employee_id ORDER BY Phrase_num"
            UpdateCommand="UPDATE [PHRASE] SET [comment] = @comment WHERE [Phrase_num] = @Phrase_num">
            <DeleteParameters>
                <asp:Parameter Name="Phrase_num" Type="Decimal" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="comment" Type="String" />
                <asp:Parameter Name="Phrase_num" Type="Decimal" />
            </UpdateParameters>
            <InsertParameters>
                <asp:SessionParameter Name="employee_id" SessionField="user_id" Type="String" />
                <asp:Parameter Name="comment" Type="String" />
            </InsertParameters>
            <SelectParameters>
                <asp:SessionParameter Name="employee_id" SessionField="user_id" />
            </SelectParameters>
        </asp:SqlDataSource>   
        
    </div>
    </form>
</body>
</html>
