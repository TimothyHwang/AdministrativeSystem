<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00108.aspx.vb" Inherits="Source_00_MOA00108" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>單位行事曆內容</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="left" style="width:3%" >
                <a href='MOA00107.aspx?Date=<%= NoteDate.Text %>'>
                <img  align='absmiddle' title="回上頁" border="0" src="../../Image/backtop.gif" class="btn_image"/></a>
            </td><td align="center" style="width:97%" >
                    <asp:Label ID="Label3" runat="server" CssClass="toptitle" Text="單位行事曆內容" Width="100%"></asp:Label>
            </td></tr>
        </table>
 	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
            <tr><td noWrap align='right'>
                <asp:Label ID="Label1" runat="server" Text="日期：" CssClass="form"></asp:Label>                    
            </td><td>
                <asp:TextBox ID="NoteDate" runat="server" Width="60px" OnKeyDown="return false"></asp:TextBox>
            </td><td noWrap align='right'>
                <asp:Label ID="Label2" runat="server" Text="事項：" CssClass="form"></asp:Label>                    
            </td><td>
                <asp:TextBox ID="NoteContent" runat="server" Width="150px"></asp:TextBox>
            </td><td>
                <asp:ImageButton ID="ImgInsert" runat="server" ImageUrl="../../Image/add.gif" title="新增"/>
		    </td></tr>
        </table>
        <asp:Label ID="ErrMsg" runat="server" ForeColor='Red' Text="" ></asp:Label>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" DataKeyNames="Note_Num" DataSourceID="SqlDataSource1" Width="100%">
            <Columns>
                <asp:BoundField  DataField="NoteContent" HeaderText="事項" SortExpression="NoteContent">
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <EditItemTemplate>
                        <asp:Button ID="Button1" runat="server" CausesValidation="True" CommandName="Update"
                            Text="更新" />&nbsp;<asp:Button ID="Button2" runat="server" CausesValidation="False"
                                CommandName="Cancel" Text="取消" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Button ID="Button1" runat="server" CausesValidation="False" CommandName="Edit" Text="修改" />
                        <asp:Button ID="Button2" runat="server" CausesValidation="False" OnClientClick="return confirm('確定刪除嗎?')"
                                CommandName="Delete" Text="刪除" />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [P_1201] WHERE [Note_Num] = @Note_Num" 
            InsertCommand=""
            SelectCommand=""
            UpdateCommand="UPDATE [P_1201] SET [NoteContent] = @NoteContent WHERE [Note_Num] = @Note_Num"
            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>">
            <DeleteParameters>
                <asp:Parameter Name="Note_Num" Type="Int32" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="NoteContent" Type="String" />
                <asp:Parameter Name="Note_Num" Type="Int32" />
            </UpdateParameters>
        </asp:SqlDataSource>        
    </form>
</body>
</html>
