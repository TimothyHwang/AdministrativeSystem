<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00105.aspx.vb" Inherits="Source_00_MOA00105" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>檔案上傳</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:FileUpload ID="UpFile" runat="server" Width="388px" />
        <asp:Button ID="btnUpload" runat="server" Text="上傳" />
        <asp:Label ID="ErrLab" runat="server" Width="300px" ForeColor="Red"></asp:Label>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
            AutoGenerateColumns="False" CellPadding="4" DataKeyNames="Upload_id" DataSourceID="SqlDataSource1"
            ForeColor="#333333" GridLines="None" Width="686px">
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <Columns>
                <asp:BoundField DataField="Upload_id" HeaderText="Upload_id" InsertVisible="False"
                    ReadOnly="True" SortExpression="Upload_id" Visible="False" />
                <asp:BoundField DataField="eformsn" HeaderText="eformsn" SortExpression="eformsn"
                    Visible="False" />
                <asp:BoundField DataField="FileName" HeaderText="檔案名稱" SortExpression="FileName">
                    <HeaderStyle HorizontalAlign="Center" Width="50%" />
                </asp:BoundField>
                <asp:BoundField DataField="FilePath" HeaderText="FilePath" SortExpression="FilePath"
                    Visible="False" />
                <asp:BoundField DataField="Upload_Time" HeaderText="上傳時間" SortExpression="Upload_Time">
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="30%" />
                </asp:BoundField>
                <asp:TemplateField SortExpression="FileName">
                    <ItemTemplate>
                        <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl='<%# "/MOA/M_Source/99/" & Eval("FileName") %>'
                            Text='檔案下載'></asp:HyperLink>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:ImageButton ID="ImageButton1" runat="server" CausesValidation="False" CommandName="Delete"
                            OnClientClick="return confirm('確定刪除嗎?')" ImageUrl="~/Image/delete.gif"  ToolTip="刪除" />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EmptyDataTemplate>
                <asp:Label ID="Label1" runat="server" Text="無任何檔案上傳" Width="197px"></asp:Label>
            </EmptyDataTemplate>
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [UPLOAD] WHERE [Upload_id] = @Upload_id" InsertCommand="INSERT INTO [UPLOAD] ([eformsn], [FileName], [FilePath], [Upload_Time]) VALUES (@eformsn, @FileName, @FilePath, @Upload_Time)"
            SelectCommand="SELECT [Upload_id], [eformsn], [FileName], [FilePath], [Upload_Time] FROM [UPLOAD]"
            UpdateCommand="UPDATE [UPLOAD] SET [eformsn] = @eformsn, [FileName] = @FileName, [FilePath] = @FilePath, [Upload_Time] = @Upload_Time WHERE [Upload_id] = @Upload_id">
            <DeleteParameters>
                <asp:Parameter Name="Upload_id" Type="Decimal" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="eformsn" Type="String" />
                <asp:Parameter Name="FileName" Type="String" />
                <asp:Parameter Name="FilePath" Type="String" />
                <asp:Parameter Name="Upload_Time" Type="DateTime" />
                <asp:Parameter Name="Upload_id" Type="Decimal" />
            </UpdateParameters>
            <InsertParameters>
                <asp:Parameter Name="eformsn" Type="String" />
                <asp:Parameter Name="FileName" Type="String" />
                <asp:Parameter Name="FilePath" Type="String" />
                <asp:Parameter Name="Upload_Time" Type="DateTime" />
            </InsertParameters>
        </asp:SqlDataSource>
    </div>
    </form>
</body>
</html>
