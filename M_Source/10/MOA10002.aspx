<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA10002.aspx.vb" Inherits="M_Source_10_MOA10002" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>新增主官</title>
    <link href='<%#ResolveUrl("~/css/site.css") %>' rel="stylesheet" type="text/css" />
    <link href='<%#ResolveUrl("~/Styles.css") %>' rel="stylesheet" type="text/css" />
    <script src='<%#ResolveUrl("~/Script/jquery-1.10.2.min.js") %>' type="text/javascript"></script>
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">       
        <div id="divContent" class="cContent">           
            <asp:DetailsView ID="dvManager" runat="server" AutoGenerateRows="False" 
               BackColor="#CCCCCC" BorderColor="#999999" BorderStyle="Solid" 
                BorderWidth="3px" CellPadding="4" CellSpacing="2" EmptyDataText="no data show" 
               EnableModelValidation="True" ForeColor="Black" Width="1024px" 
               DataSourceID="SqlDataSourceManagerEdit" Caption="新增主官" CaptionAlign="Top" 
                CssClass="tableView">
               <EditRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
               <Fields>
                   <asp:TemplateField HeaderText="職銜" SortExpression="Title" Visible="False">
                       <InsertItemTemplate>
                           <asp:TextBox ID="txtTitle" runat="server" Text='<%# Bind("Title") %>'></asp:TextBox>
                       </InsertItemTemplate>
                       <EditItemTemplate>
                           <asp:TextBox ID="txtTitle" runat="server" Text='<%# Bind("Title") %>'></asp:TextBox>
                       </EditItemTemplate>
                       <ItemTemplate>
                           <asp:Label ID="lblTitle" runat="server" Text='<%# Bind("Title") %>'></asp:Label>
                       </ItemTemplate>
                       <HeaderStyle Width="400px" />
                   </asp:TemplateField>
                   <asp:TemplateField HeaderText="名稱" SortExpression="Name">
                       <InsertItemTemplate>
                           <asp:TextBox ID="txtName" runat="server" Text='<%# Bind("Name") %>'></asp:TextBox>
                           <asp:Label ID="Label1" runat="server" ForeColor="Red" 
                               Text="*主官名稱請少於2行16個字以內，以免影響頁面排版"></asp:Label>
                       </InsertItemTemplate>
                       <EditItemTemplate>
                           <asp:TextBox ID="txtName" runat="server" Text='<%# Bind("Name") %>'></asp:TextBox>
                       </EditItemTemplate>
                       <ItemTemplate>
                           <asp:Label ID="lblName" runat="server" Text='<%# Bind("Name") %>'></asp:Label>
                       </ItemTemplate>
                   </asp:TemplateField>
                   <asp:TemplateField>
                       <InsertItemTemplate>
                           <asp:ImageButton ID="ibnAddOK" runat="server" ImageUrl="~/Image/apply.gif" 
                               onclick="ibnAddOK_Click" />
                           <asp:ImageButton ID="ibnCancel" runat="server" ImageUrl="~/Image/cancel.gif" 
                               onclick="ibnCancel_Click"/>
                       </InsertItemTemplate>
                   </asp:TemplateField>
               </Fields>
               <FooterStyle BackColor="#CCCCCC" />
               <HeaderStyle CssClass="toptitle" HorizontalAlign="Center" VerticalAlign="Middle" />
               <PagerStyle BackColor="#CCCCCC" ForeColor="Black" HorizontalAlign="Left" />
               <RowStyle BackColor="White" />
           </asp:DetailsView>
           <asp:SqlDataSource ID="SqlDataSourceManagerEdit" runat="server" 
               ConflictDetection="CompareAllValues" 
               ConnectionString="<%$ ConnectionStrings:ConnectionString %>" 
               InsertCommand="INSERT INTO [P_10] ([Title], [Name], [RankId], [Status], [CreateDate], [Creator]) VALUES (@Title, @Name, @RankId, @Status, @CreateDate, @Creator)" 
               OldValuesParameterFormatString="original_{0}" 
               SelectCommand="SELECT * FROM [P_10] ">
               <InsertParameters>
                   <asp:Parameter Name="Title" Type="String" />
                   <asp:Parameter Name="Name" Type="String" />
                   <asp:Parameter Name="RankId" Type="Int32" />
                   <asp:Parameter Name="Status" Type="Byte" />
                   <asp:Parameter Name="CreateDate" Type="DateTime" />
                   <asp:Parameter Name="Creator" Type="String" />
               </InsertParameters>
           </asp:SqlDataSource>
       </div>
    </form>
</body>
</html>
