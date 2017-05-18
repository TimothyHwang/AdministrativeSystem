<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00032.aspx.vb" Inherits="OA_System_OrgRight" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>組織人員</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
        <asp:ImageButton ID="btnIns" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" />
        <asp:ImageButton ID="btnUpd" runat="server" ImageUrl="~/Image/update.gif" ToolTip="修改" />
        <asp:ImageButton ID="btnDel" runat="server" ImageUrl="~/Image/delete.gif" OnClientClick="return confirm('確定刪除單位嗎?')" ToolTip="刪除" />
            <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
            AutoGenerateColumns="False" DataKeyNames="empuid" DataSourceID="SqlDataSource1"
            EmptyDataText="查無任何資料" CellPadding="4" ForeColor="#333333" Width="100%" GridLines="None" Height="50px">
            <Columns>
                <asp:BoundField DataField="employee_id" HeaderText="帳號" SortExpression="employee_id" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="empemail" HeaderText="E-Mail" SortExpression="empemail" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="emp_chinese_name" HeaderText="姓名" SortExpression="emp_chinese_name" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="上一級主管">
                    <ItemTemplate>
                        <asp:GridView ID="GridView2" runat="server" DataSourceID="SqlDataSource2" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" ShowHeader="False" Width="142px">
                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            <Columns>
                                <asp:BoundField DataField="emp_chinese_name" ShowHeader="False" SortExpression="emp_chinese_name" />
                            </Columns>
                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <EditRowStyle BackColor="#999999" />
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        </asp:GridView>
                        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>" SelectCommand="select emp_chinese_name from employee where org_uid in (select PARENT_ORG_UID from ADMINGROUP where  admingroup.org_uid = @org_uid)">
                            <SelectParameters>
                                <asp:QueryStringParameter Name="org_uid" QueryStringField="org_uid" />
                            </SelectParameters>
                        </asp:SqlDataSource>
                    </ItemTemplate>
                    <HeaderStyle Width="200px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
            </Columns>
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <EditRowStyle BackColor="#999999" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server"
            DeleteCommand="DELETE FROM [EMPLOYEE] WHERE [empuid] = @empuid" InsertCommand="INSERT INTO [EMPLOYEE] ([employee_id], [member_uid], [ORG_UID], [empemail], [emp_chinese_name], [emp_english_name], [password], [create_date], [modify_date], [leave], [ErrCount], [title_id], [ArriveDate], [LeaveDate]) VALUES (@employee_id, @member_uid, @ORG_UID, @empemail, @emp_chinese_name, @emp_english_name, @password, @create_date, @modify_date, @leave, @ErrCount, @title_id, @ArriveDate, @LeaveDate)"
            SelectCommand="SELECT [empuid], [employee_id], [member_uid], [ORG_UID], [empemail], [emp_chinese_name], [emp_english_name], [password], [create_date], [modify_date], [leave], [ErrCount], [title_id], [ArriveDate], [LeaveDate] FROM [EMPLOYEE] where ORG_UID = @org_uid"
            UpdateCommand="UPDATE [EMPLOYEE] SET [employee_id] = @employee_id, [member_uid] = @member_uid, [ORG_UID] = @ORG_UID, [empemail] = @empemail, [emp_chinese_name] = @emp_chinese_name, [emp_english_name] = @emp_english_name, [password] = @password, [create_date] = @create_date, [modify_date] = @modify_date, [leave] = @leave, [ErrCount] = @ErrCount, [title_id] = @title_id, [ArriveDate] = @ArriveDate, [LeaveDate] = @LeaveDate WHERE [empuid] = @empuid" ConnectionString="<%$ ConnectionStrings:ConnectionString %>">
            <InsertParameters>
                <asp:Parameter Name="employee_id" Type="String" />
                <asp:Parameter Name="member_uid" Type="String" />
                <asp:Parameter Name="ORG_UID" Type="String" />
                <asp:Parameter Name="empemail" Type="String" />
                <asp:Parameter Name="emp_chinese_name" Type="String" />
                <asp:Parameter Name="emp_english_name" Type="String" />
                <asp:Parameter Name="password" Type="String" />
                <asp:Parameter Name="create_date" Type="DateTime" />
                <asp:Parameter Name="modify_date" Type="DateTime" />
                <asp:Parameter Name="leave" Type="String" />
                <asp:Parameter Name="ErrCount" Type="Int32" />
                <asp:Parameter Name="title_id" Type="String" />
                <asp:Parameter Name="ArriveDate" Type="DateTime" />
                <asp:Parameter Name="LeaveDate" Type="DateTime" />
            </InsertParameters>
            <UpdateParameters>
                <asp:Parameter Name="employee_id" Type="String" />
                <asp:Parameter Name="member_uid" Type="String" />
                <asp:Parameter Name="ORG_UID" Type="String" />
                <asp:Parameter Name="empemail" Type="String" />
                <asp:Parameter Name="emp_chinese_name" Type="String" />
                <asp:Parameter Name="emp_english_name" Type="String" />
                <asp:Parameter Name="password" Type="String" />
                <asp:Parameter Name="create_date" Type="DateTime" />
                <asp:Parameter Name="modify_date" Type="DateTime" />
                <asp:Parameter Name="leave" Type="String" />
                <asp:Parameter Name="ErrCount" Type="Int32" />
                <asp:Parameter Name="title_id" Type="String" />
                <asp:Parameter Name="ArriveDate" Type="DateTime" />
                <asp:Parameter Name="LeaveDate" Type="DateTime" />
                <asp:Parameter Name="empuid" Type="Decimal" />
            </UpdateParameters>
            <DeleteParameters>
                <asp:Parameter Name="empuid" Type="Decimal" />
            </DeleteParameters>
            <SelectParameters>
                <asp:QueryStringParameter Name="org_uid" QueryStringField="org_uid" />
            </SelectParameters>
        </asp:SqlDataSource>
    </form>
</body>
</html>
