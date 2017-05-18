<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00041.aspx.vb" Inherits="Source_00_MOA00041" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>人員新增</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">
    <div>
        <table  border="0" style="Z-INDEX: 101; LEFT: 13px; WIDTH: 656px; POSITION: absolute; TOP: 17px; HEIGHT: 354px">
            <tr>
                <td align="center">
                        <asp:Label ID="LabTitle" runat="server" CssClass="toptitle" Text="人員新增" Width="100%"></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:DetailsView ID="DetailsView1" runat="server" AutoGenerateRows="False" DataKeyNames="empuid"
                        DataSourceID="SqlDataSource1" Height="301px" Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
                        <Fields>
                            <asp:TemplateField HeaderText="empuid" InsertVisible="False" SortExpression="empuid" ShowHeader="False">
                                <EditItemTemplate>
                                    <asp:Label ID="LabEmpuidUpd" runat="server" Text='<%# Eval("empuid") %>' Visible="False"></asp:Label>
                                </EditItemTemplate>
                                <ItemTemplate>
                                    <asp:Label ID="LabEmpuid" runat="server" Text='<%# Bind("empuid") %>' Visible="False"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="人員帳號" SortExpression="employee_id">
                                <EditItemTemplate>
                                    <asp:TextBox ID="employee_id" runat="server" Text='<%# Bind("employee_id") %>' ReadOnly="True" MaxLength="10"></asp:TextBox>
                                </EditItemTemplate>
                                <InsertItemTemplate>
                                    <asp:TextBox ID="employee_id" runat="server" Text='<%# Bind("employee_id") %>' MaxLength="10"></asp:TextBox>
                                </InsertItemTemplate>
                                <ItemTemplate>
                                    <asp:Label ID="Label2" runat="server" Text='<%# Bind("employee_id") %>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle Wrap="True" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="member_uid" SortExpression="member_uid" ShowHeader="False">
                                <EditItemTemplate>
                                    <asp:TextBox ID="member_uid" runat="server" Text='<%# Bind("member_uid") %>'></asp:TextBox>
                                </EditItemTemplate>
                                <InsertItemTemplate>
                                    <asp:TextBox ID="member_uid" runat="server" Text='<%# Bind("member_uid") %>'></asp:TextBox>
                                </InsertItemTemplate>
                                <ItemTemplate>
                                    <asp:Label ID="Label7" runat="server" Text='<%# Bind("member_uid") %>' Visible="False"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="單位" SortExpression="ORG_UID">
                                <EditItemTemplate>
                                    <asp:DropDownList ID="DropDownList1" runat="server" DataSourceID="SqlDataSource4"
                                        DataTextField="ORG_NAME" DataValueField="ORG_UID" SelectedValue='<%# Bind("ORG_UID") %>'>
                                    </asp:DropDownList>
                                </EditItemTemplate>
                                <InsertItemTemplate>
                                    <asp:DropDownList ID="DropDownList1" runat="server" DataSourceID="SqlDataSource4"
                                        DataTextField="ORG_NAME" DataValueField="ORG_UID" SelectedValue='<%# Bind("ORG_UID") %>'>
                                    </asp:DropDownList>
                                </InsertItemTemplate>
                                <ItemTemplate>
                                    <asp:Label ID="Label3" runat="server" Text='<%# Bind("ORG_UID") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="E-Mail" SortExpression="empemail">
                                <EditItemTemplate>
                                    <asp:TextBox ID="TextBox3" runat="server" Text='<%# Bind("empemail") %>' MaxLength="100" Width="200px"></asp:TextBox>
                                </EditItemTemplate>
                                <InsertItemTemplate>
                                    <asp:TextBox ID="TextBox3" runat="server" Text='<%# Bind("empemail") %>' MaxLength="100" Width="200px"></asp:TextBox>
                                </InsertItemTemplate>
                                <ItemTemplate>
                                    <asp:Label ID="Label4" runat="server" Text='<%# Bind("empemail") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="人員名稱" SortExpression="emp_chinese_name">
                                <EditItemTemplate>
                                    <asp:TextBox ID="TextBox4" runat="server" Text='<%# Bind("emp_chinese_name") %>' MaxLength="10" Width="130px"></asp:TextBox>
                                </EditItemTemplate>
                                <InsertItemTemplate>
                                    <asp:TextBox ID="TextBox4" runat="server" Text='<%# Bind("emp_chinese_name") %>' MaxLength="10" Width="130px"></asp:TextBox>
                                </InsertItemTemplate>
                                <ItemTemplate>
                                    <asp:Label ID="Label5" runat="server" Text='<%# Bind("emp_chinese_name") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="在職" SortExpression="leave">
                                <EditItemTemplate>
                                    <asp:DropDownList ID="DDLLeave1" runat="server" SelectedValue='<%# Bind("leave") %>'>
                                        <asp:ListItem Value="Y" Selected="True">是</asp:ListItem>
                                        <asp:ListItem Value="N">否</asp:ListItem>
                                    </asp:DropDownList>
                                </EditItemTemplate>
                                <InsertItemTemplate>
                                    <asp:DropDownList ID="DDLLeave2" runat="server" SelectedValue='<%# Bind("leave") %>'>
                                        <asp:ListItem Value="Y">是</asp:ListItem>
                                        <asp:ListItem Value="N">否</asp:ListItem>
                                    </asp:DropDownList>
                                </InsertItemTemplate>
                                <ItemTemplate>
                                    <asp:Label ID="LabLeave" runat="server" Text='<%# Bind("leave") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="職稱" SortExpression="title_id">
                                <EditItemTemplate>
                                    <asp:DropDownList ID="DropDownList5" runat="server" DataSourceID="SqlDataSource2"
                                        DataTextField="Title_Name" DataValueField="Title_ID" SelectedValue='<%# Bind("title_id") %>'>
                                    </asp:DropDownList><asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                                        SelectCommand="SELECT [Title_ID], [Title_Name] FROM [TITLES] ORDER BY [Title_Name]">
                                    </asp:SqlDataSource>
                                </EditItemTemplate>
                                <InsertItemTemplate>
                                    <asp:DropDownList ID="DropDownList6" runat="server" DataSourceID="SqlDataSource2"
                                        DataTextField="Title_Name" DataValueField="Title_ID" SelectedValue='<%# Bind("title_id") %>'>
                                    </asp:DropDownList><asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                                        SelectCommand="SELECT [Title_ID], [Title_Name] FROM [TITLES] ORDER BY [Title_Name]">
                                    </asp:SqlDataSource>
                                </InsertItemTemplate>
                                <ItemTemplate>
                                    <asp:Label ID="Label8" runat="server" Text='<%# Bind("title_id") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="職級" SortExpression="TU_ID">
                                <EditItemTemplate><asp:DropDownList ID="DropDownList7" runat="server" DataSourceID="SqlDataSource3"
                                        DataTextField="TU_Name" DataValueField="TU_ID" SelectedValue='<%# Bind("TU_ID") %>'>
                                </asp:DropDownList><asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                                        SelectCommand="SELECT TU_ID, TU_Name FROM TITLES_U ORDER BY TU_Name"></asp:SqlDataSource>
                                </EditItemTemplate>
                                <InsertItemTemplate><asp:DropDownList ID="DropDownList8" runat="server" DataSourceID="SqlDataSource3"
                                        DataTextField="TU_Name" DataValueField="TU_ID" SelectedValue='<%# Bind("TU_ID") %>'>
                                </asp:DropDownList><asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                                        SelectCommand="SELECT TU_ID, TU_Name FROM TITLES_U ORDER BY TU_Name"></asp:SqlDataSource>
                                </InsertItemTemplate>
                                <ItemTemplate>
                                    <asp:Label ID="Label6" runat="server" Text='<%# Bind("TU_ID") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField ShowHeader="False">
                                <EditItemTemplate>
                                    <asp:ImageButton ID="btnImgUpd" runat="server" CommandName="Update" ImageUrl="~/Image/apply.gif"
                                        OnClientClick="return confirm('確定修改人員資料嗎?')" OnClick="btnImgUpd_Click" ToolTip="確認" />
                                    <asp:ImageButton ID="btnImgBack" runat="server" ImageUrl="~/Image/backtop.gif"
                                        ToolTip="回上頁" OnClick="btnImgBack_Click" />
                                </EditItemTemplate>
                                <InsertItemTemplate>
                                    <asp:ImageButton ID="btnImgIns" runat="server" CommandName="Insert" ImageUrl="~/Image/apply.gif"
                                    OnClick="btnImgIns_Click" ToolTip="確認" OnClientClick="return confirm('確定新增人員資料嗎?')" />
                                    <asp:ImageButton ID="btnImgBackIns" runat="server" ImageUrl="~/Image/backtop.gif" ToolTip="回上頁" OnClick="btnImgBackIns_Click" />
                                </InsertItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="password" ShowHeader="False" SortExpression="password">
                                <EditItemTemplate>
                                    <asp:TextBox ID="PW" runat="server" Text='<%# Bind("password") %>'></asp:TextBox>
                                </EditItemTemplate>
                                <InsertItemTemplate>
                                    <asp:TextBox ID="PW" runat="server" Text='<%# Bind("password") %>'></asp:TextBox>
                                </InsertItemTemplate>
                                <ItemTemplate>
                                    <asp:Label ID="Label1" runat="server" Text='<%# Bind("password") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Fields>
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <CommandRowStyle BackColor="#E2DED6" Font-Bold="True" />
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FieldHeaderStyle BackColor="#E9ECF1" Font-Bold="True" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <EditRowStyle BackColor="#E9ECF1" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    </asp:DetailsView>
                    <asp:Label ID="LabErr" runat="server" ForeColor="Red" Visible="False" Width="500px"></asp:Label><br />
                    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                        DeleteCommand="DELETE FROM [EMPLOYEE] WHERE [empuid] = @empuid" InsertCommand="INSERT INTO [EMPLOYEE] ([employee_id], [member_uid], [ORG_UID], [empemail], [emp_chinese_name], [emp_english_name], [password],  [leave], [title_id], [TU_ID], [ArriveDate], [LeaveDate]) VALUES (@employee_id, @member_uid, @ORG_UID, @empemail, @emp_chinese_name, @emp_english_name, @password, @leave, @title_id, @TU_ID, @ArriveDate, @LeaveDate)"
                        SelectCommand="SELECT * FROM [EMPLOYEE] WHERE ([empuid] = @empuid)"
                        UpdateCommand="UPDATE [EMPLOYEE] SET [employee_id] = @employee_id, [ORG_UID] = @ORG_UID, [empemail] = @empemail, [emp_chinese_name] = @emp_chinese_name, [emp_english_name] = @emp_english_name, [modify_date] = @modify_date, [leave] = @leave, [title_id] = @title_id, [TU_ID] = @TU_ID, [ArriveDate] = @ArriveDate, [LeaveDate] = @LeaveDate WHERE [empuid] = @empuid">
                        <DeleteParameters>
                            <asp:Parameter Name="empuid" Type="Decimal" />
                        </DeleteParameters>
                        <UpdateParameters>
                            <asp:Parameter Name="employee_id" Type="String" />
                            <asp:Parameter Name="ORG_UID" Type="String" />
                            <asp:Parameter Name="empemail" Type="String" />
                            <asp:Parameter Name="emp_chinese_name" Type="String" />
                            <asp:Parameter Name="emp_english_name" Type="String" />
                            <asp:Parameter Name="modify_date" Type="DateTime" />
                            <asp:Parameter Name="leave" Type="String" />
                            <asp:Parameter Name="title_id" Type="String" />
                            <asp:Parameter Name="TU_ID" Type="String" />
                            <asp:Parameter Name="ArriveDate" Type="DateTime" />
                            <asp:Parameter Name="LeaveDate" Type="DateTime" />
                            <asp:Parameter Name="empuid" Type="Decimal" />
                        </UpdateParameters>
                        <SelectParameters>
                            <asp:QueryStringParameter Name="empuid" QueryStringField="empuid" />
                        </SelectParameters>
                        <InsertParameters>
                            <asp:Parameter Name="employee_id" Type="String" />
                            <asp:Parameter Name="member_uid" Type="String" />
                            <asp:Parameter Name="ORG_UID" Type="String" />
                            <asp:Parameter Name="empemail" Type="String" />
                            <asp:Parameter Name="emp_chinese_name" Type="String" />
                            <asp:Parameter Name="emp_english_name" Type="String" />
                            <asp:Parameter Name="password" Type="String" />
                            <asp:Parameter Name="leave" Type="String" />
                            <asp:Parameter Name="title_id" Type="String" />
                            <asp:Parameter Name="TU_ID" Type="String" />
                            <asp:Parameter Name="ArriveDate" Type="DateTime" />
                            <asp:Parameter Name="LeaveDate" Type="DateTime" />
                        </InsertParameters>
                    </asp:SqlDataSource>
                <asp:SqlDataSource ID="SqlDataSource4" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                                        SelectCommand="SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] ORDER BY [ORG_NAME]">
                                    </asp:SqlDataSource>
                </td>
            </tr>
            
        </table>
    
    </div>
    </form>
</body>
</html>
