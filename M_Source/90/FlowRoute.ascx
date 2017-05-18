<%@ Control Language="VB" AutoEventWireup="false" CodeFile="FlowRoute.ascx.vb" Inherits="Source_90_FlowRoute" %>
<asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" CellPadding="4"
    DataSourceID="SqlDataSource1" ForeColor="#333333" GridLines="None" Width="95%">
    <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
    <RowStyle BackColor="#EFF3FB" />
    <Columns>
        <asp:BoundField DataField="emp_chinese_name" HeaderText="人員" SortExpression="emp_chinese_name">
            <ItemStyle HorizontalAlign="Center" />
            <HeaderStyle HorizontalAlign="Center" Width="10%" />
        </asp:BoundField>
        <asp:BoundField DataField="group_name" HeaderText="批核群組" SortExpression="group_name">
            <ItemStyle HorizontalAlign="Center" />
            <HeaderStyle HorizontalAlign="Center" Width="20%" />
        </asp:BoundField>
        <asp:BoundField DataField="hddate" HeaderText="批核日期" SortExpression="hddate">
            <ItemStyle HorizontalAlign="Center" />
            <HeaderStyle HorizontalAlign="Center" Width="30%" />
        </asp:BoundField>
        <asp:BoundField DataField="comment" HeaderText="批核意見" SortExpression="comment">
            <HeaderStyle HorizontalAlign="Center" Width="30%" />
            <ItemStyle HorizontalAlign="Center" />
        </asp:BoundField>
        <asp:TemplateField HeaderText="狀態" SortExpression="gonogo">
            <EditItemTemplate>
                <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("gonogo") %>'></asp:TextBox>
            </EditItemTemplate>
            <ItemTemplate>
                <asp:Label ID="lb_eformid" runat="server" Text='<%# Bind("eformid") %>' Visible="false" />
                <asp:Label ID="Label1" runat="server" Text='<%# FunStatus("gonogo") %>'></asp:Label>
            </ItemTemplate>
            <ItemStyle HorizontalAlign="Center" />
            <HeaderStyle HorizontalAlign="Center" Width="10%" />
        </asp:TemplateField>
    </Columns>
    <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
    <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
    <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
    <EditRowStyle BackColor="#2461BF" />
    <AlternatingRowStyle BackColor="White" />
</asp:GridView>
<asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
    SelectCommand="SELECT [emp_chinese_name], [group_name],comment, [hddate], [gonogo],eformid FROM [flowctl] WHERE ([eformsn] = @eformsn) ORDER BY [eformsn], [flowsn]">
    <SelectParameters>
        <asp:QueryStringParameter Name="eformsn" QueryStringField="eformsn" Type="String" />
    </SelectParameters>
</asp:SqlDataSource>
