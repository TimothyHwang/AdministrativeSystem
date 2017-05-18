<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00210.aspx.vb" Inherits="M_Source_00_MOA00210" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>公告管理</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <link href="../../css/jquery.datepick.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.10.2.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick-zh-TW.js" type="text/javascript"></script>
    <script type="text/javascript" language="javascript">
        $(function () {
            $("#SDate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#SDate").val(), showTrigger: '#calImg' });
            $("#EDate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#EDate").val(), showTrigger: '#calImg' });            
        });
        function selectdate(obj) {
            //alert($("#" + obj.id).val());     
            $("#" + obj.id).datepick({ formats: 'yyyy/m/d', defaultDate: $("#" + obj.id).val(), showTrigger: '#calImg' });
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr>
                <td align="center">
                    <asp:Label ID="Label6" runat="server" CssClass="toptitle" Text="公告管理" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>       
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap width="10%" class="form">
                    <asp:Label ID="Label2" runat="server" CssClass="form" Text="標題："></asp:Label></td>
			    <td noWrap width="25%" class="form" style="height: 25px"><asp:TextBox ID="T_Content" runat="server" Width="150px" MaxLength="100"></asp:TextBox></td>
			    
			    <td noWrap width="10%" style="height: 25px">			    
			        <asp:Label ID="Label3" runat="server" Text="日期：" CssClass="form"></asp:Label></td>
			    <td noWrap width="35%" style="height: 25px">
                    <asp:TextBox ID="SDate" runat="server" Width="80px" OnKeyDown="return false"></asp:TextBox>
                    <div style="display: none;">
                        <img id="calImg" src="../../Image/calendar.gif" alt="選擇日期" />
                    </div>
                    <asp:Label ID="Lab3" runat="server" Text="~" CssClass="form" ></asp:Label>
                    <asp:TextBox ID="EDate" runat="server" Width="80px" OnKeyDown="return false"></asp:TextBox>
                </td>
			    <td noWrap width="10%" class="form" style="height: 25px">
                <asp:ImageButton ID="ImgBtn2" runat="server" ImageUrl="~/Image/Search.gif" ToolTip="查詢" /></td>
                
			    <td noWrap width="10%" class="form" style="height: 25px">
			    <asp:ImageButton ID="ImgBtn1" runat="server" ImageUrl="~/Image/add.gif" ToolTip="新增" /></td>
		    </tr>
	    </table>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
            AutoGenerateColumns="False" CellPadding="4" DataKeyNames="ANNID_i" DataSourceID="SqlDataSource1"
            ForeColor="#333333" GridLines="None" HorizontalAlign="Left" Width="100%">
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <Columns>
                <asp:BoundField DataField="DATE_d" HeaderText="登錄日期" SortExpression="DATE_d" ReadOnly="True" >
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="公告單位" SortExpression="uicode">
                    <EditItemTemplate>
                        <asp:DropDownList ID="DropDownList1" runat="server" SelectedValue='<%# Bind("uicode") %>'>
                            <asp:ListItem Value="0" Selected="True">單位內部</asp:ListItem>
                            <asp:ListItem Value="1">國防部全體</asp:ListItem>
                            <asp:ListItem Value="2">系統公告</asp:ListItem>
                        </asp:DropDownList>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# FunShare("uicode") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="標題" SortExpression="TITLE_nvc">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox2" runat="server" Text='<%# Bind("TITLE_nvc") %>' MaxLength="100"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label3" runat="server" Text='<%# Bind("TITLE_nvc") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="內容" SortExpression="CONTENT_nvc">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox3" runat="server" Text='<%# Bind("CONTENT_nvc") %>' Rows="5" TextMode="MultiLine" Width="200px" MaxLength="1000"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label4" runat="server" Text='<%# Bind("CONTENT_nvc") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="25%" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="姓名" SortExpression="ANN_Name">
                    <EditItemTemplate>                        
                        <asp:Label ID="TextBox4" runat="server" Text='<%# Bind("ANN_Name") %>'></asp:Label>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label5" runat="server" Text='<%# Bind("ANN_Name") %>'></asp:Label>                        
                       <asp:HiddenField ID="HiddenField1" runat="server" Value='<%# Bind("employeeid") %>' />                       
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="電話" SortExpression="ANN_Phone">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox5" runat="server" Text='<%# Bind("ANN_Phone") %>' Width="100px" MaxLength="20"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label6" runat="server" Text='<%# Bind("ANN_Phone") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="截止日期" SortExpression="DATE_e">
                    <EditItemTemplate>
                        <asp:TextBox ID="OverDate" OnKeyDown="return false" runat="server" Text='<%# Bind("DATE_e") %>' Width="70px" onclick="javascript:selectdate(this)"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("DATE_e") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField ShowHeader="False">
                    <EditItemTemplate>
                        <asp:ImageButton ID="ImageButton1" runat="server" CausesValidation="True" CommandName="Update"
                            ImageUrl="~/Image/apply.gif" ToolTip="確認" />&nbsp;<asp:ImageButton ID="ImageButton2"
                                runat="server" CausesValidation="False" CommandName="Cancel" ImageUrl="~/Image/cancel.gif"
                                ToolTip="取消" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:ImageButton ID="ImageButton1" runat="server" CausesValidation="False" CommandName="Edit"
                            ImageUrl="~/Image/update.gif" ToolTip="修改" />&nbsp;<br />
                        <asp:ImageButton ID="ImageButton2"
                                runat="server" CausesValidation="False" CommandName="Delete" ImageUrl="~/Image/delete.gif"
                                OnClientClick="return confirm('確定刪除嗎?')" ToolTip="刪除" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="5%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
            </Columns>
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
        </asp:GridView>
        
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [T_ANNOUNTCEMENT] WHERE [ANNID_i] = @ANNID_i" InsertCommand="INSERT INTO [T_ANNOUNTCEMENT] ([DATE_d], [DEPT_i], [TITLE_nvc], [CONTENT_nvc], [uicode], [DATE_e], [ANN_Name], [ANN_Phone]) VALUES (@DATE_d, @DEPT_i, @TITLE_nvc, @CONTENT_nvc, @uicode, @DATE_e, @ANN_Name, @ANN_Phone)"
            SelectCommand=""
            UpdateCommand="UPDATE [T_ANNOUNTCEMENT] SET [TITLE_nvc] = @TITLE_nvc, [CONTENT_nvc] = @CONTENT_nvc, [uicode] = @uicode, [DATE_e] = @DATE_e, [ANN_Name] = @ANN_Name, [ANN_Phone] = @ANN_Phone WHERE [ANNID_i] = @ANNID_i">
            <DeleteParameters>
                <asp:Parameter Name="ANNID_i" Type="Int32" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="TITLE_nvc" Type="String" />
                <asp:Parameter Name="CONTENT_nvc" Type="String" />
                <asp:Parameter Name="uicode" Type="String" />
                <asp:Parameter DbType="DateTime" Name="DATE_e" />
                <asp:Parameter Name="ANN_Name" Type="String" />
                <asp:Parameter Name="ANN_Phone" Type="String" />
                <asp:Parameter Name="ANNID_i" Type="Int32" />
            </UpdateParameters>
            <InsertParameters>
                <asp:Parameter DbType="Date" Name="DATE_d" />
                <asp:Parameter Name="DEPT_i" Type="String" />
                <asp:Parameter Name="TITLE_nvc" Type="String" />
                <asp:Parameter Name="CONTENT_nvc" Type="String" />
                <asp:Parameter Name="uicode" Type="String" />
                <asp:Parameter DbType="Date" Name="DATE_e" />
                <asp:Parameter Name="ANN_Name" Type="String" />
                <asp:Parameter Name="ANN_Phone" Type="String" />
            </InsertParameters>
        </asp:SqlDataSource>                                                                                                    
    </form>
</body>
</html>
