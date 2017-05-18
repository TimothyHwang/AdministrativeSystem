<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA08006.aspx.vb" Inherits="M_Source_08_MOA08006" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>影印設備與單位管理</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <script type="text/JavaScript">
          function Check() {
              var re = /</;
              var d = document.form1;
              if (re.test(document.getElementById("DetailsView1_tb_Printer_No").value)) {
                  alert("您輸入的機器號碼有包含危險字元！");
                  document.getElementById("DetailsView1_tb_Printer_No").value = "";
                  return false;
              }

              if (re.test(document.getElementById("DetailsView1_tb_memo").value)) {
                  alert("您輸入的簡述有包含危險字元！");
                  document.getElementById("DetailsView1_tb_memo").value = "";
                  return false;
              }

              return true;
          }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="影印設備與單位管理(設備管理者設定)" Width="100%"></asp:Label>
            </td></tr>
        </table>
        <asp:DetailsView ID="DetailsView1" runat="server" AutoGenerateRows="False" 
        DataKeyNames="Guid_ID" DataSourceID="SqlDataSource1" DefaultMode="Insert" Width="100%" CssClass="form">
            <Fields>
                <asp:TemplateField ShowHeader="False">
                    <InsertItemTemplate>                    
                        <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
                            <tr>
                                <td noWrap style="height: 25px; width: 5%;">
                                <asp:Label ID="Lab1" runat="server" Text="部門名稱：" CssClass="form"></asp:Label>                    
                            </td>
			                <td noWrap style="height: 25px; width: 20%;">
                                <asp:DropDownList id="OrgSel"
                                    DataSourceID="SqlDataSource_ADMINGROUP"
                                    DataValueField="ORG_UID"
                                    DataTextField="ORG_NAME"
                                    runat="server" AutoPostBack="True">
                                </asp:DropDownList></td>
			                <td noWrap style="height: 25px; width: 3%;">
                                <asp:Label ID="Label1" runat="server" Text="姓名：" CssClass="form"></asp:Label>
                            </td>
			                <td align="left" noWrap style="height: 25px; width: 10%;" >
                                <asp:DropDownList ID="UserSel" runat="server" DataSourceID="SqlDataSource_EMPLOYEE" 
                                DataTextField="emp_chinese_name" DataValueField="employee_id" >
                                </asp:DropDownList>
                            </td>
                                <td noWrap width="5%" class="form">
                                    <asp:Label ID="Lab2" runat="server" Text="機器號碼：" />
                                </td>
                                <td noWrap width="5%" class="form">
                                    <asp:TextBox ID="tb_Printer_No" runat="server" Text='<%# Bind("Printer_No") %>'  MaxLength="6" Width="50px" />
                                </td>
                      </tr>
                       <tr>
                                 <td noWrap width="10%" class="form">
                                    <asp:Label ID="Label3" runat="server" Text="名稱：" />
                                </td>
                                <td noWrap width="20%" class="form">
                                    <asp:TextBox ID="tb_Printer_Name" runat="server" Text='<%# Bind("Printer_Name") %>'  MaxLength="50" Width="150px" />
                                </td>
                                <td noWrap width="3%" class="form">
                                    <asp:Label ID="Label2" runat="server" Text="簡述：" />
                                </td>
                                <td noWrap width="25%" class="form">
                                    <asp:TextBox ID="tb_memo" runat="server" Text='<%# Bind("memo") %>'  MaxLength="100" Width="200px" />
                                </td>
                                <td colspan ="2" noWrap width="10%" class="form" align = "right" >
                                    <asp:ImageButton ID="ImgInsert" runat="server" ImageUrl="~/Image/add.gif" OnClick="ImgInsert_Click"
                                        ToolTip="新增" OnClientClick = "return Check();" />
                                </td>
                            </tr>
                        </table> 

                    </InsertItemTemplate>
                </asp:TemplateField>
                <asp:CommandField />
            </Fields>
        </asp:DetailsView>
        <center><asp:Label ID="ErrMsg" runat="server" ForeColor='Red' /></center><br />
        <asp:SqlDataSource ID="SqlDataSource_ADMINGROUP" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [ORG_UID], [ORG_NAME] FROM [ADMINGROUP] ORDER BY [ORG_NAME]">
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource_EMPLOYEE" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [employee_id], [emp_chinese_name], [ORG_UID] FROM [EMPLOYEE] WHERE ([ORG_UID] = @ORG_UID) ORDER BY [emp_chinese_name]">
            <SelectParameters>
                <asp:ControlParameter ControlID="DetailsView1$OrgSel" Name="ORG_UID" PropertyName="SelectedValue" Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>  
        <asp:Label ID="lb_ORG_UID" runat="server" Visible = "false" />
        <asp:GridView ID="GV_Printer" runat="server" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" 
        DataKeyNames="Guid_ID" DataSourceID="SqlDataSource1" EmptyDataText="查無任何資料" Width="100%" CellPadding="4" 
        ForeColor="#333333" GridLines="None">
            <Columns>
            <asp:templatefield HeaderText="流水號" Visible="false" >
                   <ItemTemplate>
                     <asp:Label runat="server" id="lb_Guid_ID" Text='<%# Bind("Guid_ID") %>' ></asp:Label>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="5%" />
              </asp:templatefield> 
              <asp:templatefield HeaderText="部門名稱" >
                   <EditItemTemplate>
                        <asp:DropDownList ID="ddl_ORG_UID" runat="server" AppendDataBoundItems="true" 
                        DataSourceID = "SqlDataSource_ADMINGROUP" DataKeyNames="ORG_UID" DataValueField="ORG_UID" 
                        DataTextField ="ORG_NAME" AutoPostBack="True" 
                            SelectedValue='<%# Eval("ORG_UID") %>' 
                            onselectedindexchanged="ddl_ORG_UID_SelectedIndexChanged" />
                   </EditItemTemplate>
                <ItemTemplate>
                    <%# ShowORG_Name(Eval("ORG_UID"))%>
                </ItemTemplate>
                <HeaderStyle HorizontalAlign="Center" Width="15%" />
                <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Wrap="False" />
            </asp:templatefield> 
                 <asp:TemplateField ShowHeader="False" HeaderText="管理人員" >
                   <EditItemTemplate>
                        <asp:DropDownList ID="ddl_Employee" runat="server" AppendDataBoundItems="true" />
                    </EditItemTemplate>
                    <ItemTemplate>
                       <asp:DropDownList ID="ddl_Employee" runat="server" AppendDataBoundItems="true" Visible="false" />
                      <%# Showemp_chinese_name(Eval("employee_id"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                </asp:TemplateField>
                 <asp:TemplateField ShowHeader="False" visible = "false" >
                    <ItemTemplate>
                      <asp:Label ID="lb_ORG_UID" runat="server" text = '<%# Eval("ORG_UID") %>' />
                      <asp:Label ID="lb_employee_id" runat="server" text = '<%# Eval("employee_id") %>' />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="機器號碼">
                    <EditItemTemplate>
                        <asp:TextBox ID="tb_Printer_No" runat="server" Text='<%# Bind("Printer_No") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lb_Printer_No" runat="server" Text='<%# Bind("Printer_No") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="影印機名稱">
                    <EditItemTemplate>
                        <asp:TextBox ID="tb_PrinterName" runat="server" Text='<%# Bind("Printer_Name") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lb_PrinterName" runat="server" Text='<%# Bind("Printer_Name") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="簡述">
                    <EditItemTemplate>
                        <asp:TextBox ID="tb_memo" runat="server" Text='<%# Bind("memo") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lb_memo" runat="server" Text='<%# Bind("memo") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="left" Width="20%" />
                    <ItemStyle HorizontalAlign="left" />
                </asp:TemplateField>
                <asp:TemplateField ShowHeader="False">
                    <EditItemTemplate>
                        <asp:ImageButton ID="btnImgOK" runat="server" CommandName="Update" ImageUrl="~/Image/apply.gif"
                            ToolTip="確認" />&nbsp;<asp:ImageButton ID="btnImgCancel" runat="server" CommandName="Cancel"
                                ImageUrl="~/Image/cancel.gif" ToolTip="取消" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:ImageButton ID="ImgUpd" runat="server" CommandName="Edit" ImageUrl="~/Image/update.gif"
                            ToolTip="修改" />&nbsp;<asp:ImageButton ID="ImgDel" runat="server" CommandName="Delete"
                                ImageUrl="~/Image/delete.gif" OnClientClick="return confirm('確定刪除嗎?')" ToolTip="刪除" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
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
            DeleteCommand="DELETE FROM [P_0803] WHERE [Guid_ID] = @Guid_ID" 
            InsertCommand="INSERT INTO [P_0803] ([Printer_No],[Printer_Name],[employee_id],[ORG_UID],[Memo]) VALUES (@Printer_No,@Printer_Name,@employee_id,@ORG_UID,@Memo)"
            SelectCommand="SELECT [Guid_ID],[Printer_No],[Printer_Name],[employee_id],[ORG_UID],[Memo] from P_0803 ORDER BY [Guid_ID]"
            UpdateCommand="update P_0803 set ORG_UID=@ORG_UID,employee_id=@employee_id,Printer_No=@Printer_No,Printer_Name=@Printer_Name,Memo=@Memo where Guid_ID=@Guid_ID"
            OnInserted ="On_Inserted">
            <DeleteParameters>
                <asp:Parameter Name="Guid_ID" Type="Int32" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="Guid_ID" Type="Int32" />
                <asp:Parameter Name="ORG_UID" Type="String" />
                <asp:Parameter Name="employee_id" Type="String" />
                <asp:Parameter Name="Printer_No" Type="String" />
                <asp:Parameter Name="Printer_Name" Type="String" />
                <asp:Parameter Name="Memo" Type="String" />
            </UpdateParameters>
            <InsertParameters>
                <asp:ControlParameter ControlID="DetailsView1$OrgSel" Name="ORG_UID" PropertyName="SelectedValue" />
                <asp:ControlParameter ControlID="DetailsView1$UserSel" Name="employee_id" PropertyName="SelectedValue" />
                <asp:Parameter Name="Printer_No" Type="String" />
                <asp:Parameter Name="Printer_Name" Type="String" />
                <asp:Parameter Name="Memo" Type="String" />
            </InsertParameters>
        </asp:SqlDataSource>       
        
    </form>
</body>
</html>
