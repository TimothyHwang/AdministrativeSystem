<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04101.aspx.vb" Inherits="M_Source_04_MOA04101" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>領料申請管理</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <base target='_self'>
    <script language="javascript" type="text/javascript" >
        function applyCheck(obj, stock, applied) {
            if (trim(obj.value) != "") {
                if (isNaN(trim(obj.value))) {
                    alert("請輸入數字,或輸入空白為取消");
                    obj.focus();
                }
                else {
                    var i_stock = parseInt(stock);
                    var i_applied = parseInt(applied);
                    var i_inputobj = parseInt(obj.value);
                    if (i_inputobj == 0) {
                        alert("請輸入正確的領料申請數量,或輸入空白為取消");
                        obj.focus();
                    }
                    if (i_inputobj > 0 && i_inputobj > i_stock) {
                        alert("不能申請大於庫存量,或輸入空白為取消");
                        obj.focus();
                    }
                    if (i_inputobj < 0 && (i_inputobj * -1) > i_applied) {
                        alert("不能退大於已申請量,或輸入空白為取消");
                        obj.focus();
                    }
                }
            }
        }
        function trim(strvalue) {
            ptntrim = /(^\s*)|(\s*$)/g;
            return strvalue.replace(ptntrim, "");
        }
    </script>
</head>
<body background="../../Image/main_bg.jpg" >
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="領料申請" Width="100%"></asp:Label>
            </td></tr>
        </table>
                          
                        <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
                            <tr>
                                <td  width="10%" class="form" align="right">
                                    <asp:Label ID="Lab1" runat="server" Text="物料名稱："></asp:Label>
                                </td>
                                <td  width="25%" class="form">
                                    <asp:TextBox ID="It_Name" runat="server" MaxLength='15' Width="218px"></asp:TextBox>
                                </td>
                                
                                <td  width="65%" class="form">
                                    <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif"
                                        ToolTip="搜尋" />
                                </td>
                            </tr>
                        </table> 
<div style=" height:500px; position:relative ;overflow:auto; border:1px; border-color:Gray; border-style:solid">                  
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" AllowSorting="True"
            EmptyDataText="查無任何資料" Width="100%" CellPadding="4" ForeColor="#333333" 
            GridLines="None" DataSourceID="SqlDataSource1">
            <Columns>
                <asp:templatefield HeaderText="物料代碼">
                    <itemtemplate>
                        <asp:label id="lblIt_code"
                        text= '<%# Bind("it_code") %>'
                        runat="server"/>
                    </itemtemplate>
                    <ItemStyle HorizontalAlign="Left" Width="5%" />
                    <HeaderStyle HorizontalAlign="Left" />
                </asp:templatefield>
                <asp:templatefield HeaderText="物料名稱">
                    <itemtemplate>
                        <asp:label id="lblIt_name"
                        text= '<%# Bind("it_name") %>'
                        runat="server"/>
                    </itemtemplate>
                    <ItemStyle HorizontalAlign="Left" Wrap="true" Width="20%"  />
                    <HeaderStyle HorizontalAlign="Left" />
                </asp:templatefield>
                <asp:templatefield HeaderText="物料規格">
                    <itemtemplate>
                        <asp:label id="lblIt_spec"
                        text= '<%# Bind("it_spec") %>'
                        runat="server"/>
                    </itemtemplate>
                    <ItemStyle HorizontalAlign="Left"  Wrap="true" Width="5%" />
                    <HeaderStyle HorizontalAlign="Left" />
                </asp:templatefield>
                <asp:templatefield HeaderText="物料單位">
                    <itemtemplate>
                        <asp:label id="lblIt_unit"
                        text= '<%# Bind("it_unit") %>'
                        runat="server"/>
                    </itemtemplate>
                    <ItemStyle HorizontalAlign="Left" Width="6%" />
                    <HeaderStyle HorizontalAlign="Left" />
                </asp:templatefield>
                <asp:templatefield HeaderText="庫存量">
                    <itemtemplate>
                         <asp:label id="lblIt_stock"        
                        runat="server"/>
                    </itemtemplate>
                    <ItemStyle HorizontalAlign="Center" Width="5%" />
                </asp:templatefield>
                <asp:templatefield HeaderText="已申請量">
                    <itemtemplate>
                          <asp:label id="lblIt_applied"
                        runat="server"/>
                    </itemtemplate>
                    <ItemStyle HorizontalAlign="Center" Width="5%" />
                </asp:templatefield>
                <asp:templatefield HeaderText="申請量">
                    <itemtemplate>
                        <asp:TextBox ID="txtIt_apply" runat="server" width ="50px"></asp:TextBox>
                    </itemtemplate>
                    <ItemStyle HorizontalAlign="Center" Width="5%" />
                </asp:templatefield>
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
          </div>
          <div style="text-align :center; margin:5px">
          <asp:Button ID="Button1" runat="server" Text="申請確認" Height="32px"  />
</div>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
            ConnectionString="<%$ ConnectionStrings:ConnectionString %>" 
            SelectCommand="SELECT * FROM [P_0407] WHERE ([it_code] LIKE '' + @it_code + '%')">
            <SelectParameters>
                <asp:QueryStringParameter Name="it_code" QueryStringField="itcode" 
                    Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>      
        
          
    </form>
</body>
</html>
