<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA07001.aspx.vb" Inherits="Source_07_MOA07001" %>

<%@ Register Src="../90/FlowRoute.ascx" TagName="FlowRoute" TagPrefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
<script language="javascript" type="text/javascript">
<!--
document.oncontextmenu = new Function("return false");
function window_onload() {
  <%if do_sql.G_errmsg <>"" then %>  
    alert('<%= do_sql.G_errmsg.Replace("'", "") %>');
  <% end if %>
}

// -->
</script>
</head>
<body language="javascript" onload="return window_onload()">
    <form id="form1" runat="server">
    <div>
         <fieldset id="tableB" style="width: 800px">
               <table style="width: 800px; height: 57px">
                <tr>
                    <td style="width: 289px">
                        <asp:Label ID="Label1" runat="server" Text="填表人單位：" Width="104px" ForeColor="Black" CssClass="form"></asp:Label>
                        <asp:Label ID="Lab_PWUNIT" runat="server" ForeColor="Black" Width="175px" CssClass="form"></asp:Label>
                    <td style="width: 241px">
                        <asp:Label ID="Label2" runat="server" Text="姓名：" Width="61px" ForeColor="Black" CssClass="form"></asp:Label>
                        <asp:Label ID="Lab_PWNAME" runat="server" ForeColor="Black" Width="106px" CssClass="form"></asp:Label></td>
                    <td>
                        <asp:Label ID="Label3" runat="server" Text="級職：" Width="74px" ForeColor="Black" CssClass="form"></asp:Label>
                        <asp:Label ID="Lab_PWTITLE" runat="server" ForeColor="Black" Width="175px" CssClass="form"></asp:Label></td>
                </tr>
                <tr>
                    <td style="width: 289px">
                        <asp:Label ID="Label4" runat="server" Text="申請人單位：" Width="104px" ForeColor="Black" CssClass="form"></asp:Label>
                        <asp:Label ID="Lab_PAUNIT" runat="server" ForeColor="Black" Width="175px" CssClass="form"></asp:Label></td>
                    <td style="width: 241px">
                        <asp:Label ID="Label5" runat="server" Text="姓名：" Width="58px" ForeColor="Black" CssClass="form"></asp:Label>&nbsp;
                        <asp:DropDownList ID="DrDown_PANAME" runat="server" Width="143px" AutoPostBack="True">
                        </asp:DropDownList>
                    <td>
                        <asp:Label ID="Label6" runat="server" Text="級職：" Width="74px" ForeColor="Black" CssClass="form"></asp:Label>
                        <asp:Label ID="Lab_PATITLE" runat="server" ForeColor="Black" Width="175px" CssClass="form"></asp:Label></td>
                </tr>                
            </table>  
             <asp:Label ID="Lab_eformsn" runat="server" CssClass="form" ForeColor="Black" Text="申請時間："
                 Visible="False"></asp:Label></fieldset> 
                 
            <table border="0" style="width: 800px; height: 57px; color:Red" >
                <tr>
                    <td style="width: 96px; height: 23px;">*
                        <asp:Label ID="Label7" runat="server" ForeColor="Black" Text="申請時間：" CssClass="form" ></asp:Label>
                        &nbsp; &nbsp;</td> 
                    <td style="height: 26px; width: 640px;" >
                        <asp:Label ID="Label8" runat="server" ForeColor="Black" CssClass="form" Width="252px" ></asp:Label></td>                                                          
                </tr>
                <tr>
                    <td style="width: 96px">*
                        <asp:Label ID="Label9" runat="server" ForeColor="Black" Text="軍線：" CssClass="form"></asp:Label></td>
                    <td style="height: 26px; width: 640px;" >
                        <asp:TextBox ID="TXT_nTel" runat="server"  MaxLength="10"></asp:TextBox></td>                     
                </tr>
                <tr>
                    <td style="width: 96px">*
                        <asp:Label ID="Label11" runat="server" ForeColor="Black" Text="儲位：" CssClass="form"></asp:Label></td>
                    <td style="height: 26px; width: 640px;" >
                        <asp:TextBox ID="Txt_nSeat" runat="server"  MaxLength="50"></asp:TextBox></td>                     
                </tr>
            </table> 
            
              <table border="0" style="width: 790px; height: 57px">                   
                  <tr>
                    <td style="width: 75px; height: 23px;">
                        <asp:Label ID="Label18" runat="server" ForeColor="Black" Text="財產編號" CssClass="form" ></asp:Label></td>
                    <td style="width: 128px; height: 23px;">
                        <asp:Label ID="Label19" runat="server" ForeColor="Black" Text="財產名稱" CssClass="form" ></asp:Label></td> 
                    <td style="width: 128px; height: 23px;">
                        <asp:Label ID="Label12" runat="server" ForeColor="Black" Text="Bios密碼" CssClass="form" ></asp:Label></td>  
                    <td style="width: 132px; height: 23px;">
                        <asp:Label ID="Label13" runat="server" ForeColor="Black" Text="標籤" CssClass="form" ></asp:Label></td>     
                    <td style="width: 128px; height: 23px;">
                        <asp:Label ID="Label14" runat="server" ForeColor="Black" Text="標籤號碼" CssClass="form" ></asp:Label></td>                         
                    <td style="width: 63px; height: 23px;">
                        <asp:Label ID="Label21" runat="server" ForeColor="Black" Text="問題類別" CssClass="form" ></asp:Label></td>   
                    <td style="width: 178px; height: 23px;">
                        <asp:Label ID="Label10" runat="server" ForeColor="Black" Text="說明" CssClass="form" ></asp:Label></td>                         
                    <td style="width: 80px; height: 23px;"></td>
                                    
                  </tr> 
                  <tr>
                    <td style="width: 75px; height: 23px;">
                        <asp:TextBox ID="Txt_nAssetNum" runat="server"  MaxLength="110" Width=80px></asp:TextBox></td>
                    <td style="width: 128px; height: 23px;">
                        <asp:TextBox ID="Txt_nAssetName" runat="server"  MaxLength="110" Width=80px></asp:TextBox></td>   
                     <td style="width: 128px; height: 23px;">
                        <asp:TextBox ID="Txt_nBios" runat="server"  MaxLength="20" Width=80px></asp:TextBox></td>                    
                     <td style="width: 132px; height: 23px;">
                        <asp:DropDownList ID="DrDown_nLabel" runat="server" AutoPostBack="True" Width="100px">
                        </asp:DropDownList></td>
                     <td style="width: 128px; height: 23px;">
                        <asp:TextBox ID="Txt_nLabelNum" runat="server"  MaxLength="20" Width=80px></asp:TextBox></td>    
                     <td style="width: 63px; height: 23px;">
                        <asp:DropDownList ID="DrDown_nREASON" runat="server" AutoPostBack="True" Width="120px">
                        </asp:DropDownList></td>
                     <td style="width: 178px; height: 23px;">
                        <asp:TextBox ID="TXT_nContent" runat="server"  MaxLength="255" Width="130px"></asp:TextBox></td> 
                     <td style="width: 80px; height: 23px;">
                         <asp:Button ID="But_ins" runat="server" Text="新增" CssClass="form" /></td>                  
                  </tr>
              </table> 
    
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderStyle="Double"
            CssClass="form" DataKeyNames="Info_Num" DataSourceID="SqlDataSource1" Width="787px">
            <Columns>
                <asp:BoundField DataField="Info_Num" HeaderText="Info_Num" InsertVisible="False"
                    ReadOnly="True" SortExpression="Info_Num" Visible="False" />
                <asp:BoundField DataField="EFORMSN" HeaderText="EFORMSN" SortExpression="EFORMSN"
                    Visible="False" />
                <asp:BoundField DataField="nAssetNum" HeaderText="財產編號" SortExpression="nAssetNum">
                    <ItemStyle Width="90px" />
                </asp:BoundField>
                <asp:BoundField DataField="nAssetName" HeaderText="財產名稱" SortExpression="nAssetName">
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="nBios" HeaderText="Bios密碼" SortExpression="nAssetName">
                    <ItemStyle Width="90px" />
                </asp:BoundField>
                <asp:BoundField DataField="nLabel" HeaderText="標籤" SortExpression="nAssetName">
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="nLabelNum" HeaderText="標籤號碼" SortExpression="nAssetName">
                    <ItemStyle Width="70px" />
                </asp:BoundField>
                <asp:BoundField DataField="nAmount" HeaderText="數量" SortExpression="nAmount">
                    <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" Width="40px" />
                </asp:BoundField>
                <asp:BoundField DataField="nReason" HeaderText="問題類別" SortExpression="nReason">
                    <ItemStyle Width="120px" />
                </asp:BoundField>
                <asp:BoundField DataField="nContent" HeaderText="說明" SortExpression="nContent">
                    <ItemStyle Width="160px" />
                </asp:BoundField>
                <asp:CommandField ShowDeleteButton="True" >
                    <ItemStyle Width="40px" />
                </asp:CommandField>
            </Columns>
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [P_0701] WHERE [Info_Num] = @Info_Num" InsertCommand="INSERT INTO [P_0701] ([EFORMSN], [nAssetNum], [nAssetName], [nReason], [nAmount], [nContent]) VALUES (@EFORMSN, @nAssetNum, @nAssetName, @nReason, @nAmount, @nContent)"
            SelectCommand="SELECT [Info_Num], [EFORMSN], [nAssetNum], [nAssetName], [nBios], [nLabel], [nLabelNum], [nAmount], [nREASON], [nContent] FROM [P_0701] WHERE ([EFORMSN] = @EFORMSN)"
            UpdateCommand="UPDATE [P_0701] SET [EFORMSN] = @EFORMSN, [nAssetNum] = @nAssetNum, [nAssetName] = @nAssetName, [nReason] = @nReason, [nAmount] = @nAmount, [nContent] = @nContent WHERE [Info_Num] = @Info_Num">
            <DeleteParameters>
                <asp:Parameter Name="Info_Num" Type="Decimal" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="EFORMSN" Type="String" />
                <asp:Parameter Name="nAssetNum" Type="String" />
                <asp:Parameter Name="nAssetName" Type="String" />
                <asp:Parameter Name="nReason" Type="String" />
                <asp:Parameter Name="nAmount" Type="Int32" />
                <asp:Parameter Name="nContent" Type="String" />
                <asp:Parameter Name="Info_Num" Type="Decimal" />
            </UpdateParameters>
            <SelectParameters>
                <asp:ControlParameter ControlID="Lab_eformsn" Name="EFORMSN" PropertyName="Text"
                    Type="String" />
            </SelectParameters>
            <InsertParameters>
                <asp:Parameter Name="EFORMSN" Type="String" />
                <asp:Parameter Name="nAssetNum" Type="String" />
                <asp:Parameter Name="nAssetName" Type="String" />
                <asp:Parameter Name="nReason" Type="String" />
                <asp:Parameter Name="nAmount" Type="Int32" />
                <asp:Parameter Name="nContent" Type="String" />
            </InsertParameters>
        </asp:SqlDataSource>
        <table border="0" style="width:100%; height: 57px">
        
                <% If read_only = "2" Then%>
                <tr>
                    <td style="height: 26px; width: 160px;" >
                        &nbsp;<asp:Label ID="Label28" runat="server" ForeColor="Black" Text="批核意見：" CssClass="form" Width="100px"></asp:Label><br />
                    </td>
                    <td colspan=3 >
                        <asp:TextBox ID="txtcomment" runat="server" Height="59px" MaxLength="255" Rows="3"
                            TextMode="MultiLine" Width="620px"></asp:TextBox>
                    </td>
                </tr> 
                <% End If%>
                
                <tr>
                   <td align="center" colspan="4">
                  <asp:Button ID="But_exe" runat="server" Text="送件" />&nbsp;<asp:Button ID="backBtn"
                      runat="server" Text="駁回" />&nbsp;<asp:Button ID="tranBtn" runat="server" Text="呈轉" /></td>
                </tr> 
              </table>
       
        </DIV>
        <uc1:FlowRoute ID="FlowRoute1" runat="server" />
    </form>
</body>
</html>
