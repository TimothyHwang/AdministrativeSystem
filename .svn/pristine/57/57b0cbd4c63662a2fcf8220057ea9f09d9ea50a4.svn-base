<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA06001.aspx.vb" Inherits="Source_06_MOA06001" %>

<%@ Register Src="../90/FlowRoute.ascx" TagName="FlowRoute" TagPrefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>資訊設備媒體攜出入申請單</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" language="javascript">
        document.oncontextmenu = new Function("return false");
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <table width="770" border=1 cellspacing=0 cellpadding=5 bgcolor=#ffffff bordercolor=#6699cc bordercolorlight=#74a3d6 bordercolordark=#000000 style="LEFT: 20px; TOP: 10px">
        <tr>
            <td valign=bottom bgcolor=#6699cc bordercolorlight=#66aaaa bordercolordark=#ffffff><font color=white><b>&nbsp;資訊設備媒體攜出入申請單</b></font>
            </td>
        </tr>
        
        <tr>
            <td>
    
            <fieldset id="tableB" style="width: 760px">
            <table style="width: 750px; height: 57px">
                <tr>
                    <td style="width: 250px">
                        <asp:Label ID="Label1" runat="server" Text="填表人單位：" Width="80px" ForeColor="Black" CssClass="form"></asp:Label>
                        <asp:Label ID="Lab_PWUNIT" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label>
                    <td style="width: 241px">
                        <asp:Label ID="Label2" runat="server" Text="姓名：" Width="61px" ForeColor="Black" CssClass="form"></asp:Label>
                        <asp:Label ID="Lab_PWNAME" runat="server" ForeColor="Black" Width="106px" CssClass="form"></asp:Label></td>
                    <td>
                        <asp:Label ID="Label3" runat="server" Text="級職：" Width="80px" ForeColor="Black" CssClass="form"></asp:Label>
                        <asp:Label ID="Lab_PWTITLE" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label></td>
                </tr>
                <tr>
                    <td style="width: 250px">
                        <asp:Label ID="Label4" runat="server" Text="申請人單位：" Width="80px" ForeColor="Black" CssClass="form"></asp:Label>
                        <asp:Label ID="Lab_PAUNIT" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label></td>
                    <td style="width: 241px">
                        <asp:Label ID="Label5" runat="server" Text="姓名：" Width="58px" ForeColor="Black" CssClass="form"></asp:Label>&nbsp;
                        <asp:DropDownList ID="DrDown_PANAME" runat="server" Width="143px" AutoPostBack="True">
                        </asp:DropDownList>
                    <td>
                        <asp:Label ID="Label6" runat="server" Text="級職：" Width="80px" ForeColor="Black" CssClass="form"></asp:Label>
                        <asp:Label ID="Lab_PATITLE" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label></td>
                </tr>                
            </table>  
            <asp:Label ID="Lab_eformsn" runat="server" CssClass="form" ForeColor="Black" Text="申請時間："
                 Visible="False"></asp:Label>
            </fieldset>
            
			<table border="0" style="width: 760px; height: 57px; color:Red" >
                <tr>
                    <td style="width: 160px; height: 23px;">*
                        <asp:Label ID="Label7" runat="server" ForeColor="Black" Text="申請時間：" CssClass="form" ></asp:Label></td>
                    <td style="width: 600px; height: 23px;">
                        <asp:Label ID="Label8" runat="server" ForeColor="Black" CssClass="form" Width="150px" ></asp:Label></td>                    
                </tr> 
                <tr>
                    <td style="width: 160px">*
                        <asp:Label ID="Label9" runat="server" ForeColor="Black" Text="申請事由：" CssClass="form"></asp:Label></td>
                    <td style="width: 600px"  >
                        <asp:DropDownList ID="DrDown_nREASON" runat="server" AutoPostBack="True">
                        </asp:DropDownList></td>                        
                </tr> 
                <tr>
                    <td style="height: 30px; width: 160px;">*
                        <asp:Label ID="Label14" runat="server" ForeColor="Black" Text="申請出入日期：" CssClass="form" ></asp:Label></td>
                    <td style="height: 30px; width: 600px;" >
                        <asp:TextBox ID="Txt_nDATE" runat="server" EnableTheming="False" CausesValidation="True" Width="89px" AutoPostBack="True" OnKeyDown="return false"></asp:TextBox>
                        <asp:ImageButton
                            ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" /></tr> 
                <tr>
                    <td style="height: 26px; width: 160px;" >*
                        <asp:Label ID="Label12" runat="server" ForeColor="Black" Text="使用地點：" CssClass="form"></asp:Label></td>
                    <td style="height: 26px; width: 600px;" >
                        <asp:TextBox ID="TXT_nPLACE" runat="server"  MaxLength="25" Width="400px"></asp:TextBox></td>
                </tr> 
                
            </table>  
                
             <table border="0" style="width: 760px;">
              <tr>
                  <td>
                  <table border="0" style="width: 750px; height: 57px">                   
                      <tr>
                        <td style="width: 80px; height: 23px;">
                            <asp:Label ID="Label18" runat="server" ForeColor="Black" Text="區分：" CssClass="form" ></asp:Label></td>
                        <td style="width: 120px; height: 23px;">
                            <asp:Label ID="Label19" runat="server" ForeColor="Black" Text="名稱機型：" CssClass="form" ></asp:Label></td>  
                        <td style="width: 120px; height: 23px;">
                            <asp:Label ID="Label20" runat="server" ForeColor="Black" Text="編(文)號：" CssClass="form" ></asp:Label></td>  
                        <td style="width: 61px; height: 23px;">
                            <asp:Label ID="Label21" runat="server" ForeColor="Black" Text="數量：" CssClass="form" ></asp:Label></td>   
                        <td style="width: 190px; height: 23px;">
                            <asp:Label ID="Label10" runat="server" ForeColor="Black" Text="內容概要：" CssClass="form" ></asp:Label></td>   
                        <td style="width: 90px; height: 23px;">
                            <asp:Label ID="Label11" runat="server" ForeColor="Black" Text="機密等級：" CssClass="form" ></asp:Label></td>   
                        <td style="width: 60px; height: 23px;"></td>
                                        
                      </tr> 
                      <tr>
                        <td>
                            <asp:DropDownList ID="DrDown_nKind" runat="server" AutoPostBack="True" Width="82px">
                            </asp:DropDownList></td>
                        <td>
                            <asp:TextBox ID="Txt_nMName" runat="server"  MaxLength="14" Width=110px></asp:TextBox></td>                    
                         <td>
                            <asp:TextBox ID="Txt_nDocnum" runat="server"  MaxLength="10" Width=113px></asp:TextBox></td> 
                         <td style="width: 61px">
                            <asp:TextBox ID="Txt_nAmount" runat="server"  MaxLength="4" Width="42px"></asp:TextBox></td> 
                         <td style="width: 190px">
                            <asp:TextBox ID="Txt_nContent" runat="server"  MaxLength="32" Width="170px"></asp:TextBox></td> 
                         <td style="width: 90px">
                            <asp:DropDownList ID="DrDown_nClass" runat="server" AutoPostBack="True" Width="82px">
                            </asp:DropDownList></td> 
                         <td style="width: 60px">
                             <asp:Button ID="But_ins" runat="server" Text="新增" /></td>                  
                      </tr>
                  </table> 
                  <asp:GridView ID="GridView1" runat="server" Width="787px" AutoGenerateColumns="False" BorderStyle="Double" CssClass="form" DataSourceID="SqlDataSource1" DataKeyNames="Info_Num" ForeColor="Black">
                <Columns>
                    <asp:BoundField DataField="Info_Num" HeaderText="Info_Num" InsertVisible="False"
                        ReadOnly="True" SortExpression="Info_Num" Visible="False" />
                    <asp:BoundField DataField="EFORMSN" HeaderText="EFORMSN" SortExpression="EFORMSN"
                        Visible="False" />
                    <asp:BoundField DataField="nKind" HeaderText="區分" SortExpression="nKind" >
                        <ItemStyle Width="87px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="nMName" HeaderText="名稱機型" SortExpression="nMName" >
                        <ItemStyle Width="120px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="nDocnum" HeaderText="編(文)號" SortExpression="nDocnum" >
                        <ItemStyle Width="120px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="nAmount" HeaderText="數量" SortExpression="nAmount">
                        <ItemStyle HorizontalAlign="Right" Width="63px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="nContent" HeaderText="內容概要" SortExpression="nContent">
                        <ItemStyle HorizontalAlign="Left" Width="150px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="nClass" HeaderText="機密等級" SortExpression="nClass">
                        <ItemStyle HorizontalAlign="Left" Width="82px" />
                    </asp:BoundField>
                    <asp:CommandField ShowDeleteButton="True" >
                        <ItemStyle Width="80px" />
                    </asp:CommandField>
                </Columns>
            </asp:GridView> 
              <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                  DeleteCommand="DELETE FROM [P_0601] WHERE [Info_Num] = @Info_Num" InsertCommand="INSERT INTO [P_0601] ([EFORMSN], [nKind], [nMName], [nDocnum], [nAmount], [nContent], [nClass]) VALUES (@EFORMSN, @nKind, @nMName, @nDocnum, @nAmount, @nContent, @nClass)"
                  SelectCommand="SELECT [Info_Num], [EFORMSN], [nKind], [nMName], [nDocnum], [nAmount], [nContent], [nClass] FROM [P_0601] WHERE ([EFORMSN] = @EFORMSN)"
                  UpdateCommand="UPDATE [P_0601] SET [EFORMSN] = @EFORMSN, [nKind] = @nKind, [nMName] = @nMName, [nDocnum] = @nDocnum, [nAmount] = @nAmount, [nContent] = @nContent, [nClass] = @nClass WHERE [Info_Num] = @Info_Num">
                  <DeleteParameters>
                      <asp:Parameter Name="Info_Num" Type="Decimal" />
                  </DeleteParameters>
                  <UpdateParameters>
                      <asp:Parameter Name="EFORMSN" Type="String" />
                      <asp:Parameter Name="nKind" Type="String" />
                      <asp:Parameter Name="nMName" Type="String" />
                      <asp:Parameter Name="nDocnum" Type="String" />
                      <asp:Parameter Name="nAmount" Type="Int32" />
                      <asp:Parameter Name="nContent" Type="String" />
                      <asp:Parameter Name="nClass" Type="String" />
                      <asp:Parameter Name="Info_Num" Type="Decimal" />
                  </UpdateParameters>
                  <SelectParameters>
                      <asp:ControlParameter ControlID="Lab_eformsn" Name="EFORMSN" PropertyName="Text"
                          Type="String" />
                  </SelectParameters>
                  <InsertParameters>
                      <asp:Parameter Name="EFORMSN" Type="String" />
                      <asp:Parameter Name="nKind" Type="String" />
                      <asp:Parameter Name="nMName" Type="String" />
                      <asp:Parameter Name="nDocnum" Type="String" />
                      <asp:Parameter Name="nAmount" Type="Int32" />
                      <asp:Parameter Name="nContent" Type="String" />
                      <asp:Parameter Name="nClass" Type="String" />
                  </InsertParameters>
              </asp:SqlDataSource>
                      &nbsp;
              <table border="0" style="width: 100%; height: 57px">
              
                <% If read_only = "2" Then%>
                <tr>
                    <td style="height: 26px; width: 160px;" >
                        &nbsp;<asp:Label ID="Label28" runat="server" ForeColor="Black" Text="批核意見：" CssClass="form"></asp:Label></td>
                    <td colspan=3 >
                        <asp:TextBox ID="txtcomment" runat="server" Height="59px" MaxLength="255" Rows="3"
                            TextMode="MultiLine" Width="520px"></asp:TextBox>
                        <asp:Button ID="But_PHRASE" runat="server" Text="批核片語" /></td>
                </tr> 
                <% End If%>
              
                <tr>
                   <td align="center" colspan="4">
                  <asp:Button ID="But_exe" runat="server" Text="送件" Visible="False" />&nbsp;<asp:Button ID="backBtn"
                      runat="server" Text="駁回" />&nbsp;<asp:Button ID="tranBtn" runat="server" Text="呈轉" />
                       <asp:ImageButton ID="ImgPrint" runat="server" ImageUrl="~/Image/SavePrint.gif" ToolTip="列印" /></td>
                </tr> 
              </table>
                      <uc1:FlowRoute ID="FlowRoute1" runat="server" />
                  </td>
              </tr> 
            </table>
        
              </td>
          </tr> 
        </table>     
    </div>
    
    <div id="Div_grid" runat="server" style="border-right: lightslategray 2px solid;
        border-top: lightslategray 2px solid; display: block; z-index: 3; left: 14px;
        border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
        position: absolute; top: 840px; height: 150pt; background-color: white" visible="false">
        <asp:Calendar ID="Calendar1" runat="server" BackColor="White" BorderColor="#3366CC"
            BorderWidth="1px" Caption="" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana"
            Font-Size="8pt" ForeColor="#003399" Height="200px" ShowGridLines="True" Width="220px">
            <SelectedDayStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
            <TodayDayStyle BackColor="#99CCCC" ForeColor="White" />
            <SelectorStyle BackColor="#99CCCC" ForeColor="#336666" />
            <WeekendDayStyle BackColor="#CCCCFF" />
            <OtherMonthDayStyle ForeColor="#999999" />
            <NextPrevStyle Font-Size="8pt" ForeColor="#CCCCFF" />
            <DayHeaderStyle BackColor="#99CCCC" ForeColor="#336666" Height="1px" />
            <TitleStyle BackColor="#003399" BorderColor="#3366CC" BorderWidth="1px" Font-Bold="True"
                Font-Size="10pt" ForeColor="#CCCCFF" Height="25px" />
        </asp:Calendar>
        <asp:Button ID="btnClose1" runat="server" Text="關閉" Width="220px" /></div>
        
       <div id="Div_grid10" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:300pt; height:200pt; left:249px; top:893px; display:block;" visible="false">
            <asp:GridView id="GridView10" runat="server" CssClass="form" Width="100%" Height="50px" DataSourceID="SqlDataSource10" PageSize="5" AutoGenerateColumns="False" AllowPaging="True" BorderColor="Lime" BorderWidth="2px">
          <Columns>
              <asp:BoundField DataField="comment" HeaderText="批核片語" >
                  <HeaderStyle HorizontalAlign="Center" Width="90%" BackColor="#80FF80" CssClass="form"  />
              </asp:BoundField>
              <asp:BoundField DataField="Phrase_num" HeaderText="Phrase_num" Visible="False" >
                  <ItemStyle Wrap="False"  />
              </asp:BoundField>
              <asp:CommandField ShowSelectButton="True">
                  <HeaderStyle HorizontalAlign="Center" Width="10%"  />
              </asp:CommandField>
          </Columns>
          <RowStyle Height="10px"  />
      </asp:GridView>
            &nbsp;
            <asp:Button ID="Btn_PHclose" runat="server" Text="關閉" Width="389px" />
      <asp:SqlDataSource id="SqlDataSource10" runat="server" SelectCommand="SELECT Phrase_num, employee_id, comment FROM PHRASE WHERE [employee_id] = @employee_id ORDER BY Phrase_num" ConnectionString="<%$ ConnectionStrings:ConnectionString %>">                      
          <SelectParameters>
              <asp:SessionParameter Name="employee_id" SessionField="user_id" Type="String" />
          </SelectParameters>
      </asp:SqlDataSource>
            </div> 
    </form>    
         
</body>
</html>
