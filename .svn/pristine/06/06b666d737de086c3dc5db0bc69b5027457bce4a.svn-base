<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04003.aspx.vb" Inherits="Source_04_MOA04003" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>完工報告單</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />  
    <script type="text/javascript" language="javascript">
        document.oncontextmenu = new Function("return false");
    </script>
</head>
<body language="javascript" onload="return window_onload()">
    <form id="form1" runat="server">    
    
        <table width="750" border=1 cellspacing=0 cellpadding=5 bgcolor=#ffffff bordercolor=#6699cc bordercolorlight=#74a3d6 bordercolordark=#000000 style="LEFT: 20px; TOP: 10px">
        <tr>
            <td valign=bottom bgcolor=#6699cc bordercolorlight=#66aaaa bordercolordark=#ffffff><font color=white><b>&nbsp;完工報告申請單</b></font>
            </td>
        </tr>
        
        <tr>
            <td>    
    
            <fieldset id="tableB" style="width: 750px">
            <table style="width: 750px; height: 57px">
                <tr>
                    <td style="width: 270px">
                        <asp:Label ID="Label1" runat="server" Text="填表人單位：" Width="90px" ForeColor="Black" CssClass="form"></asp:Label>
                        <asp:Label ID="Lab_PWUNIT" runat="server" ForeColor="Black" Width="160px" CssClass="form"></asp:Label>
                    <td style="width: 240px">
                        <asp:Label ID="Label2" runat="server" Text="姓名：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                        <asp:Label ID="Lab_PWNAME" runat="server" ForeColor="Black" Width="110px" CssClass="form"></asp:Label></td>
                    <td style="width: 240px">
                        <asp:Label ID="Label3" runat="server" Text="級職：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                        <asp:Label ID="Lab_PWTITLE" runat="server" ForeColor="Black" Width="140px" CssClass="form"></asp:Label></td>
                </tr>
                <tr>
                    <td style="width: 270px">
                        <asp:Label ID="Label4" runat="server" Text="申請人單位：" Width="90px" ForeColor="Black" CssClass="form"></asp:Label>
                        <asp:Label ID="Lab_PAUNIT" runat="server" ForeColor="Black" Width="160px" CssClass="form"></asp:Label></td>
                    <td style="width: 240px">
                        <asp:Label ID="Label5" runat="server" Text="姓名：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>&nbsp;
                        <asp:DropDownList ID="DrDown_PANAME" runat="server" Width="159px" AutoPostBack="True">
                        </asp:DropDownList>
                    <td style="width: 240px">
                        <asp:Label ID="Label6" runat="server" Text="級職：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                        <asp:Label ID="Lab_PATITLE" runat="server" ForeColor="Black" Width="140px" CssClass="form"></asp:Label></td>
                </tr>                
            </table>  
            <asp:Label ID="Lab_eformsn" runat="server" CssClass="form" ForeColor="Black" Text="申請時間："
                Visible="False"></asp:Label></fieldset>
			<table border="0" style="width: 750px; height: 57px; color:Red" >
                <tr>
                    <td style="width: 100px; height: 23px;">*
                        <asp:Label ID="Label7" runat="server" ForeColor="Black" Text="申請時間：" CssClass="form" ></asp:Label></td>
                    <td style="width: 600px; height: 23px;">
                        <asp:Label ID="Label8" runat="server" ForeColor="Black" CssClass="form" Width="248px" ></asp:Label></td>                    
                </tr> 
                <tr>
                    <td style="width: 100px">*
                        <asp:Label ID="Label9" runat="server" ForeColor="Black" Text="申請單：" CssClass="form" Width="65px"></asp:Label></td>
                    <td style="height: 26px; width: 600px;" >
                        <asp:DropDownList ID="DrDown_nFIXITEM" runat="server" CssClass="form" Width="600px" AutoPostBack="True">
                        </asp:DropDownList></td>                      
                </tr> 
                
                <tr>
                    <td></td>
                    <td >
                        <asp:TextBox ID="txtnFIXITEM" runat="server" Height="80px" TextMode="MultiLine" Width="600px" OnKeyDown="return false"></asp:TextBox></td>                      
                </tr> 
                <tr>
                    <td style="width: 100px; height: 28px;">*
                        <asp:Label ID="Label10" runat="server" ForeColor="Black" Text="材料：" CssClass="form" Width="61px"></asp:Label></td>
                    <td style="height: 28px; width: 600px;" >
                        &nbsp;<asp:RadioButton ID="Rad_nConSume1" runat="server" ForeColor="Black" Text="相符"  Checked="True" GroupName="rad" CssClass="form" />
                        <asp:RadioButton ID="Rad_nConSume2" runat="server" ForeColor="Black" Text="不相符"  Checked="False" GroupName="rad" CssClass="form" /></td>                      
                </tr> 
                <tr>
                    <td style="width: 100px">*
                        <asp:Label ID="Label11" runat="server" ForeColor="Black" Text="效果：" CssClass="form" Width="61px"></asp:Label></td>
                    <td style="height: 26px; width: 600px;" >
                        &nbsp;<asp:RadioButton ID="Rad_nEffect1" runat="server" ForeColor="Black" Text="良好"  Checked="True" GroupName="blue" CssClass="form" />
                        <asp:RadioButton ID="Rad_nEffect2" runat="server" ForeColor="Black" Text="不好"  Checked="False" GroupName="blue" CssClass="form" /></td>                      
                </tr> 
                <tr>
                    <td style="width: 100px">*
                        <asp:Label ID="Label12" runat="server" ForeColor="Black" Text="交辦日期：" CssClass="form" Width="70px"></asp:Label></td>
                    <td style="height: 26px; width: 600px;" > 
                        &nbsp;<asp:TextBox ID="Txt_nHandleDate" runat="server" EnableTheming="False" Width="89px" AutoPostBack="True" OnKeyDown="return false"></asp:TextBox>
                        <asp:ImageButton
                                ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />&nbsp;<asp:DropDownList ID="DrDown_nHandle_HOUR" runat="server">
                        </asp:DropDownList>
                        <asp:Label ID="Label14" runat="server" ForeColor="Black" Text="時" CssClass="form"></asp:Label>
                        </td>                      
                </tr> 
                <tr>
                    <td style="width: 100px; height: 28px;">*
                        <asp:Label ID="Label13" runat="server" ForeColor="Black" Text="勘查日期：" CssClass="form" Width="70px"></asp:Label></td>
                    <td style="height: 28px; width: 600px;" >
                        &nbsp;<asp:TextBox ID="Txt_nCheckDate" runat="server" EnableTheming="False" Width="89px" OnKeyDown="return false"></asp:TextBox>
                        <asp:ImageButton
                                ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" />
                        <asp:DropDownList ID="DrDown_nCheck_HOUR" runat="server">
                        </asp:DropDownList>
                        <asp:Label ID="Label15" runat="server" ForeColor="Black" Text="時" CssClass="form"></asp:Label></td>                      
                </tr> 
                <tr>
                    <td style="width: 100px">*
                        <asp:Label ID="Label16" runat="server" ForeColor="Black" Text="發料日期：" CssClass="form" Width="70px"></asp:Label></td>
                    <td style="height: 26px; width: 600px;" >
                        &nbsp;<asp:TextBox ID="Txt_nDoleDate" runat="server" EnableTheming="False" CausesValidation="True" Width="89px" OnKeyDown="return false"></asp:TextBox>
                        <asp:ImageButton
                                ID="ImgDate3" runat="server" ImageUrl="~/Image/calendar.gif" />&nbsp;<asp:DropDownList ID="DrDown_nDole_HOUR" runat="server">
                        </asp:DropDownList>
                        <asp:Label ID="Label17" runat="server" ForeColor="Black" Text="時" CssClass="form"></asp:Label></td>                      
                </tr> 
                <tr>
                    <td style="width: 100px">*
                        <asp:Label ID="Label18" runat="server" ForeColor="Black" Text="完工日期：" CssClass="form" Width="70px"></asp:Label></td>
                    <td style="height: 26px; width: 600px;" >
                        &nbsp;<asp:TextBox ID="Txt_nFinishDate" runat="server" EnableTheming="False" CausesValidation="True" Width="89px" OnKeyDown="return false"></asp:TextBox>
                        <asp:ImageButton
                                ID="ImgDate4" runat="server" ImageUrl="~/Image/calendar.gif" />&nbsp;<asp:DropDownList ID="DrDown_nFinish_HOUR" runat="server">
                        </asp:DropDownList>
                        <asp:Label ID="Label19" runat="server" ForeColor="Black" Text="時" CssClass="form"></asp:Label></td>                      
                </tr> 
            </table>  
                
            <table border="0" style="width: 750px;">
            <tr>
              <td style="width: 750px">
              <table border="0" style="width: 750px; height: 57px">                   
                  <tr>
                    <td style="width: 320px; height: 23px;" align="left" valign="middle">
                        <asp:Label ID="Label21" runat="server" ForeColor="Black" Text="物品名稱：" CssClass="form" ></asp:Label></td>
                    <td style="height: 23px; width: 190px;" align="left">
                        <asp:Label ID="Label22" runat="server" ForeColor="Black" Text="物品單位：" CssClass="form" ></asp:Label></td>  
                    <td style="height: 23px; width: 80px;" align="left">
                        <asp:Label ID="Label23" runat="server" ForeColor="Black" Text="物品數量：" CssClass="form" ></asp:Label></td>  
                    <td style="height: 23px; width: 60px;"> </td>  
                  </tr> 
                  <tr>
                    <td style="width: 320px; height: 23px;">
                        <asp:TextBox ID="Txt_Block_Name" runat="server"  MaxLength="50" Width="300px"></asp:TextBox></td>
                    <td style="width: 190px; height: 23px;">
                        <asp:TextBox ID="Txt_Block_Unit" runat="server"  MaxLength="10" Width="160px"></asp:TextBox></td>                    
                     <td style="height: 23px; width: 80px;">
                        <asp:TextBox ID="Txt_Block_Amount" runat="server"  MaxLength="4" Width="70px" ></asp:TextBox></td>   
                     <td style="height: 23px; width: 60px;"> 
                        <asp:Button ID="But_ins" runat="server" Text="新增" CssClass="form" ForeColor="Black" /></td>                                       
                  </tr> 
              </table> 
              <asp:GridView ID="GridView1" runat="server" Width="720px" AutoGenerateColumns="False" BorderStyle="Double" CssClass="form" DataSourceID="SqlDataSource1" DataKeyNames="Block_Num">
                <Columns>
                <asp:BoundField DataField="Block_Num" HeaderText="Block_Num" InsertVisible="False"
                    ReadOnly="True" SortExpression="Block_Num" Visible="False" />
                <asp:BoundField DataField="EFORMSN" HeaderText="EFORMSN" SortExpression="EFORMSN" Visible="False" />
                <asp:BoundField DataField="Block_Name" HeaderText="品名" SortExpression="Block_Name" >
                    <ItemStyle Width="365px" />
                </asp:BoundField>
                <asp:BoundField DataField="Block_Unit" HeaderText="單位" SortExpression="Block_Unit" />
                <asp:BoundField DataField="Block_Amount" HeaderText="數量" SortExpression="Block_Amount" >
                    <ItemStyle HorizontalAlign="Right" Width="100px" />
                </asp:BoundField>
                <asp:CommandField ShowDeleteButton="True" ButtonType="Image" DeleteImageUrl="~/Image/delete.gif">
                    <ItemStyle Width="90px" />
                </asp:CommandField>
                </Columns>
                  <HeaderStyle BackColor="SkyBlue" />
                </asp:GridView> 
                  <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                      DeleteCommand="DELETE FROM [P_0402] WHERE [Block_Num] = @Block_Num" 
                      SelectCommand="SELECT [Block_Num], [EFORMSN], [Block_Name], [Block_Unit], [Block_Amount] FROM [P_0402] WHERE ([EFORMSN] = @EFORMSN) ORDER BY [Block_Num]">
                      <DeleteParameters>
                          <asp:Parameter Name="Block_Num" Type="Decimal" />
                      </DeleteParameters>                      
                      <SelectParameters>
                          <asp:ControlParameter ControlID="Lab_eformsn" Name="EFORMSN" PropertyName="Text"
                              Type="String" />
                      </SelectParameters>                      
                  </asp:SqlDataSource>
                  &nbsp; &nbsp;
                  <table border="0" style="width: 730px; height: 57px">
                  
                    <% If read_only = "2" Then%>
                    <tr>
                        <td style="height: 26px; width: 160px;" >
                            &nbsp;<asp:Label ID="Label28" runat="server" ForeColor="Black" Text="批核意見：" CssClass="form"></asp:Label><br />
                        </td>
                        <td >
                            <asp:TextBox ID="txtcomment" runat="server" Height="59px" MaxLength="255" Rows="3"
                                TextMode="MultiLine" Width="464px"></asp:TextBox>
                            <asp:Button ID="But_PHRASE" runat="server" Text="批核片語" /></td>
                    </tr> 
                    <% End If%>
                    
                    <tr>
                       <td style="width: 100%; height: 23px;"  align="center" colspan="2">
                      <asp:Button ID="But_exe" runat="server" Text="送件" />
                           <asp:Button ID="backBtn" runat="server" Text="駁回" />&nbsp;<asp:Button ID="tranBtn"
                               runat="server" Text="呈轉" /></td>
                    </tr> 
                  </table>
              </td>
            </tr> 
            </table>  
          
          
            </td>
        </tr>
        </table>
          
    
    <div id="Div_grid" runat="server" style="border-right: lightslategray 2px solid;
        border-top: lightslategray 2px solid; display: block; z-index: 3; left: 13px;
        border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
        position: absolute; top: 787px; height: 150pt; background-color: white" visible="false">
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
        
    <div id="Div_grid2" runat="server" style="border-right: lightslategray 2px solid;
        border-top: lightslategray 2px solid; display: block; z-index: 3; left: 248px;
        border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
        position: absolute; top: 789px; height: 150pt; background-color: white" visible="false">
        <asp:Calendar ID="Calendar2" runat="server" BackColor="White" BorderColor="#3366CC"
            BorderWidth="1px" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana"
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
        <asp:Button ID="btnClose2" runat="server" Text="關閉" Width="220px" /></div>
        
    <div id="Div_grid3" runat="server" style="border-right: lightslategray 2px solid; border-top: lightslategray 2px solid;
        display: block; z-index: 3; left: 481px; border-left: lightslategray 2px solid;
        width: 165pt; border-bottom: lightslategray 2px solid; position: absolute; top: 791px;
        height: 150pt; background-color: white" visible="false">
        <asp:Calendar ID="Calendar3" runat="server" BackColor="White" BorderColor="#3366CC"
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
        <asp:Button ID="btnClose3" runat="server" Text="關閉" Width="220px" /></div>
        
    <div id="Div_grid4" runat="server" style="border-right: lightslategray 2px solid; border-top: lightslategray 2px solid;
        display: block; z-index: 3; left: 712px; border-left: lightslategray 2px solid;
        width: 165pt; border-bottom: lightslategray 2px solid; position: absolute; top: 791px;
        height: 150pt; background-color: white" visible="false">
        <asp:Calendar ID="Calendar4" runat="server" BackColor="White" BorderColor="#3366CC"
            BorderWidth="1px" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana"
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
        <asp:Button ID="btnClose4" runat="server" Text="關閉" Width="220px" /></div>
        
        <div id="Div_grid10" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:300pt; height:80pt; left:249px; top:893px; display:block;" visible="false">
            <asp:GridView id="GridView10" runat="server" CssClass="form" Width="100%" Height="50px" DataSourceID="SqlDataSource10" PageSize="7" AutoGenerateColumns="False" AllowPaging="True" BorderColor="Lime" BorderWidth="2px">
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
    <script language=javascript>
    
function window_onload() {
     <%if do_sql.G_errmsg <>"" then %>  
          alert('<%= do_sql.G_errmsg%>');
     <%end if    %>
 }

    </script>

</body>
</html>
