<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA01001.aspx.vb" Inherits="Source_01_MOA01001" %>

<%@ Register Src="../90/FlowRoute.ascx" TagName="FlowRoute" TagPrefix="uc1" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>差假申請單</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />    
    <script type="text/javascript" language="javascript">
        document.oncontextmenu = new Function("return false");
    </script>
</head>
<body language="javascript" onload="return window_onload()">
    <form id="form1" runat="server">
     <div>
     
        <table width="740" border=1 cellspacing=0 cellpadding=5 bgcolor=#ffffff bordercolor=#6699cc bordercolorlight=#74a3d6 bordercolordark=#000000 style="LEFT: 20px; TOP: 10px">
        <tr>
            <td valign=bottom bgcolor=#6699cc bordercolorlight=#66aaaa bordercolordark=#ffffff style="height: 33px; width: 727px;"><font color=white><b>&nbsp;差假申請單</b></font>
            </td>
        </tr>
        
        <tr>
            <td style="width: 727px">
            
            <fieldset id="tableB" style="width: 740px">
               <table style="width: 740px">
                    <tr>
                        <td style="width: 230px; height: 24px;">
                            <asp:Label ID="Label1" runat="server" Text="填表人單位：" Width="80px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_ORG_NAME_1" runat="server" ForeColor="Black" Width="140px" CssClass="form"></asp:Label>
                        <td style="width: 250px; height: 24px;">
                            <asp:Label ID="Label2" runat="server" Text="姓名：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_emp_chinese_name" runat="server" ForeColor="Black" Width="190px" CssClass="form"></asp:Label></td>
                        <td style="width: 210px">
                            <asp:Label ID="Label3" runat="server" Text="級職：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_title_name_1" runat="server" ForeColor="Black" Width="140px" CssClass="form"></asp:Label></td>
                    </tr>
                    <tr>
                        <td style="width: 230px">
                            <asp:Label ID="Label4" runat="server" Text="申請人單位：" Width="80px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_ORG_NAME_2" runat="server" ForeColor="Black" Width="140px" CssClass="form"></asp:Label></td>
                        <td style="width: 250px">
                            <asp:Label ID="Label5" runat="server" Text="姓名：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:DropDownList ID="DrDown_emp_chinese_name" runat="server" Width="143px" AutoPostBack="True">
                            </asp:DropDownList>
                        <td style="width: 210px">
                            <asp:Label ID="Label6" runat="server" Text="級職：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_title_name_2" runat="server" ForeColor="Black" Width="140px" CssClass="form"></asp:Label></td>
                    </tr>                
                </table>  
		    </fieldset> 
                        
            <table border="0" style="width: 740px; height: 57px; color:Red" >
                    <tr>
                        <td style="width: 120px; height: 23px;">*
                            <asp:Label ID="Label7" runat="server" ForeColor="Black" Text="申請類別：" CssClass="form" ></asp:Label></td>
                        <td style="width: 260px; height: 23px;">
                            <asp:RadioButton ID="Radio1" runat="server" ForeColor="Black" Text="假前申請"  Checked="True" GroupName="rad" CssClass="form" AutoPostBack="True" />
                            <asp:RadioButton ID="Radio2" runat="server" ForeColor="Black" Text="假後補請"  GroupName="rad" CssClass="form" AutoPostBack="True" />
                            <asp:Label ID="Label29" runat="server" Font-Size="10pt" Text="(當日視同)"></asp:Label>
                        </td>
                        <td style="width: 120px; height: 23px;">&nbsp;&nbsp;
                            <asp:Label ID="Label15" runat="server" ForeColor="Black" Text="申請時間：" CssClass="form" ></asp:Label></td>
                        <td style="width: 200px; height: 23px;">
                            <asp:Label ID="Lab_time" runat="server" ForeColor="Black" CssClass="form" ></asp:Label></td>
                    </tr> 
                    <tr>
                        <td style="width: 120px">*
                            <asp:Label ID="Label8" runat="server" ForeColor="Black" Text="假別：" CssClass="form"></asp:Label></td>
                        <td style="width: 260px"  >
                            <asp:DropDownList ID="DrDown_nTYPE" runat="server" AutoPostBack="True">
                            </asp:DropDownList></td>
                        <td style="width: 120px" >&nbsp;&nbsp;
                            <asp:Label ID="Label16" runat="server" ForeColor="Black" Text="證明：" CssClass="form"></asp:Label></td>
                        <td style="width: 200px" ><asp:RadioButton ID="RadioButton2" runat="server" ForeColor="Black" Text="有" GroupName="pro" CssClass="form" /><asp:RadioButton ID="RadioButton3" runat="server" ForeColor="Black" Text="無" Checked="True" GroupName="pro" CssClass="form" /></td>
                    </tr> 
                    <tr>
                        <td style="height: 23px; width: 120px;">*
                            <asp:Label ID="Label9" runat="server" ForeColor="Black" Text="起始時間：" CssClass="form" ></asp:Label></td>
                        <td style="height: 23px; width: 260px;" >
                            <asp:TextBox ID="Txt_nSTARTTIME" runat="server" EnableTheming="False" CausesValidation="True" OnKeyDown="return false" Width="80px"></asp:TextBox>
                            <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />
                            <asp:DropDownList ID="DrDown_nSTHOUR" runat="server" AutoPostBack="True">
                            </asp:DropDownList>
                            <asp:Label ID="Label14" runat="server" ForeColor="Black" Text="時" CssClass="form"></asp:Label>
                            </td>
                        <td style="height: 23px; width: 120px;">*
                            <asp:Label ID="Label17" runat="server" ForeColor="Black" Text="結束時間：" CssClass="form"></asp:Label></td>
                        <td style="height: 23px; width: 200px;">
                            <asp:TextBox ID="Txt_nENDTIME" runat="server" EnableTheming="False" OnKeyDown="return false" Width="80px" ></asp:TextBox>
                            <asp:ImageButton ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" />
                            <asp:DropDownList ID="DrDown_nETHOUR" runat="server" AutoPostBack="True">
                            </asp:DropDownList>
                            <asp:Label ID="Label23" runat="server" ForeColor="Black" Text="時" CssClass="form"></asp:Label>
                            </td>
                    </tr> 
                    <tr>
                        <td style="width: 120px; height: 23px;" >&nbsp;&nbsp;
                            <asp:Label ID="Label10" runat="server" ForeColor="Black" Text="請假天數：" CssClass="form"></asp:Label></td>
                        <td style="width: 260px; height: 23px;" >
                            <asp:Label ID="Label21" runat="server" ForeColor="Black" Text="共" CssClass="form" Width="20px"></asp:Label>
                            <asp:Label ID="Label22" runat="server" ForeColor="Black" Width="20px" CssClass="form"></asp:Label>
                            <asp:Label ID="Label24" runat="server" ForeColor="Black" Text="天" CssClass="form" Width="20px"></asp:Label>
                            <asp:Label ID="Label25" runat="server" ForeColor="Black" Width="20px" CssClass="form"></asp:Label>
                            <asp:Label ID="Label26" runat="server" ForeColor="Black" Text="小時" CssClass="form" Width="30px"></asp:Label></td>
                        <td style="width: 120px; height: 23px;" ></td>
                        <td style="width: 200px; height: 23px;"  >
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td style="width: 120px; height: 26px;" >*
                            <asp:Label ID="Label18" runat="server" ForeColor="Black" Text="職務代理人(一)：" CssClass="form"></asp:Label></td>
                            <td style="height: 26px; width: 260px;" >
                            <asp:DropDownList ID="DrDown_nAGENT1" runat="server">
                            </asp:DropDownList>
                            <asp:DropDownList ID="DrDown_nAGENT1_title" runat="server" Visible="False">
                            </asp:DropDownList></td>
                        <td style="height: 26px; width: 120px;" ></td>
                        <td style="height: 26px; width: 200px;" ></td>
                    </tr> 
                    <tr>
                        <td style="width: 120px; height: 26px;"  >&nbsp;&nbsp;
                            <asp:Label ID="Label11" runat="server" ForeColor="Black" Text="職務代理人(二)：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 260px;" >
                            <asp:DropDownList ID="DrDown_nAGENT2" runat="server">
                            </asp:DropDownList></td>
                        <td style="height: 26px; width: 120px;" >&nbsp;&nbsp;
                            <asp:Label ID="Label19" runat="server" ForeColor="Black" Text="職務代理人(三)：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 200px;" >
                            <asp:DropDownList ID="DrDown_nAGENT3" runat="server">
                            </asp:DropDownList></td>
                    </tr> 
                    <tr>
                        <td style="height: 26px; width: 120px;" >*
                            <asp:Label ID="Label12" runat="server" ForeColor="Black" Text="到達地點：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 260px;" >
                            <asp:DropDownList ID="DDL_nPlace" runat="server">
                                <asp:ListItem Selected="True">請選擇</asp:ListItem>
                                <asp:ListItem>台北市</asp:ListItem>
                                <asp:ListItem>新北市</asp:ListItem>
                                <asp:ListItem>高雄市</asp:ListItem>
                                <asp:ListItem>桃園縣市</asp:ListItem>
                                <asp:ListItem>新竹縣市</asp:ListItem>
                                <asp:ListItem>台中市</asp:ListItem>
                                <asp:ListItem>南投縣</asp:ListItem>
                                <asp:ListItem>苗栗縣</asp:ListItem>
                                <asp:ListItem>彰化縣</asp:ListItem>
                                <asp:ListItem>雲林縣</asp:ListItem>
                                <asp:ListItem>嘉義縣市</asp:ListItem>
                                <asp:ListItem>台南市</asp:ListItem>
                                <asp:ListItem>屏東縣</asp:ListItem>
                                <asp:ListItem>基隆市</asp:ListItem>
                                <asp:ListItem>宜蘭縣</asp:ListItem>
                                <asp:ListItem>花蓮縣</asp:ListItem>
                                <asp:ListItem>台東縣</asp:ListItem>
                                <asp:ListItem>金門地區</asp:ListItem>
                                <asp:ListItem>澎湖地區</asp:ListItem>
                                <asp:ListItem>馬祖地區</asp:ListItem>
                                <asp:ListItem>出國</asp:ListItem>
                            </asp:DropDownList></td>
                        <td style="height: 26px; width: 120px;" >&nbsp;&nbsp;
                            <asp:Label ID="Label20" runat="server" ForeColor="Black" Text="連絡電話：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 200px;" >
                            <asp:TextBox ID="TXT_nPHONE" runat="server"  MaxLength="20"></asp:TextBox></td>
                    </tr> 
                    <tr>
                        <td style="height: 26px; width: 120px;" >*
                            <asp:Label ID="Label13" runat="server" ForeColor="Black" Text="差勤事由：" CssClass="form"></asp:Label>
                        </td>
                        <td colspan=3 width="560">
                            <asp:TextBox ID="TXT_nREASON" runat="server" Height="59px" MaxLength="255" Rows="3"
                                TextMode="MultiLine" Width="560px"></asp:TextBox>
                        </td>
                    </tr> 
                    <% If read_only = "2" Then%>
                    <tr>
                        <td style="height: 26px; width: 120px;" >
                            &nbsp;&nbsp;
                            <asp:Label ID="Label28" runat="server" ForeColor="Black" Text="批核意見：" CssClass="form"></asp:Label><br />
                        </td>
                        <td colspan=3 width="600">
                            <asp:TextBox ID="txtcomment" runat="server" Height="59px" MaxLength="255" Rows="3"
                                TextMode="MultiLine" Width="510px"></asp:TextBox>
                            <asp:Button ID="But_PHRASE" runat="server" Text="批核片語" /></td>
                    </tr> 
                    <% End If%>
                    <tr>
                        <td colspan="2">

                        &nbsp;&nbsp;&nbsp;
                            <asp:Label ID="lbAidAndComfort" runat="server" ForeColor="Black" 
                                Text="本年度慰勞假日數：" CssClass="form"></asp:Label>

                        </td>
                    </tr>
            </table>   
             <table style="width: 720px; height: 37px; color:Red" >
                    <tr>
                        <td align="center" style="height: 29px; width: 720px;">
                        <asp:Button ID="but_exe" runat="server" Text="送件" style="height: 21px" />&nbsp;
                        <asp:Button ID="backBtn" runat="server" Text="駁回" />&nbsp;
                        <asp:Button ID="tranBtn" runat="server" Text="呈轉" />
                        <asp:Button ID="Button3" runat="server" Text="列印" />
                        </td>
                    </tr> 
             </table> 
            
            <asp:GridView ID="GridView1" runat="server" Width="720px" AutoGenerateColumns="False" BorderStyle="None" CssClass="form" BackColor="White" BorderColor="#3366CC" BorderWidth="1px" CellPadding="4">
                <Columns>
                    <asp:BoundField DataField="State_Name" HeaderText="假別" ReadOnly="True">
                        <ItemStyle HorizontalAlign="Left" />
                        <HeaderStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="nday" HeaderText="天" ReadOnly="True">
                        <ItemStyle HorizontalAlign="Right" />
                        <HeaderStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="nhour" HeaderText="時" ReadOnly="True">
                        <ItemStyle HorizontalAlign="Right" />
                        <HeaderStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                </Columns>
                <FooterStyle BackColor="#99CCCC" ForeColor="#003399" />
                <RowStyle BackColor="White" ForeColor="#003399" />
                <PagerStyle BackColor="#99CCCC" ForeColor="#003399" HorizontalAlign="Left" />
                <SelectedRowStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
                <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="#CCCCFF" />
            </asp:GridView>
            
            <uc1:FlowRoute ID="FlowRoute1" runat="server" />
            
            
            </td>
        </tr>
        
        
        </table>
        
		      
         </div>
         
        <div id="Div_grid" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:165pt; height:150pt; left:105px; top:592px; display:block;" visible="false">
         
            <asp:Calendar ID="Calendar1" runat="server" BackColor="White" BorderColor="#3366CC"
            BorderWidth="1px" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana"
            Font-Size="8pt" ForeColor="#003399" Height="200px" Width="220px" Caption="" ShowGridLines="True">
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
        
        <div id="Div_grid2" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:165pt; height:150pt; left:360px; top:592px; display:block;" visible="false">
        <asp:Calendar ID="Calendar2" runat="server" BackColor="White" BorderColor="#3366CC"
            BorderWidth="1px" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana"
            Font-Size="8pt" ForeColor="#003399" Height="200px" Width="220px" ShowGridLines="True">
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
            
             
        <div id="Div_grid10" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:300pt; height:80pt; left:249px; top:893px; display:block;" visible="false">
            <asp:GridView id="GridView2" runat="server" CssClass="form" Width="100%" Height="50px" DataSourceID="SqlDataSource2" PageSize="5" AutoGenerateColumns="False" AllowPaging="True" BorderColor="Lime" BorderWidth="2px">
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
      <asp:SqlDataSource id="SqlDataSource2" runat="server" SelectCommand="SELECT Phrase_num, employee_id, comment FROM PHRASE WHERE [employee_id] = @employee_id ORDER BY Phrase_num" ConnectionString="<%$ ConnectionStrings:ConnectionString %>">                      
          <SelectParameters>
              <asp:SessionParameter Name="employee_id" SessionField="user_id" Type="String" />
          </SelectParameters>
      </asp:SqlDataSource>
            </div>   
                 
    </form> 
    
    <iframe id="lst" frameborder=0 width=0 height=0 src="/blank.htm"></iframe>   
    <script language=javascript>
    
function window_onload() {
   <%if do_sql.G_errmsg <>"" then %>  
    alert('<%= do_sql.G_errmsg%>');
  <%end if 
  if date1_flag = True then %>
      run_div.style.left =<% =date1_x %> ;
      run_div.style.top =<% =date1_y %> ;   
      run_div.style.display="block";     
  <%elseif date2_flag = True then %>
      run_div2.style.left =<% =date2_x %> ;
      run_div2.style.top =<% =date2_y %> ;   
      run_div2.style.display="block";    
  <%end if 
      if print_file <>"" then %>
        lst.navigate("<%=print_file%>");              
     <%end if %>
  
}

    </script>
    
    
    

</body>
</html>
