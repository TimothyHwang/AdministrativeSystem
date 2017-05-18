<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04001.aspx.vb" Inherits="Source_04_MOA04001" %>

<%@ Register Src="../90/FlowRoute.ascx" TagName="FlowRoute" TagPrefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>房屋水電修繕申請單</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
     <%--   <script type="text/javascript" language="javascript">
            document.oncontextmenu = new Function("return false");
    </script>--%>
</head>
<body language="javascript" onload="return window_onload()" >
    <form id="form1" runat="server">
    <div>   
    
        <table width="750" border=1 cellspacing=0 cellpadding=5 bgcolor=#ffffff bordercolor=#6699cc bordercolorlight=#74a3d6 bordercolordark=#000000 style="LEFT: 20px; TOP: 10px">
        <tr>
            <td valign=bottom bgcolor=#6699cc bordercolorlight=#66aaaa bordercolordark=#ffffff><font color=white><b>&nbsp;房屋水電修繕申請單</b></font>
            </td>
        </tr>
        
        <tr>
            <td>
            
             <fieldset id="tableB" style="width: 750px">
                   <table style="width: 750px; height: 57px">
                    <tr>
                        <td style="width: 240px">
                            <asp:Label ID="Label1" runat="server" Text="填表人單位：" Width="80px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_PWUNIT" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label>
                        <td style="width: 220px">
                            <asp:Label ID="Label2" runat="server" Text="姓名：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_PWNAME" runat="server" ForeColor="Black" Width="106px" CssClass="form"></asp:Label></td>
                        <td style="width: 250px">
                            <asp:Label ID="Label3" runat="server" Text="級職：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_PWTITLE" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label></td>
                    </tr>
                    <tr>
                        <td style="width: 240px">
                            <asp:Label ID="Label4" runat="server" Text="申請人單位：" Width="80px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_PAUNIT" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label></td>
                        <td style="width: 220px">
                            <asp:Label ID="Label5" runat="server" Text="姓名：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>&nbsp;
                            <asp:DropDownList ID="DrDown_PANAME" runat="server" Width="143px" AutoPostBack="True">
                            </asp:DropDownList>
                        <td style="width: 250px">
                            <asp:Label ID="Label6" runat="server" Text="級職：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_PATITLE" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label></td>
                    </tr>                
                    </table>  
				</fieldset>
    			
			    <table border="0" style="width: 750px; height: 57px; color:Red" >
                    <tr>
                        <td style="width: 120px; height: 23px;">*
                            <asp:Label ID="Label7" runat="server" ForeColor="Black" Text="申請時間：" CssClass="form" ></asp:Label></td>
                        <td style="width: 550px; height: 23px;">
                            <asp:Label ID="Label8" runat="server" ForeColor="Black" CssClass="form" Width="310px" ></asp:Label></td>                    
                    </tr> 
                    <tr>
                        <td style="width: 120px">*
                            <asp:Label ID="Label9" runat="server" ForeColor="Black" Text="電話：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 550px;" >
                            <asp:TextBox ID="Txt_nPHONE" runat="server"  MaxLength="25" Width="160px"></asp:TextBox></td>                      
                    </tr> 
                    
                    <tr>
                        <td style="height: 26px; width: 120px;" >*
                            <asp:Label ID="Label12" runat="server" ForeColor="Black" Text="請修地點：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 550px;" >
                            <asp:TextBox ID="Txt_nPLACE" runat="server"  MaxLength="100" Width="340px"></asp:TextBox></td>
                    </tr> 
                    <tr>
                        <td style="height: 26px; width: 120px;" >*
                            <asp:Label ID="Label10" runat="server" ForeColor="Black" Text="請修事項：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 550px;" >
                            <asp:TextBox ID="Txt_nFIXITEM" runat="server"  MaxLength="100" Width="500px" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
                    </tr> 
                    
                    <% If read_only <> "" Or stepChk = "1" Then%>
                    <tr>
                        <td style="width: 120px" >
                            &nbsp;
                            <asp:Label ID="Label15" runat="server" ForeColor="Black" Text="承辦類別：" CssClass="form"></asp:Label></td>
                        <td style="width: 550px" >
                            <asp:RadioButtonList
                                ID="rdo_nExternal" runat="server" ForeColor="Black" RepeatDirection="Horizontal" Width="150px" CssClass="form">
                                <asp:ListItem Selected="True">內派</asp:ListItem>
                                <asp:ListItem>外包</asp:ListItem>
                            </asp:RadioButtonList></td>
                    </tr> 
                    <tr>
                        <td style="height: 26px; width: 120px;" >
                            &nbsp;
                            <asp:Label ID="Label16" runat="server" ForeColor="Black" Text="設施編號：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 550px;" >
                            <asp:TextBox ID="Txt_nFacilityNo" runat="server"  MaxLength="25" Width="150px" CssClass="form"></asp:TextBox></td>
                    </tr>
                    <% end if%> 
                    
                    <% If read_only <> "" Or stepChk = "2" Then%>
                    <tr>
                        <td style="height: 26px; width: 120px;" >
                            &nbsp;
                            <asp:Label ID="Label11" runat="server" ForeColor="Black" Text="維修時間：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 550px;" >
                            <asp:TextBox ID="Txt_nFIXDATE" runat="server" EnableTheming="False" Width="89px" OnKeyDown="return false"></asp:TextBox>
                            <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" /></td>
                    </tr> 
                    
                    <tr>
                        <td style="height: 26px; width: 120px;" >
                            &nbsp;
                            <asp:Label ID="Label13" runat="server" ForeColor="Black" Text="派員人數：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 550px;" >
                            <asp:TextBox ID="Txt_nPacthCount" runat="server"  MaxLength="4" Width="85px" CssClass="form"></asp:TextBox></td>
                    </tr> 
                    <tr>
                        <td style="height: 26px; width: 120px;" >
                            &nbsp;
                            <asp:Label ID="Label14" runat="server" ForeColor="Black" Text="施工人員：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 550px;" >
                            <asp:TextBox ID="Txt_nPacthPer" runat="server"  MaxLength="20" Width="214px"  CssClass="form"></asp:TextBox></td>
                    </tr> 
                    
                    
                    <% End If%>
                    
                    <% If read_only = "2" Then%>
                    <tr>
                        <td style="height: 26px; width: 120px;" >
                            &nbsp;<asp:Label ID="Label28" runat="server" ForeColor="Black" Text="批核意見：" CssClass="form"></asp:Label><br />
                        </td>
                        <td colspan=3 >
                            <asp:TextBox ID="txtcomment" runat="server" Height="59px" MaxLength="255" Rows="3"
                                TextMode="MultiLine" Width="529px"></asp:TextBox>
                            <asp:Button ID="But_PHRASE" runat="server" Text="批核片語" /></td>
                    </tr> 
                    <% End If%>
                    
                    </table>  
                    <table border="0" style="width: 750px; height: 57px">
                        <tr>                  
                        <td align="center">
                        <asp:Button ID="But_exe" runat="server" Text="送件" />
                        <asp:Button ID="But_prt" runat="server" Text="列印" />&nbsp;<asp:Button ID="backBtn"
                            runat="server" Text="駁回" />&nbsp;<asp:Button ID="tranBtn" runat="server" Text="呈轉" /></td>                    
                        </tr> 
                    </table> 
            <uc1:FlowRoute ID="FlowRoute1" runat="server" />
            
            
            
            </td>
        </tr>
        </table>
    
    
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="">
        </asp:SqlDataSource>
        &nbsp;
    </div>
    
        <div id="Div_grid" runat="server" style="border-right: lightslategray 2px solid;
            border-top: lightslategray 2px solid; display: block; z-index: 3; left: 229px;
            border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
            position: absolute; top: 573px; height: 150pt; background-color: white" visible="false">
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
            
            <div id="Div_grid10" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:300pt; height:80pt; left:249px; top:893px; display:block;" visible="false">
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
    
    <script language=javascript>
    
function window_onload() {
  <%if do_sql.G_errmsg <>"" then %>  
    alert('<%= do_sql.G_errmsg %>');
  <%end if %>
  
  //顯示報表檔
  <%if pfilename <>"" then %>  
    window.location.replace('<%= pfilename %>');
  <%end if %>
  
  //產生列印檔過程有錯誤
  <%if ErrMsg <>"" then %>  
    alert('<%= ErrMsg %>');
  <%end if %>
}

    </script>
</body>
</html>
