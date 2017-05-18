<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA03001.aspx.vb" Inherits="Source_03_MOA03001" %>

<%@ Register Src="../90/FlowRoute.ascx" TagName="FlowRoute" TagPrefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>派車申請單</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
     <script type="text/javascript" language="javascript">
         document.oncontextmenu = new Function("return false");
    </script>
</head>
<body language="javascript" onload="return window_onload()">
    <form id="form1" runat="server">
    <div>
        
        <table width="730" border=1 cellspacing=0 cellpadding=5 bgcolor=#ffffff bordercolor=#6699cc bordercolorlight=#74a3d6 bordercolordark=#000000 style="LEFT: 10px; TOP: 10px">
        <tr>
            <td valign=bottom bgcolor=#6699cc bordercolorlight=#66aaaa bordercolordark=#ffffff style="width: 734px"><font color=white><b>&nbsp;派車申請單</b></font>
            </td>
        </tr>
        
        <tr>
            <td style="width: 734px">           
            
            <fieldset id="tableB" style="width: 720px">
                <table style="width: 710px; height: 57px">
                    <tr>
                        <td style="width: 240px">
                            <asp:Label ID="Label1" runat="server" Text="填表人單位：" Width="80px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_PWUNIT" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label>
                        <td style="width: 210px">
                            <asp:Label ID="Label2" runat="server" Text="姓名：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_PWNAME" runat="server" ForeColor="Black" Width="106px" CssClass="form"></asp:Label></td>
                        <td style="width: 230px">
                            <asp:Label ID="Label3" runat="server" Text="級職：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_PWTITLE" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label></td>
                    </tr>
                    <tr>
                        <td style="width: 240px">
                            <asp:Label ID="Label4" runat="server" Text="申請人單位：" Width="80px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_PAUNIT" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label></td>
                        <td style="width: 210px">
                            <asp:Label ID="Label5" runat="server" Text="姓名：" Width="51px" ForeColor="Black" CssClass="form"></asp:Label>&nbsp;<asp:DropDownList ID="DrDown_PANAME" runat="server" Width="143px" AutoPostBack="True">
                            </asp:DropDownList>
                        <td style="width: 230px">
                            <asp:Label ID="Label6" runat="server" Text="級職：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_PATITLE" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label></td>
                    </tr>                
                </table>  
             </fieldset> 
             
			    <table border="0" style="width: 720px; height: 57px; color:Red" >
                    <tr>
                        <td style="width: 120px; height: 12px;">&nbsp;&nbsp;
                        <asp:Label ID="Label15" runat="server" ForeColor="Black" Text="申請時間：" CssClass="form" ></asp:Label></td>
                        <td style="width: 240px; height: 12px;">
                            <asp:Label ID="Lab_time" runat="server" ForeColor="Black" CssClass="form" Width="190px" ></asp:Label></td>
                        <td style="width: 100px; height: 12px;">*
                            <asp:Label ID="Label7" runat="server" ForeColor="Black" Text="聯絡電話：" CssClass="form" ></asp:Label></td>
                        <td style="height: 12px; width: 240px;" >
                            <asp:TextBox ID="TXT_nPHONE" runat="server"  MaxLength="20" Width="140px"></asp:TextBox></td>
                    </tr>                             
                   
                    <tr>
                        <td style="height: 26px; width: 120px;" >*
                            <asp:Label ID="Label12" runat="server" ForeColor="Black" Text="任務理由：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 240px;" >
                            <asp:TextBox ID="TXT_nREASON" runat="server"  MaxLength="255" Width="209px"></asp:TextBox></td>
                        <td style="height: 26px; width: 100px;" >*
                            <asp:Label ID="Label20" runat="server" ForeColor="Black" Text="人員項目：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 240px;" >
                            <asp:TextBox ID="TXT_nITEM" runat="server"  MaxLength="3" Width="50px"></asp:TextBox>&nbsp;<asp:Label 
                                ID="Label32" runat="server" CssClass="form" ForeColor="Black" Text="員"></asp:Label>
                        </td>
                    </tr> 
                    <tr>
                        <td style="height: 26px; width: 120px;" >*
                            <asp:Label ID="Label8" runat="server" ForeColor="Black" Text="車輛報到地點：" CssClass="form" Width="100px"></asp:Label></td>
                        <td style="height: 26px; width: 240px;" >
                            <asp:TextBox ID="TXT_nARRIVEPLACE" runat="server"  MaxLength="50" Width="180px"></asp:TextBox></td>
                        <td style="height: 26px; width: 100px;" >*
                            <asp:Label ID="Label9" runat="server" ForeColor="Black" Text="向何人報到：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 240px;" >
                            <asp:TextBox ID="TXT_nARRIVETO" runat="server"  MaxLength="50" Width="150px"></asp:TextBox></td>
                    </tr>
                    <tr>
                        <td style="height: 23px; width: 120px;">*
                            <asp:Label ID="Label16" runat="server" ForeColor="Black" Text="車輛報到日期：" CssClass="form" ></asp:Label></td>
                        <td colspan=3 style="height: 23px; width: 580px;" >
                            <asp:TextBox ID="Txt_nARRDATE" runat="server" EnableTheming="False" OnKeyDown="return false" CausesValidation="True" Width="89px"></asp:TextBox>
                            <asp:ImageButton ID="ImgCalender" runat="server" ImageUrl="~/Image/calendar.gif" />
                            <asp:DropDownList ID="DrDown_nSTHOUR" runat="server">
                            </asp:DropDownList><asp:Label ID="Label17" runat="server" ForeColor="Black" Text="時" CssClass="form" Width="20px"></asp:Label><asp:DropDownList ID="DrDown_nEDHOUR" runat="server">
                            </asp:DropDownList><asp:Label ID="Label18" runat="server" ForeColor="Black" Text="分" CssClass="form" Width="20px"></asp:Label>
                        </td>
                    </tr> 
                    <tr>
                        <td style="height: 26px; width: 120px;" >*
                            <asp:Label ID="Label13" runat="server" ForeColor="Black" Text="起點：" CssClass="form" Width="80px"></asp:Label></td>
                        <td style="height: 26px; width: 240px;" >
                            <asp:TextBox ID="Txt_nSTARTPOINT" runat="server"  MaxLength="50" Width="180px"></asp:TextBox></td>
                        <td style="height: 26px; width: 100px;" >*
                            <asp:Label ID="Label14" runat="server" ForeColor="Black" Text="目的地：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 240px;" >
                            <asp:TextBox ID="Txt_nENDPOINT" runat="server"  MaxLength="50" Width="180px"></asp:TextBox></td>
                    </tr> 
                </table>   
            
            <fieldset id="Fieldset1" style="width: 720px">
                <legend class="form"><font color="red">車輛資訊</font></legend>
                <table border="0" style="width: 720px; height: 23px;color:Red">
                   <tr>
                    <td style="width: 110px; height: 23px;">&nbsp;&nbsp;
                        <asp:Label ID="Label19" runat="server" ForeColor="Black" Text="車輛類型：" CssClass="form" ></asp:Label></td>
                    <td style="width: 560px; height: 23px;">
                        &nbsp;<asp:RadioButton ID="Radio1" runat="server" ForeColor="Black" Text="一般車"  Checked="True" GroupName="rad" CssClass="form" AutoPostBack="True" />
                        <asp:RadioButton ID="Radio2" runat="server" ForeColor="Black" Text="經常性支援"  GroupName="rad" CssClass="form" AutoPostBack="True" />
                        <asp:RadioButton ID="Radio3" runat="server" ForeColor="Black" Text="主官管"  GroupName="rad" CssClass="form" AutoPostBack="True" /></td>                  
                  </tr> 
                  <tr>
                    <td style="width: 110px; height: 23px;">*
                        <asp:Label ID="Label21" runat="server" ForeColor="Black" Text="車輛品名型式：" CssClass="form" ></asp:Label></td>
                    <td style="width: 560px; height: 23px;">
                        &nbsp;<asp:DropDownList ID="DrDown_nSTYLE" runat="server">
                        </asp:DropDownList>
                        <asp:Button ID="Button2" runat="server" Text="查詢車輛數"/>
                        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;<asp:Label ID="LabCarTitle" runat="server"
                            CssClass="form" ForeColor="Black" Text="車號：" Visible="False" Width="70px"></asp:Label>
                        <asp:Label ID="LabCarNum" runat="server" CssClass="form" ForeColor="Black" Visible="False"
                            Width="180px"></asp:Label></td>                  
                  </tr> 
                  <tr>
                    <td style="height: 22px; width: 110px;">*
                        <asp:Label ID="Label22" runat="server" ForeColor="Black" Text="任務使用時間：" CssClass="form" ></asp:Label></td>
                    <td colspan=3 style="height: 22px; width: 560px;" >
                        <asp:TextBox ID="Txt_nUSEDATE" runat="server" EnableTheming="False" OnKeyDown="return false" CausesValidation="True" Width="70px"></asp:TextBox><asp:Label
                            ID="Label29" runat="server" CssClass="form" ForeColor="Black" Text="日" Width="20px"></asp:Label><asp:DropDownList ID="DrDown_nSTUSEHOUR" runat="server">
                        </asp:DropDownList><asp:Label ID="Label23" runat="server" ForeColor="Black" Text="時" CssClass="form" Width="20px"></asp:Label><asp:DropDownList ID="DrDown_nSTUSEMIN" runat="server">
                        </asp:DropDownList><asp:Label ID="Label24" runat="server" ForeColor="Black" Text="分至" CssClass="form" Width="30px"></asp:Label><asp:TextBox ID="Txt_nEDUSEDATE" runat="server" EnableTheming="False" OnKeyDown="return false" CausesValidation="True" Width="70px"></asp:TextBox><asp:Label
                            ID="Label30" runat="server" CssClass="form" ForeColor="Black" Text="日" Width="20px"></asp:Label><asp:DropDownList ID="DrDown_nEDUSEHOUR" runat="server">
                        </asp:DropDownList><asp:Label ID="Label25" runat="server" ForeColor="Black" Text="時" CssClass="form" Width="20px"></asp:Label><asp:DropDownList ID="DrDown_nEDUSEMIN" runat="server">
                        </asp:DropDownList><asp:Label ID="Label26" runat="server" ForeColor="Black" Text="止" CssClass="form" Width="20px"></asp:Label></td>                   
                </tr> 
                </table>
            </fieldset>
            
                <table border="0" style="width: 720px; height: 23px">
                       <tr>
                        <td style="height: 26px; width: 80px;" >&nbsp;&nbsp;
                            <asp:Label ID="Label27" runat="server" ForeColor="Black" Text="從：" CssClass="form" Width="30px"></asp:Label></td>
                        <td style="height: 26px; width: 280px;" >
                            <asp:TextBox ID="Txt_GoLocal" runat="server"  MaxLength="50" Width="190px"></asp:TextBox></td>
                        <td style="height: 26px; width: 80px;" >
                            &nbsp; &nbsp;
                            <asp:Label ID="Label28" runat="server" ForeColor="Black" Text="至：" CssClass="form" Width="30px"></asp:Label></td>
                        <td style="height: 26px; width: 280px;" >
                            <asp:TextBox ID="Txt_EndLocal" runat="server"  MaxLength="50" Width="189px"></asp:TextBox>
                            <asp:Button ID="But_ins" runat="server" Text="新增" /></td>
                    </tr>                 
                </table> 
                <asp:GridView ID="GridView1" runat="server" Width="720px" AutoGenerateColumns="False" BorderStyle="Double" CssClass="form" DataSourceID="SqlDataSource1" DataKeyNames="Local_Num" Height="50px" ForeColor="Black">
                <Columns>
                    <asp:BoundField DataField="Local_Num" HeaderText="Local_Num" InsertVisible="False"
                        ReadOnly="True" SortExpression="Local_Num" Visible="False" />
                    <asp:BoundField DataField="EFORMSN" HeaderText="EFORMSN" SortExpression="EFORMSN"
                        Visible="False">
                        <HeaderStyle Width="30px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="GoLocal" HeaderText="起點" SortExpression="GoLocal" />
                    <asp:BoundField DataField="EndLocal" HeaderText="終點" SortExpression="EndLocal" />
                    <asp:CommandField ShowDeleteButton="True">
                        <HeaderStyle Height="15px" />
                        <ItemStyle Width="50px" />
                    </asp:CommandField>
                </Columns>
                </asp:GridView>             
            
                <table border="0" style="width: 720; height: 23px">
         
                    <% If read_only = "2" Then%>
                    <tr>
                        <td style="height: 26px; width: 120px;" >
                            &nbsp;&nbsp;
                            <asp:Label ID="Label31" runat="server" ForeColor="Black" Text="批核意見：" CssClass="form" Width="80px"></asp:Label><br />
                        </td>
                        <td style="width: 600px">
                            <asp:TextBox ID="txtcomment" runat="server" Height="59px" MaxLength="255" Rows="3"
                                TextMode="MultiLine" Width="510px"></asp:TextBox>
                            <asp:Button ID="But_PHRASE" runat="server" Text="批核片語" /></td>
                    </tr> 
                    <% End If%>
                    
                    <tr>
                       <td align="center" colspan=2  style="width: 720px">
                           <asp:Button ID="But_exe" runat="server" Text="送件" />
                           <asp:Button ID="backBtn" runat="server" Text="駁回" />
                           <asp:Button ID="tranBtn" runat="server" Text="呈轉" />
                           <asp:Button ID="supplyBtn" runat="server" Text="補登" />&nbsp;
                       <% If strDisplayPrintButton.Equals("YES") Then%>
                           <input id="btnPrint" onclick="javascript:window.open('MOA03011.aspx?EFORMSN=<%=eformsn %>');" type="button" value="列印" />
                           <asp:ImageButton ID="ImagePrint" runat="server" ImageUrl="~/Image/print.gif" ToolTip="列印" Visible="False" />
                       <% End If %>
                       </td>
                    </tr> 
                </table>
        <uc1:FlowRoute ID="FlowRoute1" runat="server" />
            
            
            </td>
        </tr>
        </table>
    
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [P_0305] WHERE [Local_Num] = @Local_Num" InsertCommand="INSERT INTO [P_0305] ([EFORMSN], [GoLocal], [EndLocal]) VALUES (@EFORMSN, @GoLocal, @EndLocal)"
            SelectCommand="SELECT [Local_Num], [EFORMSN], [GoLocal], [EndLocal] FROM [P_0305] WHERE ([EFORMSN] = @EFORMSN)"
            UpdateCommand="UPDATE [P_0305] SET [EFORMSN] = @EFORMSN, [GoLocal] = @GoLocal, [EndLocal] = @EndLocal WHERE [Local_Num] = @Local_Num">
            <DeleteParameters>
                <asp:Parameter Name="Local_Num" Type="Decimal" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="EFORMSN" Type="String" />
                <asp:Parameter Name="GoLocal" Type="String" />
                <asp:Parameter Name="EndLocal" Type="String" />
                <asp:Parameter Name="Local_Num" Type="Decimal" />
            </UpdateParameters>
            <SelectParameters>
                <asp:QueryStringParameter Name="EFORMSN" QueryStringField="eformsn" Type="String" />
            </SelectParameters>
            <InsertParameters>
                <asp:Parameter Name="EFORMSN" Type="String" />
                <asp:Parameter Name="GoLocal" Type="String" />
                <asp:Parameter Name="EndLocal" Type="String" />
            </InsertParameters>
        </asp:SqlDataSource>
        &nbsp; &nbsp;
        &nbsp; &nbsp;
    </div>
    
        <div id="Div_grid2" runat="server" style="position:absolute ; z-index:3; border:2 solid lightslategray; background-color:white; width:161pt; height:150pt; left:101px; top:698px; display:block;" visible="false">
        <asp:Calendar ID="Calendar1" runat="server" BackColor="White" BorderColor="#3366CC"
            BorderWidth="1px" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana"
            Font-Size="8pt" ForeColor="#003399" Height="200px" Width="220px" CssClass="form" ShowGridLines="True">
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
    
    <div id="Div_grid" runat="server" style="position:absolute;z-index:99; border:2 solid lightslategray; background-color:white; width:225pt; height:193pt; left:366px; top:700px; display:block;" visible="false">
        <asp:GridView ID="GridView2" runat="server" Height="229px" Width="302px" AutoGenerateColumns="False" AllowPaging="True" CssClass="form" PageSize="8" DataSourceID="SqlDataSource2">
            <Columns>
                <asp:BoundField DataField="pck_name" HeaderText="車輛名稱">
                    <ItemStyle Width="220px" />
                    <HeaderStyle CssClass="form" />
                </asp:BoundField>
                <asp:BoundField DataField="num_car" HeaderText="可用數量" ReadOnly="True">
                    <ItemStyle Width="80px" HorizontalAlign="Right" />
                    <HeaderStyle CssClass="form" />
                </asp:BoundField>
            </Columns>
            <EditRowStyle Height="10px" />
            <RowStyle Height="10px" />
            <HeaderStyle Height="10px" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [PCK_Name], [PCN_Use] FROM [P_0306] WHERE ([PCN_Date] = @PCN_Date)">
            <SelectParameters>
                <asp:Parameter DefaultValue="2008/01/01" Name="PCN_Date" Type="DateTime" />
            </SelectParameters>
        </asp:SqlDataSource>
        <asp:Button ID="But_cencel" runat="server" Text="關閉" Width="299px"/></div>     
        
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
    <iframe id="lst" frameborder=0 width=0 height=0 src="/blank.htm"></iframe>   
   <script language=javascript>
   var print_file='<%= print_file%>';
function window_onload() {
    <%if do_sql.G_errmsg <>"" then %>  
         alert('<%= do_sql.G_errmsg.Replace("'", "") %>');
  <%end if 
  if date1_flag = True then %>
      run_div.style.left =<% =date1_x %> ;
      run_div.style.top =<% =date1_y %> ;   
      run_div.style.display="block";      
  <%end if %>
  if (print_file != '') 
  { lst.navigate(print_file); }
}
   </script>
</body>
</html>
