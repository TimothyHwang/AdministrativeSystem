<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA05001.aspx.vb" Inherits="Source_05_MOA05001" %>

<%@ Register Src="../90/FlowRoute.ascx" TagName="FlowRoute" TagPrefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>會客洽公申請單</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.7.2.min.js" type="text/javascript"></script>
    <script type="text/javascript" language="javascript">
        document.oncontextmenu = new Function("return false");
    </script>
    <link href="../../css/jquery.datepick.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.10.2.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick-zh-TW.js" type="text/javascript"></script>
    <script type="text/javascript" language="javascript">
        $(function () {
            $("#Txt_nRECDATE").datepick({ formats: 'yyyy/m/d', defaultDate: $("#Txt_nRECDATE").val(), showTrigger: '#calImg' });                        
        });        
    </script>
</head>
<body language="javascript" onload="return window_onload()" >
    <form id="form1" runat="server">
    <div>
    
        <table width="750" border=1 cellspacing=0 cellpadding=5 bgcolor=#ffffff bordercolor=#6699cc bordercolorlight=#74a3d6 bordercolordark=#000000 style="LEFT: 20px; TOP: 10px">
        <tr>
            <td valign=bottom bgcolor=#6699cc bordercolorlight=#66aaaa bordercolordark=#ffffff><font color=white><b>&nbsp;會客洽公申請單</b></font>
            </td>
        </tr>
        
        <tr>
            <td>            
            
             <fieldset id="tableB" style="width: 750px">
                   <table style="width: 750px; height: 57px">
                    <tr>
                        <td style="width: 249px; height: 14px;">
                            <asp:Label ID="Label1" runat="server" Text="填表人單位：" Width="80px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_ORG_NAME_1" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label>
                        <td style="width: 240px; height: 14px;">
                            <asp:Label ID="Label2" runat="server" Text="姓名：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_emp_chinese_name" runat="server" ForeColor="Black" Width="106px" CssClass="form"></asp:Label></td>
                        <td style="height: 14px; width: 240px;">
                            <asp:Label ID="Label3" runat="server" Text="級職：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_title_name_1" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label></td>
                    </tr>
                    <tr>
                        <td style="width: 249px; height: 21px;">
                            <asp:Label ID="Label4" runat="server" Text="申請人單位：" Width="80px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_ORG_NAME_2" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label></td>
                        <td style="width: 240px; height: 21px;">
                            <asp:Label ID="Label5" runat="server" Text="姓名：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>&nbsp;
                            <asp:DropDownList ID="DrDown_emp_chinese_name" runat="server" Width="143px" AutoPostBack="True" CssClass="form" ForeColor="Black">
                            </asp:DropDownList>
                        <td style="height: 21px; width: 240px;">
                            <asp:Label ID="Label6" runat="server" Text="級職：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_title_name_2" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label></td>
                    </tr>                
                </table>
                 <asp:Label ID="Lab_eformsn" runat="server" CssClass="form" ForeColor="Black" Text="申請時間："
                     Visible="False"></asp:Label>
                 </fieldset> 
                 
			    <table border="0" style="width: 750px; height: 57px; color:Red" >
                    <tr>
                        <td style="width: 90px; height: 15px;">*
                            <asp:Label ID="Label7" runat="server" ForeColor="Black" Text="申請時間：" CssClass="form" ></asp:Label></td>
                        <td style="width: 160px; height: 15px;">
                            <asp:Label ID="Label8" runat="server" ForeColor="Black" CssClass="form" Width="151px" ></asp:Label>
                 </td>   
                        <td style="width: 80px; height: 15px;">*
                            <asp:Label ID="Label9" runat="server" ForeColor="Black" Text="會客室：" CssClass="form"></asp:Label></td>
                        <td style="width: 400px; height: 15px;"  >
                            <asp:DropDownList ID="DrDown_nRECROOM" runat="server" AutoPostBack="True" CssClass="form" Width="150px" ForeColor="Black">
                            </asp:DropDownList></td>                   
                    </tr>                 
                    
                    <tr>
                        <td style="width: 90px; height: 8px;">
                            <asp:Label ID="Label10" runat="server" ForeColor="Black" Text="會客入口：" CssClass="form"></asp:Label></td>
                        <td style="width: 160px; height: 8px;"  >
                            <asp:DropDownList ID="DrDown_nRECEXIT" runat="server" AutoPostBack="True" 
                                CssClass="form" Width="110px" ForeColor="Black" Enabled="False">
                            </asp:DropDownList></td>  
                        <td style="height: 8px; width: 80px;" >*
                            <asp:Label ID="Label12" runat="server" ForeColor="Black" Text="地點：" CssClass="form"></asp:Label></td>
                        <td style="height: 8px; width: 400px;" >
                            <asp:TextBox ID="TXT_nPLACE" runat="server"  MaxLength="50" Width="220px"></asp:TextBox></td>                      
                    </tr> 
                    
                    <tr>
                        <td style="height: 5px; width: 90px;" >*
                            <asp:Label ID="Label11" runat="server" ForeColor="Black" Text="電話：" CssClass="form"></asp:Label></td>
                        <td style="height: 5px; width: 160px;" >
                            <asp:TextBox ID="Txt_nPHONE" runat="server"  MaxLength="50" Width="101px"></asp:TextBox></td>
                        <td style="height: 5px; width: 80px;">*
                            <asp:Label ID="Label14" runat="server" ForeColor="Black" Text="時間：" CssClass="form" ></asp:Label></td>
                        <td style="height: 5px; width: 400px;" >
                            <asp:TextBox ID="Txt_nRECDATE" runat="server" EnableTheming="False" CausesValidation="True" Enabled="True" Width="70px" OnKeyDown="return false"></asp:TextBox>
                            <div style="display: none;">
                                <img id="calImg" src="../../Image/calendar.gif" alt="選擇日期" />
                            </div>
                            &nbsp;<asp:DropDownList ID="DrDown_nSTARTTIME" runat="server" AutoPostBack="True" CssClass="form" ForeColor="Black"></asp:DropDownList>
                            <asp:Label ID="Label22" runat="server" CssClass="form" ForeColor="Black" Text="時"></asp:Label>
                            <asp:DropDownList ID="DrDown_nSMin" runat="server" CssClass="form" ForeColor="Black">
                                        <asp:ListItem>00</asp:ListItem>
                                        <asp:ListItem>30</asp:ListItem>
                                    </asp:DropDownList><asp:Label ID="Label15" runat="server" ForeColor="Black" Text="分至" CssClass="form"></asp:Label><asp:DropDownList ID="DrDown_nENDTIME" runat="server" AutoPostBack="True" CssClass="form" ForeColor="Black">
                            </asp:DropDownList><asp:Label ID="Label23" runat="server" ForeColor="Black" Text="時" CssClass="form"></asp:Label><asp:DropDownList
                                ID="DrDown_nEMin" runat="server" CssClass="form" ForeColor="Black">
                                <asp:ListItem>00</asp:ListItem>
                                <asp:ListItem>30</asp:ListItem>
                            </asp:DropDownList>
                            <asp:Label ID="Label16" runat="server" CssClass="form" ForeColor="Black" Text="分止"></asp:Label>
                        </td>
                    </tr> 
                    
                    <tr>
                        <td rowspan=4; style="height: 35px; width: 90px;" 141px">
                            <br />
                            *
                            <asp:Label ID="Label13" runat="server" ForeColor="Black" Text="事由：" CssClass="form"></asp:Label><br />
                        </td>
                        <td rowspan=3; colspan=3 style="height: 32px">
                            <asp:TextBox ID="TXT_nREASON" runat="server" Height="46px" MaxLength="255" Rows="3" TextMode="MultiLine" Width="550px"></asp:TextBox></td>
                        
                    </tr> 
            </table>            
            
            </td>
        </tr>       
        
        </table>
    
    
        
    <div id="Div_grid" runat="server" 
            style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:300pt; height:150pt; left:774px; top:907px; display:block;" 
            visible="false">
    
        <asp:GridView ID="GridView2" runat="server" AllowPaging="True" AutoGenerateColumns="False"
          CssClass="form" Height="50px" PageSize="5" Width="100%" DataSourceID="SqlDataSource2">
          <Columns>
              <asp:BoundField DataField="nAPPTIME" HeaderText="申請時間" ReadOnly="True" >
                  <HeaderStyle HorizontalAlign="Center" Width="50%" />
              </asp:BoundField>
              <asp:BoundField DataField="nName" HeaderText="會客姓名" >
                  <HeaderStyle HorizontalAlign="Center" Width="20%" />
              </asp:BoundField>
              <asp:BoundField DataField="nREASON" HeaderText="事由" >
                  <HeaderStyle HorizontalAlign="Center" Width="20%" />
              </asp:BoundField>
              <asp:BoundField DataField="EFORMSN" HeaderText="EFORMSN" >
                  <ItemStyle Wrap="False" />
              </asp:BoundField>
              <asp:CommandField ShowSelectButton="True">
                  <HeaderStyle HorizontalAlign="Center" Width="10%" />
              </asp:CommandField>
          </Columns>
          <RowStyle Height="10px" />
      </asp:GridView>
        <asp:Button ID="CloseVistor" runat="server" Text="關閉" Width="100%" />
      <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
          SelectCommand="SELECT Top 10 *,'' As nName FROM [P_05] WHERE (DATEDIFF(day, nAPPTIME, GETDATE()) < 90) and ([PAIDNO] = @PAIDNO) and p_05.eformsn in (select p_0501.eformsn from p_0501 where P_0501.nKIND IS NOT NULL) ORDER BY [nAPPTIME] DESC">                      
          <SelectParameters>
              <asp:ControlParameter ControlID="DrDown_emp_chinese_name" Name="PAIDNO" PropertyName="SelectedValue"
                  Type="String" />
          </SelectParameters>
      </asp:SqlDataSource>    
        
    </div>  
                  
        <table border="0" style="width: 750px;">
          <tr>
              <td style="width: 620px">
              <table border="0" style="width: 750px; height: 57px">
                   <tr>
                    <td style="width: 100px; height: 17px;">
                        <asp:Label ID="Label17" runat="server" ForeColor="Black" Text="新增來賓資料：" CssClass="form" ></asp:Label></td>
                    <td colspan=4; style="width: 520px; height: 17px;">
                        <asp:DropDownList ID="DrDown_guest" runat="server" AutoPostBack="True">
                        </asp:DropDownList>
                        <asp:Button ID="SelOldEform" runat="server" Text="表單快速新增" /></td>                    
                  </tr> 
                  <tr>
                    <td style="width: 100px; height: 19px;">
                        <asp:Label ID="Label18" runat="server" ForeColor="Black" Text="姓名：" CssClass="form" ></asp:Label></td>
                    <td style="width: 80px; height: 19px;">
                        <asp:Label ID="Label19" runat="server" ForeColor="Black" Text="姓別：" CssClass="form" ></asp:Label></td>  
                    <td style="width: 50px; height: 19px;">
                        <asp:Label ID="Label24" runat="server" ForeColor="Black" Text="類別：" CssClass="form" ></asp:Label></td>
                    <td style="width: 200px; height: 19px;">
                        <asp:Label ID="Label20" runat="server" ForeColor="Black" Text="服務單位：" CssClass="form" ></asp:Label></td>                     
                    <td style="width: 150px; height: 19px;">
                        <asp:Label ID="Label21" runat="server" ForeColor="Black" Text="身份證字號：" CssClass="form" ></asp:Label><br />
                        <asp:RadioButton ID="RadioButton1" runat="server" Checked="True"
                            CssClass="form" ForeColor="Black" GroupName="id" Text="本國" AutoPostBack="True" /><asp:RadioButton ID="RadioButton2"
                                runat="server" CssClass="form" ForeColor="Black" GroupName="id"
                                Text="外國" AutoPostBack="True" /></td>   
                    <td style="width:100px;"><asp:Label ID="Label25" runat="server" ForeColor="Black" Text="車號：" CssClass="form" ></asp:Label></td>
                                    
                  </tr> 
                  <tr>
                    <td style="width: 100px; height: 18px;">
                        <asp:TextBox ID="Txt_nName" runat="server"  MaxLength="30" Width="90px"></asp:TextBox></td>
                    <td style="width: 80px; height: 18px;">
                        <asp:RadioButton ID="Radio1" runat="server" ForeColor="Black" Text="男"  Checked="True" GroupName="rad" CssClass="form" />
                        <asp:RadioButton ID="Radio2" runat="server" ForeColor="Black" Text="女"  Checked="False" GroupName="rad" CssClass="form" /></td>                    
                     <td style="width: 50px; height: 19px;">
                         <asp:DropDownList ID="DrDown_nkind" runat="server">
                             <asp:ListItem>民</asp:ListItem>
                             <asp:ListItem>軍</asp:ListItem>
                         </asp:DropDownList></td>
                     <td style="width: 200px; height: 18px;">
                        <asp:TextBox ID="Txt_nService" runat="server"  MaxLength="100" Width=180px></asp:TextBox></td>                      
                     <td style="width: 120px; height: 18px;">
                        <asp:TextBox ID="Txt_nID" runat="server"  MaxLength="10" Width="120px"></asp:TextBox></td>                         
                     <td style="width: 150px; height: 18px;">
                         <asp:TextBox ID="txtCarNo" runat="server"  MaxLength="10" Width="60px"></asp:TextBox>
                         <asp:Button ID="But_ins" runat="server" Text="新增" CssClass="form" ForeColor="Black" /></td>                  
                  </tr> 
              </table> 
              <asp:GridView ID="GridView1" runat="server" Width="730px" 
                      AutoGenerateColumns="False" BorderStyle="Double" CssClass="form" 
                      DataSourceID="SqlDataSource1" DataKeyNames="Receive_Num" AllowPaging="True" 
                      EnableModelValidation="True">
            <Columns>
                <asp:BoundField DataField="Receive_Num" HeaderText="Receive_Num" InsertVisible="False"
                    ReadOnly="True" SortExpression="Receive_Num" Visible="False" />
                <asp:BoundField DataField="EFORMSN" HeaderText="EFORMSN" SortExpression="EFORMSN"
                    Visible="False" />
                <asp:BoundField DataField="nNAME" HeaderText="姓名" SortExpression="nNAME">
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="nSEX" HeaderText="性別" SortExpression="nSEX">
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="20%" />
                </asp:BoundField>
                <asp:BoundField DataField="nSERVICE" HeaderText="服務單位" SortExpression="nSERVICE">
                    <HeaderStyle Width="30%" />
                </asp:BoundField>
                <asp:BoundField DataField="nKIND" HeaderText="類別" SortExpression="nKIND">
                    <HeaderStyle Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="nID" HeaderText="身份證字號" SortExpression="nID">
                    <HeaderStyle Width="20%" />
                </asp:BoundField>
                <asp:BoundField DataField="nCarNo" HeaderText="車號" SortExpression="nCarNo" />
                <asp:CommandField ShowDeleteButton="True">
                    <HeaderStyle Width="5%" />
                </asp:CommandField>
            </Columns>
                  <RowStyle Width="10px" />
        </asp:GridView>  
                  <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                      DeleteCommand="DELETE FROM [P_0501] WHERE [Receive_Num] = @Receive_Num" InsertCommand="INSERT INTO [P_0501] ([EFORMSN], [nNAME], [nSEX], [nSERVICE], [nKIND], [nID]) VALUES (@EFORMSN, @nNAME, @nSEX, @nSERVICE,@nKIND, @nID)"
                      SelectCommand="SELECT Receive_Num, EFORMSN, nNAME, nSEX, nSERVICE, nKIND, nID, nCarNo FROM P_0501 WHERE (EFORMSN = @EFORMSN)"
                      
                      
                      UpdateCommand="UPDATE [P_0501] SET [EFORMSN] = @EFORMSN, [nNAME] = @nNAME, [nSEX] = @nSEX, [nSERVICE] = @nSERVICE,[nKIND]=@nKIND, [nID] = @nID WHERE [Receive_Num] = @Receive_Num">
                      <DeleteParameters>
                          <asp:Parameter Name="Receive_Num" Type="Decimal" />
                      </DeleteParameters>
                      <UpdateParameters>
                          <asp:Parameter Name="EFORMSN" Type="String" />
                          <asp:Parameter Name="nNAME" Type="String" />
                          <asp:Parameter Name="nSEX" Type="String" />
                          <asp:Parameter Name="nSERVICE" Type="String" />
                          <asp:Parameter Name="nKIND" Type="String" />
                          <asp:Parameter Name="nID" Type="String" />
                          <asp:Parameter Name="Receive_Num" Type="Decimal" />
                      </UpdateParameters>
                      <SelectParameters>
                          <asp:ControlParameter ControlID="Lab_eformsn" Name="EFORMSN" PropertyName="Text"
                              Type="String" />
                      </SelectParameters>
                      <InsertParameters>
                          <asp:Parameter Name="EFORMSN" Type="String" />
                          <asp:Parameter Name="nNAME" Type="String" />
                          <asp:Parameter Name="nSEX" Type="String" />
                          <asp:Parameter Name="nSERVICE" Type="String" />
                          <asp:Parameter Name="nKIND" Type="String" />
                          <asp:Parameter Name="nID" Type="String" />
                      </InsertParameters>
                  </asp:SqlDataSource>
                  
                  &nbsp;&nbsp;&nbsp; &nbsp;
              <table border="0" style="width: 750px; height: 57px" >
              
                <% If read_only = "2" Then%>
                <tr>
                    <td style="height: 26px; width: 100px;" >
                        &nbsp;<asp:Label ID="Label28" runat="server" ForeColor="Black" Text="批核意見：" CssClass="form"></asp:Label><br />
                    </td>
                    <td style="width: 630px" >
                        <asp:TextBox ID="txtcomment" runat="server" Height="59px" MaxLength="255" Rows="3"
                            TextMode="MultiLine" Width="530px"></asp:TextBox>
                        <asp:Button ID="But_PHRASE" runat="server" Text="批核片語" /></td>
                </tr> 
                <% End If%>
                
                <tr>                   
                   <td style="height: 23px;" align="center" colspan="2">
                  <asp:Button ID="But_exe" runat="server" Text="送件" />
                       <asp:Button ID="Button3" runat="server"
                           Text="列印" />&nbsp;<asp:Button ID="backBtn" runat="server" Text="駁回" />&nbsp;<asp:Button
                               ID="tranBtn" runat="server" Text="呈轉" /></td>
                </tr> 
                <tr>    
                   <td colspan="2">
                  <uc1:FlowRoute ID="FlowRoute1" runat="server" />
                   </td>
                </tr>  
              </table>
              </td>
          </tr> 
        </table>    
    </div>      
     
      <div id="Div_grid10" runat="server" 
        style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:300pt; height:80pt; left:779px; top:702px; display:block;" 
        visible="false">
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
        
        
    <script language="javascript">
    var errmsg='<%= do_sql.G_errmsg %>';
    var print_file='<%= print_file%>';

    function window_onload() {
        if (errmsg != '') {
            alert(errmsg);
        }
        
        if (print_file != '') {
            lst.navigate(print_file);              
        }
    }   

    </script>
</body>
</html>
