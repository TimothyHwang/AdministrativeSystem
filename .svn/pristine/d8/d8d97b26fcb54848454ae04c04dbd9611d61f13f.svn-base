<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA02001.aspx.vb" Inherits="Source_02_MOA02001" %>

<%@ Register Src="../90/FlowRoute.ascx" TagName="FlowRoute" TagPrefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>會議室申請單</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
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
            $("#StartDay").datepick({ formats: 'yyyy/m/d', defaultDate: $("#StartDay").val(), showTrigger: '#calImg' });            
        }); 
    </script>
</head>
<body>
    <form id="form1" runat="server">
    
        <table width="750" border=1 cellspacing=0 cellpadding=5 bgcolor=#ffffff bordercolor=#6699cc bordercolorlight=#74a3d6 bordercolordark=#000000 style="LEFT: 20px; TOP: 10px">
        <tr>
            <td valign=bottom bgcolor=#6699cc bordercolorlight=#66aaaa bordercolordark=#ffffff><font color=white><b>&nbsp;會議室申請單</b></font>
            </td>
        </tr>
        
        <tr>
            <td>
            
             <fieldset id="tableB" style="width: 740px">
            
		        <table cellSpacing="0" cellPadding="0" width="750" align="center" border="0" height="57">
			        <tr>
				        <td noWrap style="width: 80px"><asp:Label ID="Label1" runat="server" Text="填表人單位：" CssClass="form" Width="80px" ForeColor="Black"></asp:Label></td>
				        <td noWrap style="width: 150px"><asp:label id="Lab_ORG_NAME_1" runat="server" CssClass="form" Width="139px" ForeColor="Black"></asp:label></td>
				        <td noWrap style="width: 60px"><asp:Label ID="Label2" runat="server" Text="姓名：" CssClass="form" Width="50px" ForeColor="Black"></asp:Label></td>
				        <td noWrap style="width: 160px"><asp:label id="Lab_emp_chinese_name" runat="server" CssClass="form" Width="140px" ForeColor="Black"></asp:label></td>
				        <td noWrap style="width: 60px"><asp:Label ID="Label3" runat="server" Text="級職：" Width="50px" CssClass="form" ForeColor="Black"></asp:Label></td>
				        <td noWrap style="width: 130px"><asp:label id="Lab_title_name_1" runat="server" CssClass="form" ForeColor="Black" Width="130px"></asp:label></td>
			        </tr>
			        <tr>
				        <td noWrap style="width: 80px"><asp:Label ID="Label4" runat="server" Text="申請人單位：" Width="80px" CssClass="form" ForeColor="Black"></asp:Label></td>
				        <td noWrap style="width: 150px"><asp:Label ID="Lab_ORG_NAME_2" runat="server" Width="140px" CssClass="form" ForeColor="Black"></asp:Label></td>
				        <td noWrap style="width: 60px"><asp:Label ID="Label5" runat="server" Text="姓名：" Width="50px" CssClass="form" ForeColor="Black"></asp:Label></td>
				        <td noWrap style="width: 160px"><asp:DropDownList ID="DrDown_emp_chinese_name" runat="server" Width="140px" AutoPostBack="True" DataSourceID="SqlDataSource1" DataTextField="emp_chinese_name" DataValueField="employee_id"></asp:DropDownList></td>
				        <td noWrap style="width: 60px"><asp:Label ID="Label6" runat="server" Text="級職：" Width="50px" CssClass="form" ForeColor="Black"></asp:Label></td>
				        <td noWrap style="width: 130px"><asp:Label ID="Lab_title_name_2" runat="server" Width="130px" CssClass="form" ForeColor="Black"></asp:Label></td>
			        </tr>
		        </table>
    			    
		    </fieldset>    		
    		
		    <table cellSpacing="0" cellPadding="0" width="750" align="center" border="0">
    		
			    <TR>
				    <td noWrap class="form" style="height: 33px; width: 10%;"><FONT color="#ff0033">* <asp:Label ID="Label12" runat="server" CssClass="form" ForeColor="Black" Text="聯絡電話："></asp:Label></FONT>
                    </TD>
				    <td noWrap class="form" style="height: 33px; width: 40%;"><asp:textbox id="PHONE" runat="server" MaxLength="10"></asp:textbox></TD>
				    <td noWrap class="form" style="height: 33px; width: 10%;">&nbsp;&nbsp;
                        <asp:Label ID="Label13" runat="server" CssClass="form" ForeColor="Black" Text="申請時間："></asp:Label></TD>
				    <td noWrap class="form" style="height: 33px; width: 40%;"><asp:label id="AppDate" runat="server" ForeColor="Black" Width="150px"></asp:label></TD>
			    </TR>
			    
			    <TR>
				    <td noWrap class="form" style="height: 33px; width: 10%;"><FONT color="#ff0033">* <asp:Label ID="Label11" runat="server" CssClass="form" ForeColor="Black"
                            Text="會議名稱："></asp:Label></FONT></TD>
				    <td noWrap class="form" style="height: 33px; width: 40%;"><asp:textbox id="MeetName" runat="server" Width="250px" MaxLength="50"></asp:textbox></TD>
				    <td noWrap class="form" style="height: 33px; width: 10%;"><FONT color="#ff0033">* <asp:Label ID="Label14" runat="server" CssClass="form" ForeColor="Black"
                            Text="會議類別："></asp:Label></FONT></TD>
				    <td noWrap class="form" style="height: 33px; width: 40%;"><FONT face="新細明體"><asp:dropdownlist id="Status" runat="server">
							    <asp:ListItem Value=" " >請選擇</asp:ListItem>
							    <asp:ListItem Value="例行會議">例行會議</asp:ListItem>
							    <asp:ListItem Value="研討會">研討會</asp:ListItem>
                        <asp:ListItem>研討會議</asp:ListItem>
                        <asp:ListItem Value="訓練課程"></asp:ListItem>
                        <asp:ListItem>專案會議</asp:ListItem>
                        <asp:ListItem>管理會議</asp:ListItem>
                        <asp:ListItem>其他</asp:ListItem>
						        </asp:dropdownlist>&nbsp;</FONT></TD>
			    </TR>			
    			
			    <TR>
				    <td noWrap class="form" style="height: 33px; width: 10%;"><FONT color="#ff0033">* <asp:Label ID="Label10" runat="server" CssClass="form" ForeColor="Black"
                            Text="主持人："></asp:Label></FONT></TD>
				    <td noWrap class="form" style="height: 33px; width: 40%;"><asp:textbox id="speaker" runat="server" Width="218px" MaxLength="50"></asp:textbox></TD>
				    <td noWrap class="form" style="height: 33px; width: 10%;"><FONT color="#ff0033">* <asp:Label ID="Label15" runat="server" CssClass="form" ForeColor="Black"
                            Text="會議地點："></asp:Label></FONT></TD>
				    <td noWrap class="form" style="height: 33px; width: 40%;"><FONT face="新細明體"><asp:dropdownlist id="Place" runat="server" DataSourceID="SqlDataSource3" DataTextField="MeetName" DataValueField="MeetSn" AutoPostBack="True"></asp:dropdownlist></FONT></TD>
			    </TR>			
			    <TR>
				    <td noWrap class="form" style="height: 33px; width: 10%;"><FONT color="#ff0033">* </FONT><asp:Label ID="Label9" runat="server" CssClass="form"
                            ForeColor="Black" Text="會議時間："></asp:Label></TD>
				    <td noWrap class="form" style="height: 33px; width: 40%;">
					    <asp:textbox id="StartDay" runat="server" OnKeyDown="return false" Width="100px"></asp:textbox>
                        <div style="display: none;">
                        <img id="calImg" src="../../Image/calendar.gif" alt="選擇日期" />
                        </div>
                    </td>
                        
                    <td noWrap class="form" style="height: 33px; width: 10%;">
                    &nbsp;&nbsp;
                        <asp:Label ID="Label16" runat="server" CssClass="form" ForeColor="Black" Text="時段："></asp:Label>
                    </td>
                    <td noWrap class="form" style="height: 33px; width: 40%;">
				    <asp:dropdownlist id="htime_start" runat="server">
					    <asp:ListItem Value="全天" Selected="True">全天</asp:ListItem>
					    <asp:ListItem Value="上午">上午</asp:ListItem>
					    <asp:ListItem Value="下午">下午</asp:ListItem>
				    </asp:dropdownlist>
                    <asp:Button ID="InsMeeting" runat="server" Text="新增會議時間" />
                        <asp:Label ID="labMeetErr" runat="server" ForeColor="Red" Width="91px"></asp:Label></td>
			    </TR>
			    <TR>
				    <td noWrap width="10%">&nbsp;&nbsp;
				    <asp:Label ID="Label7" runat="server" CssClass="form" ForeColor="Black" Text="會議時間："></asp:Label></TD>
				    <td noWrap width="90%" class="form" colSpan="3">
                        <asp:Label ID="labMeetNull" runat="server" ForeColor="Red" Width="83px"></asp:Label><asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" DataKeyNames="MTT_Num"
                            DataSourceID="SqlDataSource2" Width="600px" CssClass="form" ForeColor="Black" PageSize="5">
                            <Columns>
                                <asp:BoundField DataField="MTT_Num" HeaderText="流水號" InsertVisible="False" ReadOnly="True"
                                    SortExpression="MTT_Num" />
                                <asp:BoundField DataField="MeetName" HeaderText="會議室名稱" SortExpression="MeetName" />
                                <asp:BoundField DataField="MeetTime" HeaderText="會議日期" SortExpression="MeetTime" />
                                <asp:BoundField DataField="MeetHour" HeaderText="會議時段" SortExpression="MeetHour" />
                                <asp:CommandField ButtonType="Image" DeleteImageUrl="~/Image/delete.gif" ShowDeleteButton="True" />
                            </Columns>
                            <EmptyDataTemplate>
                                <asp:Label ID="Label8" runat="server" CssClass="form" Text="尚無開會時間"></asp:Label>
                            </EmptyDataTemplate>
                        </asp:GridView>
                        <br />
                    </td>
			    </TR>
			    <TR>
				    <td noWrap class="form" style="height: 54px; width: 10%;">&nbsp;&nbsp;
                        <asp:Label ID="Label17" runat="server" CssClass="form" ForeColor="Black" Text="參加人員："></asp:Label><br />
                        <asp:Label ID="Label22" runat="server" CssClass="form" ForeColor="Black" Text="(以255字為限)"></asp:Label></TD>
				    <td noWrap width="90%" class="form" colSpan="3" style="height: 54px"><asp:textbox id="JoinPerson" runat="server" TextMode="MultiLine" Rows="3" Width="550"></asp:textbox></td>
			    </TR>
			    <TR>
				    <td noWrap class="form" style="height: 54px; width: 10%;">&nbsp;&nbsp;
                        <asp:Label ID="Label18" runat="server" CssClass="form" ForeColor="Black" Text="會議內容："></asp:Label><br />
                        <asp:Label ID="Label21" runat="server" CssClass="form" ForeColor="Black" Text="(以255字為限)"></asp:Label></TD>
				    <td noWrap width="90%" class="form" colSpan="3" style="height: 54px"><asp:textbox id="MeetContect" runat="server" TextMode="MultiLine" Rows="3" Width="550"></asp:textbox></TD>
			    </TR>
			    <TR>
				    <td noWrap class="form" style="width: 10%; height: 38px" >&nbsp;&nbsp;
                        <asp:Label ID="Label19" runat="server" CssClass="form" ForeColor="Black" Text="會議室設備："></asp:Label></TD>
				    <td noWrap width="90%" class="form" colSpan="3" valign="middle" style="height: 38px">
                        <asp:BulletedList ID="BulletedList1" runat="server" BulletStyle="Numbered" DataSourceID="SqlDataSource4"
                            DataTextField="DeviceName" DataValueField="DeviceName" ForeColor="Black">
                        </asp:BulletedList>
                    </TD>
			    </TR>
                <% If read_only = "2" Then%>
			    <TR>
				    <td noWrap class="form" style="height: 27px; width: 10%;" >&nbsp;&nbsp;
                        <asp:Label ID="Label20" runat="server" CssClass="form" ForeColor="Black" Text="批核意見："></asp:Label></TD>
				    <td noWrap width="90%" class="form" colSpan="3" style="height: 27px" valign="middle">
                       <asp:TextBox ID="txtcomment" runat="server" Height="59px" MaxLength="255" Rows="3"
                                TextMode="MultiLine" Width="550px"></asp:TextBox>
                        <asp:Button ID="But_PHRASE" runat="server" Text="批核片語" /></TD>
			    </TR>
                <% End If%>
			    <TR>
				    <td class="form" colSpan="4" style="height: 24px">
					    <DIV align="center"><asp:button id="send" runat="server" Text="送件"></asp:button>
                            <asp:Button ID="backBtn" runat="server" Text="駁回" />
                            <asp:Button ID="tranBtn" runat="server" Text="呈轉" />
                            <asp:Button ID="supplyBtn" runat="server" Text="補登" />
                        </DIV>
				    </TD>
			    </TR>
		    </table>		
            <uc1:FlowRoute ID="FlowRoute1" runat="server" />
            
            </td>
        </tr>
        </table>
            
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT * FROM [V_EmpInfo] WHERE ([ORG_UID] = @ORG_UID) ORDER BY [emp_chinese_name]">
            <SelectParameters>
                <asp:SessionParameter Name="ORG_UID" SessionField="org_uid" Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            DeleteCommand="DELETE FROM [P_0204] WHERE [MTT_Num] = @MTT_Num" InsertCommand="INSERT INTO [P_0204] ([EFORMSN], [MeetSn], [MeetTime], [MeetHour]) VALUES (@EFORMSN, @MeetSn, @MeetTime, @MeetHour)"
            SelectCommand="SELECT P_0201.MeetName, P_0204.MTT_Num, P_0204.MeetSn,CONVERT(nvarchar,P_0204.MeetTime,111) as MeetTime, P_0204.MeetHour FROM P_0204 INNER JOIN P_0201 ON P_0204.MeetSn = P_0201.MeetSn WHERE (P_0204.EFORMSN = @EFORMSN)" UpdateCommand="UPDATE [P_0204] SET [EFORMSN] = @EFORMSN, [MeetSn] = @MeetSn, [MeetTime] = @MeetTime, [MeetHour] = @MeetHour WHERE [MTT_Num] = @MTT_Num">
            <SelectParameters>
                <asp:QueryStringParameter Name="EFORMSN" QueryStringField="eformsn" Type="String" />
            </SelectParameters>
            <DeleteParameters>
                <asp:Parameter Name="MTT_Num" Type="Decimal" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="EFORMSN" Type="String" />
                <asp:Parameter Name="MeetSn" Type="Decimal" />
                <asp:Parameter Name="MeetTime" Type="DateTime" />
                <asp:Parameter Name="MeetHour" Type="String" />
                <asp:Parameter Name="MTT_Num" Type="Decimal" />
            </UpdateParameters>
            <InsertParameters>
                <asp:Parameter Name="EFORMSN" Type="String" />
                <asp:Parameter Name="MeetSn" Type="Decimal" />
                <asp:Parameter Name="MeetTime" Type="DateTime" />
                <asp:Parameter Name="MeetHour" Type="String" />
            </InsertParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            
            SelectCommand="SELECT [MeetSn], [MeetName] FROM [P_0201] WHERE (share = @share OR Org_Uid=@Org_Uid) AND (Enabled=1) ORDER BY [MeetName]">
            <SelectParameters>
                <asp:Parameter DefaultValue="1" Name="Share" Type="Int32" />
                <asp:SessionParameter DefaultValue="" Name="Org_Uid" SessionField="org_uid" />
            </SelectParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource4" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT P_0202.DeviceName FROM P_0202 INNER JOIN P_0203 ON P_0202.DeviceSn = P_0203.DeviceSn WHERE (P_0203.MeetSn = @PLACE) ORDER BY P_0203.MeetSn">
            <SelectParameters>
                <asp:ControlParameter ControlID="Place" Name="PLACE" PropertyName="SelectedValue" />
            </SelectParameters>
        </asp:SqlDataSource>
        &nbsp;
                
         <div id="Div_grid10" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:300pt; height:80pt; left:249px; top:893px; display:block;" visible="false">
            <asp:GridView id="GridView2" runat="server" CssClass="form" Width="100%" Height="50px" DataSourceID="SqlDataSource10" PageSize="5" AutoGenerateColumns="False" AllowPaging="True" BorderColor="Lime" BorderWidth="2px">
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
