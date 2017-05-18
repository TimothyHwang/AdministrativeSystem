<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA01003.aspx.vb" Inherits="Source_01_MOA01003" %>

<%@ Register Src="../90/FlowRoute.ascx" TagName="FlowRoute" TagPrefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>銷假單</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    
        <table width="740" border=1 cellspacing=0 cellpadding=5 bgcolor=#ffffff bordercolor=#6699cc bordercolorlight=#74a3d6 bordercolordark=#000000 style="LEFT: 20px; TOP: 10px">
        <tr>
            <td width="740" valign=bottom bgcolor=#6699cc bordercolorlight=#66aaaa bordercolordark=#ffffff style="height: 33px"><font color=white><b>&nbsp;銷假申請單</b></font>
            </td>
        </tr>
        <tr>
            <td width="740">
            
            <fieldset id="tableB">  
	        <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		        <tr>
			        <td noWrap width="10%" style="height: 16px"><asp:Label ID="Label1" runat="server" Text="填表人單位：" CssClass="form" ForeColor="Black"></asp:Label></td>
			        <td noWrap width="25%" style="height: 16px"><asp:label id="Lab_ORG_NAME_1" runat="server" CssClass="form" ForeColor="Black"></asp:label></td>
			        <td noWrap width="10%" style="height: 16px"><asp:Label ID="Label2" runat="server" Text="姓名：" CssClass="form" ForeColor="Black"></asp:Label></td>
			        <td noWrap width="25%" style="height: 16px"><asp:label id="Lab_emp_chinese_name" runat="server" CssClass="form" ForeColor="Black"></asp:label></td>
			        <td noWrap style="height: 16px; width: 75px;"><asp:Label ID="Label3" runat="server" Text="級職：" CssClass="form" ForeColor="Black"></asp:Label></td>
			        <td noWrap width="20%" style="height: 16px"><asp:label id="Lab_title_name_1" runat="server" CssClass="form" ForeColor="Black"></asp:label></td>
		        </tr>
		        <tr>
			        <td style="height: 22px"><asp:Label ID="Label4" runat="server" Text="申請人單位：" CssClass="form" ForeColor="Black"></asp:Label></td>
			        <td style="height: 22px"><asp:Label ID="Lab_ORG_NAME_2" runat="server" CssClass="form" ForeColor="Black"></asp:Label></td>
			        <td style="height: 22px">
                        <asp:Label ID="Label7" runat="server" CssClass="form" ForeColor="Black" Text="姓名："></asp:Label></td>
			        <td style="height: 22px">
                        <asp:Label ID="Lab_emp_chinese_name2" runat="server" CssClass="form" ForeColor="Black"></asp:Label></td>
			        <td style="height: 22px; width: 75px;"><asp:Label ID="Label6" runat="server" Text="級職：" CssClass="form" ForeColor="Black"></asp:Label></td>
			        <td style="height: 22px"><asp:Label ID="Lab_title_name_2" runat="server" CssClass="form" ForeColor="Black"></asp:Label></td>
		        </tr>
	        </table>
		    </fieldset>
    	    
		    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
			    <TR>
                    <td noWrap width="10%" class="form" style="height: 25px" ><asp:Label ID="Label5" runat="server" Text="銷假單申請時間：" CssClass="form" ForeColor="Black" Width="79px"></asp:Label></TD>
				    <td noWrap width="90%" class="form" style="height: 25px" ><asp:label id="AppDate" runat="server" ForeColor="Black" Width="150px"></asp:label></TD>
			    </TR>
			    <TR>
                    <td colSpan="2">
                        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" DataSourceID="SqlDataSource3"
                            Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
                            <Columns>
                                <asp:BoundField DataField="nAPPTIME" HeaderText="申請日期" SortExpression="nAPPTIME">
                                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="nTYPE" HeaderText="假別" SortExpression="nTYPE">
                                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="nSTARTTIME" HeaderText="請假起日" SortExpression="nSTARTTIME">
                                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="nSTHOUR" HeaderText="時" SortExpression="nSTHOUR">
                                    <HeaderStyle HorizontalAlign="Center" Width="5%" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="nENDTIME" HeaderText="請假迄日" SortExpression="nENDTIME">
                                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="nETHOUR" HeaderText="時" SortExpression="nETHOUR">
                                    <HeaderStyle HorizontalAlign="Center" Width="5%" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="nREASON" HeaderText="事由" SortExpression="nREASON">
                                    <HeaderStyle HorizontalAlign="Center" Width="20%" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                            </Columns>
                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <EditRowStyle BackColor="#999999" />
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        </asp:GridView>
                    </TD>
			    </TR>
                <% If read_only = "2" Then%>
			    <TR>
				    <td noWrap class="form" style="height: 22px; width: 10%;" >&nbsp; 批核意見：</TD>
				    <td noWrap width="90%" class="form" colSpan="3" style="height: 22px" valign="middle">
                       <asp:TextBox ID="txtcomment" runat="server" Height="59px" MaxLength="255" Rows="3"
                                TextMode="MultiLine" Width="530px"></asp:TextBox>
                        <asp:Button ID="But_PHRASE" runat="server" Text="批核片語" /></TD>
			    </TR>
                <% End If%>
			    <TR>
				    <td class="form" colSpan="2" style="height: 24px">
					    <DIV align="center"><asp:button id="send" runat="server" Text="送件"></asp:button>
                            <asp:Button ID="backBtn" runat="server" Text="駁回" />
                            <asp:Button ID="tranBtn" runat="server" Text="呈轉" />
                        </DIV>
				    </TD>
			    </TR>			
    			
	        </table>
	        
            <uc1:FlowRoute ID="FlowRoute1" runat="server" />
	    
            </td>
        </tr>
        </table>
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [nAPPTIME], [nTYPE], CONVERT (char(12), nSTARTTIME, 111) AS nSTARTTIME, [nSTHOUR], CONVERT (char(12), nENDTIME, 111) AS nENDTIME, [nETHOUR], [nREASON] FROM [P_01] WHERE 1=2">
        </asp:SqlDataSource>
        <br />
        
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
</body>
</html>
