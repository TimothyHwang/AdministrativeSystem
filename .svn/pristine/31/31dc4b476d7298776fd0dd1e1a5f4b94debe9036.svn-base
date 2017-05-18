<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA01013.aspx.vb" Inherits="M_Source_01_MOA01013" %>

<%@ Register src="../90/FlowRoute.ascx" tagname="FlowRoute" tagprefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>加班申請</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />

    <link href="../../css/jquery.datepick.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.10.2.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick-zh-TW.js" type="text/javascript"></script>

    <script type="text/javascript" language="javascript">
        document.oncontextmenu = new Function("return false");

        $(function () {
            $("#ApplyDate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#ApplyDate").val(), showTrigger: '#calImg' });
            
        }); 
    </script>
    <style type="text/css">

.form
{
	font: 13px Verdana, Arial, Helvetica, sans-serif;
	color: #666666;
	text-decoration: none;
}

    </style>
</head>
<body>
    <form id="form1" runat="server">
     <div>
     
        <table width="740" border=1 cellspacing=0 cellpadding=5 bgcolor=#ffffff bordercolor=#6699cc bordercolorlight=#74a3d6 bordercolordark=#000000 style="LEFT: 20px; TOP: 10px">
        <tr>
            <td valign=bottom bgcolor=#6699cc bordercolorlight=#66aaaa bordercolordark=#ffffff style="height: 33px; width: 727px;"><font color=white><b>&nbsp;個人加班申請單</b></font>
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
                        <td style="width: 120px">*
                            <asp:Label ID="Label8" runat="server" ForeColor="Black" Text="申請時間：" 
                                CssClass="form"></asp:Label></td>
                        <td style="width: 260px"  >
                            <asp:TextBox ID="ApplyDate" runat="server" Width="80px" OnKeyDown="return false" ></asp:TextBox>
                    <div style="display: none;">
	                    <img id="Img1" src="../../Image/calendar.gif" alt="選擇日期" />
                    </div>
                            <asp:DropDownList ID="ddlStartHour" runat="server">
                                <asp:ListItem>0</asp:ListItem>
                                <asp:ListItem>1</asp:ListItem>
                                <asp:ListItem>2</asp:ListItem>
                                <asp:ListItem>3</asp:ListItem>
                                <asp:ListItem>4</asp:ListItem>
                                <asp:ListItem>5</asp:ListItem>
                                <asp:ListItem>6</asp:ListItem>
                                <asp:ListItem>7</asp:ListItem>
                                <asp:ListItem>8</asp:ListItem>
                                <asp:ListItem>9</asp:ListItem>
                                <asp:ListItem>10</asp:ListItem>
                                <asp:ListItem>11</asp:ListItem>
                                <asp:ListItem>12</asp:ListItem>
                                <asp:ListItem>13</asp:ListItem>
                                <asp:ListItem>14</asp:ListItem>
                                <asp:ListItem>15</asp:ListItem>
                                <asp:ListItem>16</asp:ListItem>
                                <asp:ListItem>17</asp:ListItem>
                                <asp:ListItem>18</asp:ListItem>
                                <asp:ListItem>19</asp:ListItem>
                                <asp:ListItem>20</asp:ListItem>
                                <asp:ListItem>21</asp:ListItem>
                                <asp:ListItem>22</asp:ListItem>
                                <asp:ListItem>23</asp:ListItem>
                            </asp:DropDownList>
                            <asp:Label ID="Label29" runat="server" Text="時~" ForeColor="Black"></asp:Label>
                            <asp:DropDownList ID="ddlEndHour" runat="server">
                                <asp:ListItem>0</asp:ListItem>
                                <asp:ListItem>2</asp:ListItem>
                                <asp:ListItem>3</asp:ListItem>
                                <asp:ListItem>4</asp:ListItem>
                                <asp:ListItem>5</asp:ListItem>
                                <asp:ListItem>6</asp:ListItem>
                                <asp:ListItem>7</asp:ListItem>
                                <asp:ListItem>8</asp:ListItem>
                                <asp:ListItem>9</asp:ListItem>
                                <asp:ListItem>10</asp:ListItem>
                                <asp:ListItem>11</asp:ListItem>
                                <asp:ListItem>12</asp:ListItem>
                                <asp:ListItem>13</asp:ListItem>
                                <asp:ListItem>14</asp:ListItem>
                                <asp:ListItem>15</asp:ListItem>
                                <asp:ListItem>16</asp:ListItem>
                                <asp:ListItem>17</asp:ListItem>
                                <asp:ListItem>18</asp:ListItem>
                                <asp:ListItem>19</asp:ListItem>
                                <asp:ListItem>20</asp:ListItem>
                                <asp:ListItem>21</asp:ListItem>
                                <asp:ListItem>22</asp:ListItem>
                                <asp:ListItem>23</asp:ListItem>
                            </asp:DropDownList>
                            <asp:Label ID="Label30" runat="server" Text="時" ForeColor="Black"></asp:Label>
                    </td>
                    </tr> 
                    <tr>
                        <td style="height: 26px; width: 120px;" >*
                            <asp:Label ID="Label31" runat="server" ForeColor="Black" Text="加班地點：" 
                                CssClass="form"></asp:Label>
                        </td>
                        <td width="560">
                    <div style="display: none;">
	                    <img id="Img2" src="../../Image/calendar.gif" alt="選擇日期" />
                    </div>
                            <asp:TextBox ID="txtLocation" runat="server"></asp:TextBox>
                            <asp:Label ID="Label32" runat="server" Text="*內容不可超過15個字"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td style="height: 26px; width: 120px;" >*
                            <asp:Label ID="Label13" runat="server" ForeColor="Black" Text="加班事由：" 
                                CssClass="form"></asp:Label>
                        </td>
                        <td width="560">
                            <asp:TextBox ID="TXT_nREASON" runat="server" Height="59px" MaxLength="255" Rows="3"
                                TextMode="MultiLine" Width="560px"></asp:TextBox><br />
                            <asp:Label ID="Label33" runat="server" Text="*內容不可超過25個字"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td>&nbsp;</td>
                        <td>
                            <asp:Button ID="btnSubmit" runat="server" Text="送出" />
                            <asp:Button ID="btnModify" runat="server" Text="修改" Visible="False" />
                            <asp:Button ID="btnCancel" runat="server" Text="取消" Visible="False" />
                        </td>
                    </tr>                     
                    </table>   
            
            
            </td>
        </tr>
        
        
        </table>
        
		      
         </div>
         
             
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
