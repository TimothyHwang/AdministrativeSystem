<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA03002.aspx.vb" Inherits="Source_03_MOA03002" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>車輛派遣紀錄表</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body language="javascript" onload="return window_onload()">
    <form id="form1" runat="server">
    <div>
         <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="車輛派遣紀錄表" Width="100%"></asp:Label>
            </td></tr>
        </table>
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
            <tr>
			    <td noWrap style="height: 25px; width: 10%;">
                <asp:Label ID="Label20" runat="server" Text="目的：" CssClass="form" Width="35px"></asp:Label></td>
			    <td noWrap style="height: 25px; width: 15%;">
                    <asp:TextBox ID="TXT_nREASON" runat="server"  MaxLength="50" Width="80px"></asp:TextBox></td>
			    <td noWrap style="height: 25px; width: 10%;">
                    <asp:Label ID="Label1" runat="server" Text="報到日期：" CssClass="form" Width="61px"></asp:Label></td>
			    <td noWrap style="height: 25px; width: 15%;">
                    <asp:TextBox ID="Txt_nARRDATE" runat="server"  MaxLength="50" Width="80px" OnKeyDown="return false"></asp:TextBox>
                    <asp:ImageButton
                        ID="ImageButton1" runat="server" ImageUrl="~/Image/calendar.gif" /></td>
			    <td noWrap style="height: 25px; width: 10%;">
                <asp:Label ID="Label17" runat="server" Text="駕駛人：" CssClass="form" Width="48px"></asp:Label></td>
			    <td noWrap style="height: 25px; width: 15%;">
                    <asp:TextBox ID="Txt_DriveName" runat="server"  MaxLength="10" Width="80px"></asp:TextBox></td>
			    <td noWrap style="height: 25px; width: 10%;">			    
                    <asp:Label ID="Label22" runat="server" Text="車型：" CssClass="form" Width="48px"></asp:Label>
                </td>
			    <td noWrap style="height: 25px; width: 15%;">
                    <asp:DropDownList ID="DrDown_nstyle" runat="server" Width="104px" AutoPostBack="True">
                    </asp:DropDownList>
			    </td>
             </tr>      
             <tr>
			    <td noWrap style="height: 25px; width: 10%;">
                    <asp:Label ID="Label18" runat="server" Text="車輛號碼：" CssClass="form" Width="62px"></asp:Label></td>
			    <td noWrap style="height: 25px; width: 15%;">
                    <asp:TextBox ID="Txt_CarNumber" runat="server"  MaxLength="20" Width="80px"></asp:TextBox></td>
			    <td noWrap style="height: 25px; width: 10%;">
                    <asp:Label ID="Label19" runat="server" Text="報到處：" CssClass="form" Width="51px"></asp:Label></td>
			    <td noWrap style="height: 25px; width: 15%;">
                    <asp:TextBox ID="Txt_nARRIVEPLACE" runat="server"  MaxLength="20" Width="80px"></asp:TextBox></td>
			    <td noWrap style="height: 25px; width: 10%;">
                    <asp:Label ID="Label21" runat="server" Text="向何人報到：" CssClass="form" Width="78px"></asp:Label></td>
			    <td noWrap style="height: 25px; width: 15%;">
                    <asp:TextBox ID="Txt_nARRIVETO" runat="server"  MaxLength="20" Width="80px"></asp:TextBox></td>
			    <td noWrap style="height: 25px; width: 10%;">
                    <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" /></td>	
                <td noWrap style="height: 25px; width: 15%;">
                    </td>			    
             </tr>
        </table>
            
        <asp:GridView ID="GridView1" runat="server" Width="100%" DataSourceID="SqlDataSource1" AutoGenerateColumns="False" DataKeyNames="EFORMSN" CssClass="form" AllowPaging="True" CellPadding="4" ForeColor="#333333" GridLines="None" >
            <Columns>
                <asp:TemplateField HeaderText="駕駛人">
                    <EditItemTemplate>
                        &nbsp;<asp:DropDownList ID="DropDownList1" runat="server" DataSourceID="SqlDataSource2"
                            DataTextField="DriveName" DataValueField="DriveName" SelectedValue='<%# Bind("DriveName") %>' Width="75px" AutoPostBack="True" OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged">
                        </asp:DropDownList>&nbsp;
                        <asp:DropDownList ID="DropDownList6" runat="server" DataSourceID="SqlDataSource2"
                            DataTextField="DriveTel" DataValueField="DriveTel" SelectedValue='<%# Bind("DriveTel") %>' Width="75px" AutoPostBack="True" OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged" Visible="False">
                        </asp:DropDownList>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label6" runat="server" Text='<%# Bind("DriveName") %>' Width="75px"></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="5%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="車輛號碼">
                    <EditItemTemplate>
                        &nbsp;<asp:DropDownList ID="DropDownList9" runat="server" DataSourceID="SqlDataSource3"
                            DataTextField="PCI_CarNumber" DataValueField="PCI_CarNumber" SelectedValue='<%# Bind("CarNumber") %>' Width="75px" OnSelectedIndexChanged="DropDownList9_SelectedIndexChanged" >
                        </asp:DropDownList>&nbsp;
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label7" runat="server" Text='<%# Bind("CarNumber") %>' Width="60px"></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" Wrap="True" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="報到處-何人">
                    <EditItemTemplate>
                        <asp:Label ID="Label9" runat="server" Text='<%# Eval("per") %>' Width="60px"></asp:Label>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label8" runat="server" Text='<%# Eval("per") %>' Width="60px"></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="聯絡電話">
                    <EditItemTemplate>
                        <asp:Label ID="Label11" runat="server" Text='<%# Eval("nPHONE") %>' Width="60px"></asp:Label>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label10" runat="server" Text='<%# Eval("nPHONE") %>' Width="60px"></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:BoundField DataField="nENDPOINT" HeaderText="地址" SortExpression="nENDPOINT" ReadOnly="True" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="nREASON" HeaderText="目的" SortExpression="nREASON" ReadOnly="True" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="出場時間">
                    <EditItemTemplate>
                        <asp:DropDownList ID="DropDownList2" runat="server" OnSelectedIndexChanged="DropDownList2_SelectedIndexChanged"
                            SelectedValue='<%# BIND("LeaveHour") %>'>
                            <asp:ListItem></asp:ListItem>
                            <asp:ListItem>00</asp:ListItem>
                            <asp:ListItem>01</asp:ListItem>
                            <asp:ListItem>02</asp:ListItem>
                            <asp:ListItem>03</asp:ListItem>
                            <asp:ListItem>04</asp:ListItem>
                            <asp:ListItem>05</asp:ListItem>
                            <asp:ListItem>06</asp:ListItem>
                            <asp:ListItem>07</asp:ListItem>
                            <asp:ListItem>08</asp:ListItem>
                            <asp:ListItem>09</asp:ListItem>
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
                        </asp:DropDownList><asp:DropDownList ID="DropDownList3" runat="server" OnSelectedIndexChanged="DropDownList3_SelectedIndexChanged"
                            SelectedValue='<%# BIND("LeaveMin") %>'>
                            <asp:ListItem></asp:ListItem>
                            <asp:ListItem>00</asp:ListItem>
                            <asp:ListItem>10</asp:ListItem>
                            <asp:ListItem>20</asp:ListItem>
                            <asp:ListItem>30</asp:ListItem>
                            <asp:ListItem>40</asp:ListItem>
                            <asp:ListItem>50</asp:ListItem>
                        </asp:DropDownList>
                        <asp:Label ID="Label3" runat="server" CssClass="form" Text='<%# Bind("LeaveTime") %>'
                            Width="62px" Visible="False"></asp:Label>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# Bind("LeaveTime") %>' Width="79px"></asp:Label><br />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="5%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="回場時間">
                    <EditItemTemplate>
                        <asp:DropDownList ID="DropDownList4" runat="server" OnSelectedIndexChanged="DropDownList4_SelectedIndexChanged"
                            SelectedValue='<%# BIND("ComeHour") %>'>
                            <asp:ListItem></asp:ListItem>
                            <asp:ListItem>00</asp:ListItem>
                            <asp:ListItem>01</asp:ListItem>
                            <asp:ListItem>02</asp:ListItem>
                            <asp:ListItem>03</asp:ListItem>
                            <asp:ListItem>04</asp:ListItem>
                            <asp:ListItem>05</asp:ListItem>
                            <asp:ListItem>06</asp:ListItem>
                            <asp:ListItem>07</asp:ListItem>
                            <asp:ListItem>08</asp:ListItem>
                            <asp:ListItem>09</asp:ListItem>
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
                        </asp:DropDownList><asp:DropDownList ID="DropDownList5" runat="server" OnSelectedIndexChanged="DropDownList5_SelectedIndexChanged"
                            SelectedValue='<%# BIND("ComeMin") %>'>
                            <asp:ListItem></asp:ListItem>
                            <asp:ListItem>00</asp:ListItem>
                            <asp:ListItem>10</asp:ListItem>
                            <asp:ListItem>20</asp:ListItem>
                            <asp:ListItem>30</asp:ListItem>
                            <asp:ListItem>40</asp:ListItem>
                            <asp:ListItem>50</asp:ListItem>
                        </asp:DropDownList>
                        <asp:Label ID="Label5" runat="server" Text='<%# Bind("ComeTime") %>' Width="66px" Visible="False"></asp:Label>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label4" runat="server" Text='<%# Bind("ComeTime") %>' Width="72px"></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="5%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="出場里數">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox2" runat="server" Text='<%# Bind("leaveMilage") %>'
                            Width="22px" MaxLength="4"></asp:TextBox>
                        <asp:RangeValidator ID="RangeValidator1" runat="server" ControlToValidate="TextBox2"
                            ErrorMessage="必需輸入數字" MaximumValue="9999" MinimumValue="0" SetFocusOnError="True"
                            Type="Integer"></asp:RangeValidator>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label14" runat="server" Text='<%# Bind("LeaveMilage") %>'
                            Width="30px"></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="5%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="回場里數">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox3" runat="server" Text='<%# Bind("ComeMilage") %>'
                            Width="23px"></asp:TextBox>
                        <asp:RangeValidator ID="RangeValidator2" runat="server" ControlToValidate="TextBox3"
                            ErrorMessage="必需輸入數字" MaximumValue="9999" MinimumValue="0" SetFocusOnError="True"
                            Type="Integer"></asp:RangeValidator>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label15" runat="server" Text='<%# Bind("ComeMilage") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="5%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="實際里程">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox4" runat="server" Text='<%# Bind("RealMilage") %>'
                            Width="21px"></asp:TextBox>
                        <asp:RangeValidator ID="RangeValidator3" runat="server" ControlToValidate="TextBox4"
                            ErrorMessage="必需輸入數字" MaximumValue="9999" MinimumValue="0" SetFocusOnError="True"></asp:RangeValidator>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label16" runat="server" Text='<%# Bind("RealMilage") %>'
                            Width="50px"></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="5%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="車型">
                    <EditItemTemplate>
                        <asp:Label ID="Label13" runat="server" Height="37px" Text='<%# Eval("nSTYLE") %>'
                            Width="70px"></asp:Label>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label12" runat="server" Text='<%# Eval("nSTYLE") %>' Width="70px"></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="狀態" SortExpression="JoinCar">
                    <EditItemTemplate>
                        <asp:DropDownList ID="DropDownList7" runat="server" SelectedValue='<%# Bind("JoinCar") %>'>
                            <asp:ListItem>正常</asp:ListItem>
                            <asp:ListItem>併車</asp:ListItem>
                        </asp:DropDownList>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label3" runat="server" Text='<%# Bind("JoinCar") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="5%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:CommandField ShowEditButton="True" ButtonType="Image" EditImageUrl="~/Image/update.gif" EditText="修改" CancelImageUrl="~/Image/cancel.gif" UpdateImageUrl="~/Image/apply.gif" UpdateText="確認" >
                    <HeaderStyle HorizontalAlign="Center" Width="5%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:CommandField>
                <asp:BoundField DataField="EFORMSN" HeaderText="EFORMSN" ReadOnly="True" SortExpression="EFORMSN" />
            </Columns>
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
        </asp:GridView>
        <table border="0" style="width: 100%; height: 23px">
                <tr>
                   <td align="center">
                       <asp:ImageButton
                           ID="ImagePrint" runat="server" ImageUrl="~/Image/print.gif" ToolTip="列印" />
                       <asp:ImageButton ID="Img_Export" runat="server" ImageUrl="~/Image/ExportFile.gif" /></td>
                </tr> 
         </table>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT '' As [DriveName],'' As [DriveTel] &#13;&#10;union SELECT [Drive_Name] As [DriveName],[Drive_Tel] As [DriveTel]  FROM [P_0304]   where Drive_Status ='1' ORDER BY 1"></asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConflictDetection="CompareAllValues"
            ConnectionString="<%$ ConnectionStrings:ConnectionString %>"             
            OldValuesParameterFormatString="original_{0}" SelectCommand="SELECT P_03.EFORMSN, P_0301.DriveName, P_0301.CarNumber, P_03.nARRIVEPLACE + '--' + P_03.nARRIVETO AS per, P_03.nPHONE, P_03.nSTARTPOINT, P_03.nENDPOINT, P_0301.LeaveTime, P_0301.ComeTime, P_0301.LeaveMilage, P_0301.ComeMilage, P_0301.RealMilage, P_03.nSTYLE,P_0301.JoinCar FROM P_03 INNER JOIN P_0301 ON P_03.EFORMSN = P_0301.EFORMSN where P_03.EFORMSN='-1'"
            UpdateCommand="UPDATE P_0301 SET DriveName = @DriveName, CarNumber = @CarNumber, LeaveTime = @LeaveTime, ComeTime = @ComeTime, LeaveMilage = @LeaveMilage, ComeMilage = @ComeMilage, RealMilage = @RealMilage ,DriveTel = @DriveTel,JoinCar = @JoinCar WHERE (EFORMSN = @original_EFORMSN)">            
            <UpdateParameters>
                <asp:Parameter Name="DriveName" Type="String" />
                <asp:Parameter Name="CarNumber" />
                <asp:Parameter Name="LeaveTime" Type="String" />
                <asp:Parameter Name="ComeTime" Type="String" />
                <asp:Parameter Name="LeaveMilage" />
                <asp:Parameter Name="ComeMilage" />
                <asp:Parameter Name="RealMilage" />
                <asp:Parameter Name="DriveTel" />
                <asp:Parameter Name="JoinCar" />
                <asp:Parameter Name="original_EFORMSN" />
            </UpdateParameters>                        
        </asp:SqlDataSource>
        
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT DISTINCT '' AS PCI_CarNumber UNION SELECT PCI_CarNumber FROM P_0303 where PCI_Use='1' order by PCI_CarNumber">
        </asp:SqlDataSource>
        
    </div>
    
    <div id="Div_grid" runat="server" style="position:absolute; z-index:3; border:2 solid lightslategray; background-color:white; width:165pt; height:150pt; left:157px; top:637px; display:block;" visible="false">
     
        <asp:Calendar ID="Calendar1" runat="server" BackColor="White" BorderColor="#3366CC"
        BorderWidth="1px" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana"
        Font-Size="8pt" ForeColor="#003399" Height="200px" Width="220px" Caption="">
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
        
    </form>
    <iframe id="lst" frameborder=0 width=0 height=0 src="/blank.htm"></iframe> 
    <script language=javascript>
    
function window_onload() {
   <% if display_flag = True then %>
      run_div.style.left =<% =x_point %> ;
      run_div.style.top =<% =y_point %> ;   
      run_div.style.display="block";       
   <% end if %> 
   
   <%if do_sql.G_errmsg <>"" then %>  
    alert('<%= do_sql.G_errmsg.Replace("'", "") %>');
  <%end if
     if print_file <>"" then %>
        lst.navigate("<%=print_file%>");              
     <%end if %>
}

    </script>
</body>
</html>
