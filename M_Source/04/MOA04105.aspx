<%@ Page Language="VB" EnableEventValidation = "false" AutoEventWireup="false" CodeFile="MOA04105.aspx.vb" Inherits="M_Source_04_MOA04105" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>房屋水電修繕統計表</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
      <script type="text/JavaScript">
          function Check(type) {
              var re = /</;
              if (type == 1) {
                  if (document.getElementById("tbnFacilityNo").value == "" && document.getElementById("tbit_Name").value == "") {
                      alert("查詢輸入的設備編號與物料名稱請勿2者都空白！");
                      return false;
                  }

                  if (re.test(document.getElementById("tbnFacilityNo").value)) {
                      alert("您輸入的設備編號有包含危險字元！");
                      document.getElementById("tbnFacilityNo").value = "";
                      return false;
                  }

                  if (re.test(document.getElementById("tbit_Name").value)) {
                      alert("您輸入的物料名稱有包含危險字元！");
                      document.getElementById("tbit_Name").value = "";
                      return false;
                  } 
              }

              if (type == 2) {

                  if (document.getElementById("tbtype1name").value == "" && document.getElementById("tbtype2name").value == "") {
                      alert("查詢輸入的主要工程類別與次要工程類別請勿2者都空白！");
                      return false;
                  }


                  if (re.test(document.getElementById("tbtype1name").value)) {
                      alert("您輸入的主要工程類別名稱有包含危險字元！");
                      document.getElementById("tbtype1name").value = "";
                      return false;
                  }

                  if (re.test(document.getElementById("tbtype2name").value)) {
                      alert("您輸入的次要工程類別類別名稱有包含危險字元！");
                      document.getElementById("tbtype2name").value = "";
                      return false;
                  }
              }
              return true;
          }
</script>
</head>
<body>
    <form id="form1" runat="server">
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="房屋水電修繕統計表" Width="100%"></asp:Label>
            </td></tr>
        </table>   
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap style="height: 25px; width: 5%;">
                    <asp:Label ID="Lab1" runat="server" Text="統計查詢：" CssClass="form"></asp:Label>                    
                </td>
			    <td noWrap style="height: 25px; width: 10%;">
                    <asp:DropDownList ID="ddl_QueryType" runat="server" AutoPostBack="True">
                        <asp:ListItem Value="1">故障原因統計</asp:ListItem>
                        <asp:ListItem Value="2">建築物統計</asp:ListItem>
                        <asp:ListItem Value="3">維修項目統計</asp:ListItem>
                        <asp:ListItem Value="4">報修類別統計</asp:ListItem>
                        <asp:ListItem Value="5">用料統計</asp:ListItem>
                        <asp:ListItem Value="6">設備分佈統計</asp:ListItem>
                    </asp:DropDownList>
                    </td>
                <td noWrap style="height: 25px; width: 50%;">
             
                   <div id="divYear" runat = "server" CssClass="form">
                     <asp:Label ID="Label5" runat="server" Text="完工日期：" CssClass="form" />
                     <asp:TextBox ID="nStartDATE1" runat="server" Width="65px" OnKeyDown="return false" />
                    <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />
                <asp:Label ID="Label6" runat="server" Text="~" CssClass="form"></asp:Label>
                <asp:TextBox ID="nStartDATE2" runat="server" Width="65px" OnKeyDown="return false"></asp:TextBox>
                    <asp:ImageButton ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" />  &nbsp;&nbsp;&nbsp;
                    <asp:Button ID="btQueryYear" runat="server" Text="統計查詢" /> 
                    <asp:Label ID="lbDTMsg" runat="server" CssClass="form" />
                </div>
                <div id="dvFacifity" runat = "server" visible ="false" CssClass="form">
                     <asp:Label ID="Label1" runat="server" Text="設備(物品)編號：" CssClass="form" /><asp:TextBox ID="tbnFacilityNo" runat="server" MaxLength ="50" />
                     &nbsp;&nbsp;&nbsp;
                     <asp:Label ID="Label2" runat="server" Text="物品(料)名稱：" CssClass="form" /><asp:TextBox ID="tbit_Name" runat="server"  MaxLength ="225" />
                    <asp:Button ID="btQueryFacifity" runat="server" Text="統計查詢" />
                </div>
                 <div id="dvitCode" runat = "server" visible ="false" CssClass="form">
                     <asp:Label ID="Label4" runat="server" Text="主要工程類別名稱：" CssClass="form" /><asp:TextBox ID="tbtype1name" runat="server"  MaxLength ="225" />
                     &nbsp;&nbsp;&nbsp;
                     <asp:Label ID="Label3" runat="server" Text="次要工程類別名稱：" CssClass="form" /><asp:TextBox ID="tbtype2name" runat="server"  MaxLength ="225" />
                    <asp:Button ID="btQueryitCode" runat="server" Text="統計查詢" />
                 </div>
                </td>
			    <td align="right" noWrap style="height: 25px; width: 20%;">
                    <asp:ImageButton ID="Img_Export" runat="server" 
                        ImageUrl="~/Image/ExportFile.gif" Enabled="False" /></td>
		    </tr>		    
	    </table>
        <center>
             <asp:GridView ID="gvErrCause" runat="server" AutoGenerateColumns="False" ShowFooter = "true"  Visible = "false"
        CellPadding="4" EmptyDataText="查無任何資料" ForeColor="#333333" GridLines="None" Width="500px">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EditRowStyle BackColor="#999999" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <FooterStyle BackColor="#5D1234" Font-Bold="True" ForeColor="yellow"  HorizontalAlign="Center"  />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <Columns>
                 <asp:BoundField  DataField="nErrCause" HeaderText="編號">
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                  <asp:BoundField  DataField="nErrCauseName" HeaderText="故障原因" >
                    <ItemStyle HorizontalAlign="Center" Width="50%" />
                </asp:BoundField>
                      <asp:BoundField  DataField="Anycnt" HeaderText="筆數" >
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:LinkButton ID="LBErrCauseDetail" runat="server" CausesValidation="False" CommandArgument=<%#Eval("nErrCause") %> CommandName="AnyDetail" text="明細" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="10%" />
                </asp:TemplateField> 
                
            </Columns>
        </asp:GridView>
            <asp:GridView ID="gvBuilding" runat="server" AutoGenerateColumns="False" ShowFooter = "true"  Visible = "false"
        CellPadding="4" EmptyDataText="查無任何資料" ForeColor="#333333" GridLines="None" Width="600px">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EditRowStyle BackColor="#999999" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <FooterStyle BackColor="#5D1234" Font-Bold="True" ForeColor="yellow"  HorizontalAlign="Center"  />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                <PagerStyle HorizontalAlign="Center" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <Columns>
                 <asp:BoundField  DataField="bd_code" HeaderText="編號">
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                  <asp:BoundField  DataField="bd_name" HeaderText="建築物名稱" >
                    <ItemStyle HorizontalAlign="Center" Width="50%" />
                </asp:BoundField>
                      <asp:BoundField  DataField="Anycnt" HeaderText="筆數" >
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:LinkButton ID="LBBuildingDetail" runat="server" CausesValidation="False" CommandArgument=<%#Eval("bd_code") %> CommandName="AnyDetail" text="明細" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="10%" />
                </asp:TemplateField> 
                
            </Columns>
        </asp:GridView>
            <asp:GridView ID="gvFacility" runat="server" AutoGenerateColumns="False" ShowFooter = "true"  Visible = "false"
        CellPadding="4" EmptyDataText="查無任何資料" ForeColor="#333333" GridLines="None" Width="600px">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EditRowStyle BackColor="#999999" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <FooterStyle BackColor="#5D1234" Font-Bold="True" ForeColor="yellow"  HorizontalAlign="Center"  />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <Columns>
                 <asp:BoundField  DataField="nFacilityNo" HeaderText="設備(物品)編號">
                    <ItemStyle HorizontalAlign="Center" Width="35%" />
                </asp:BoundField>
                  <asp:BoundField  DataField="it_name" HeaderText="設備(物品)名稱" >
                    <ItemStyle HorizontalAlign="Center" Width="35%" />
                </asp:BoundField>
                      <asp:BoundField  DataField="Anycnt" HeaderText="筆數" >
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:LinkButton ID="LBFacilityDetail" runat="server" CausesValidation="False" CommandArgument=<%#Eval("nFacilityNo") + "," + Eval("it_name") %> CommandName="AnyDetail" text="明細" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="10%" />
                </asp:TemplateField> 
                
            </Columns>
        </asp:GridView>
            <asp:GridView ID="gvItName" runat="server" AutoGenerateColumns="False" ShowFooter = "true"  Visible = "false"
        CellPadding="4" EmptyDataText="查無任何資料" ForeColor="#333333" GridLines="None" Width="600px">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EditRowStyle BackColor="#999999" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <FooterStyle BackColor="#5D1234" Font-Bold="True" ForeColor="yellow"  HorizontalAlign="Center"  />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <Columns>
              <asp:BoundField  DataField="type1name" HeaderText="主要工程類別名稱" >
                    <ItemStyle HorizontalAlign="Center" Width="35%" />
                </asp:BoundField>
                  <asp:BoundField  DataField="type2name" HeaderText="次要工程類別名稱" >
                    <ItemStyle HorizontalAlign="Center" Width="35%" />
                </asp:BoundField>
                      <asp:BoundField  DataField="Anycnt" HeaderText="筆數" >
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:LinkButton ID="LBItCodeDetail" runat="server" CausesValidation="False" CommandArgument=<%#Eval("itcode") + "," + Eval("type1name")+ "," + Eval("type2name") %> CommandName="AnyDetail" text="明細" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="10%" />
                </asp:TemplateField> 
            </Columns>
        </asp:GridView>
         <asp:GridView ID="gvitCodeName" runat="server" AutoGenerateColumns="False" ShowFooter = "true"  Visible = "false"
        CellPadding="4" EmptyDataText="查無任何資料" ForeColor="#333333" GridLines="None" Width="600px">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EditRowStyle BackColor="#999999" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <FooterStyle BackColor="#5D1234" Font-Bold="True" ForeColor="yellow"  HorizontalAlign="Center"  />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <Columns>
              <asp:BoundField  DataField="it_code" HeaderText="物品(料)編號" >
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                  <asp:BoundField  DataField="it_name" HeaderText="物品(料)名稱" >
                    <ItemStyle HorizontalAlign="Center" Width="50%" />
                </asp:BoundField>
                      <asp:BoundField  DataField="Anycnt" HeaderText="數量" >
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:LinkButton ID="LBItCodeDetail" runat="server" CausesValidation="False" CommandArgument=<%#Eval("it_code")  + "," + Eval("it_name") %> CommandName="AnyDetail" text="明細" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="10%" />
                </asp:TemplateField> 
            </Columns>
        </asp:GridView>
        <asp:GridView ID="gvBgFlRome" runat="server" AutoGenerateColumns="False" ShowFooter = "true"  Visible = "false"
        CellPadding="4" EmptyDataText="查無任何資料" ForeColor="#333333" GridLines="None" Width="650px">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EditRowStyle BackColor="#999999" />
            <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
            <FooterStyle BackColor="#5D1234" Font-Bold="True" ForeColor="yellow"  HorizontalAlign="Center"  />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <Columns>
              <asp:BoundField  DataField="bd_name" HeaderText="建物名稱" >
                    <ItemStyle HorizontalAlign="Center" Width="20%" />
                </asp:BoundField>
                  <asp:BoundField  DataField="fl_name" HeaderText="樓層名稱" >
                    <ItemStyle HorizontalAlign="Center" Width="25%" />
                </asp:BoundField>
                 <asp:BoundField  DataField="rnum_name" HeaderText="房間(區域)名稱" >
                    <ItemStyle HorizontalAlign="Center" Width="25%" />
                </asp:BoundField>
                      <asp:BoundField  DataField="Anycnt" HeaderText="數量" >
                    <ItemStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:LinkButton ID="LBItCodeDetail" runat="server" CausesValidation="False" CommandArgument=<%#Eval("element_code")  + "," + Eval("TypeName") %> CommandName="AnyDetail" text="明細" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="10%" />
                </asp:TemplateField> 
            </Columns>
        </asp:GridView>
        <asp:Label ID="lbEmptyData" runat="server" Font-Bold="True" ForeColor="Red" />  
            <asp:Button ID="btnPrint" runat="server" Visible="false" />
        </center>

             <div id="Div_grid" runat="server" style="border-right: lightslategray 2px solid;
            border-top: lightslategray 2px solid; display: block; z-index: 3; left: 194px;
            border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
            position: absolute; top: 824px; height: 150pt; background-color: white" visible="false">
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
            border-top: lightslategray 2px solid; display: block; z-index: 3; left: 428px;
            border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
            position: absolute; top: 825px; height: 150pt; background-color: white" visible="false">
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
    </form>
</body>
</html>