<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA08008.aspx.vb" Inherits="M_Source_08_MOA08008" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>機密資訊複(影)印資料申請單</title>
    <script type="text/javascript" src="../../script/jquery-1.7.2.min.js"></script>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript">
        $(document).ready(function () {
            var imgDate1 = $('#<%=ImgDate1.ClientID %>').offset();
            var imgDate2 = $('#<%=ImgDate2.ClientID %>').offset();
            var lbSecurityRange = $('#<%=lb_Security_Range.ClientID %>').offset()
            //設定發文時間Calender顯示座標
            $('#<%=hidImgDate1Left.ClientID %>').val(imgDate1.left);
            $('#<%=hidImgDate1Top.ClientID %>').val(imgDate1.top);
            //設定複印時間Calender顯示座標
            $('#<%=hidImgDate2Left.ClientID %>').val(imgDate2.left);
            $('#<%=hidImgDate2Top.ClientID %>').val(imgDate2.top);
            //設定保密期限/保密條件顯示座標
            $('#<%=hidSecurity_RangeLeft.ClientID %>').val($('#mon_form').width() / 2 - 240);
            $('#<%=hidSecurity_RangeTop.ClientID %>').val($('#mon_form').height() / 2 - 55);
            checkAndSetTotalSheet(); //計算及顯示合計複印張數欄位
            $('#<%=lb_PrintTotal.ClientID %>').attr("readonly", "readonly"); //合計複印張數欄位設定為 readOnly
        });

        function checkWords(obj) {
            var words = $(obj).val();
            if (words.length > 255) {
                alert("不可超過255個字數");
                $(obj).val(words.substring(0, 255));
                $(obj).focus();
                return false;
            }
        }

        function checkAndSetTotalSheet() { //計算及顯示合計複印張數欄位
            var orgNum = $('#<%=lb_Org_Num.ClientID %>').val();
            var printSetNum = $('#<%=lb_PrintSet_Cnt.ClientID %>').val();
            var totalNum = $('#<%=lb_PrintTotal.ClientID %>');
            var re = /^[0-9]+$/;
            if (orgNum.match(re) != null && printSetNum.match(re) != null) {
                totalNum.val(Number(orgNum) * Number(printSetNum));
            }
        }

        function sumCalculate(obj) { //檢驗原件張數及每份複印張數欄位
            var re = /^[0-9]+$/;
            if ($(obj).val().match(re) == null) {
                alert("請輸入數字");
                $(obj).val("");
                $(obj).focus();
                return false;
            }
            checkAndSetTotalSheet(); //計算及顯示合計複印張數欄位 
        }

        function checkIllegalChar(obj) {
            var re = /^[A-Z a-z 0-9]+$/;
            if ($(obj).val().match(re) == null) {
                alert("請輸入英文字母或數字");
                $(obj).val("");
                $(obj).focus();
                return false;
            }
        }
    </script>
</head>
<body style="font-size: large">
<form name="mon_form" id="mon_form" runat="server" >
  <asp:HiddenField ID="hidImgDate1Left" runat="server" />
  <asp:HiddenField ID="hidImgDate1Top" runat="server" />
  <asp:HiddenField ID="hidImgDate2Left" runat="server" />
  <asp:HiddenField ID="hidImgDate2Top" runat="server" />
  <asp:HiddenField ID="hidSecurity_RangeLeft" runat="server" />
  <asp:HiddenField ID="hidSecurity_RangeTop" runat="server" />
  <div id="mainDiv">
  <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
      <tr><td align="center" style="background-color:#6499CD;color:#FFFFFF;font:bold 15px 'Verdana, Arial, Helvetica, sans-serif';">
          <asp:Label ID="lb_No1ORGName" runat="server" Text = "一級單位全銜" /> 機密資訊複(影)印資料申請單
      </td></tr>
  </table>
  <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
  <tr>
      <td nowrap="nowrap" width="40%" class="form" style="font-size:12pt" align="right">
          <span style="color:Red;">*</span><asp:Label ID="lbSubject" runat="server" Text="主旨、簡由或(資料名稱)：" /></td>
      <td nowrap="nowrap" width="60%" class="form" style="font-size:12pt">
          <asp:TextBox ID="tb_Subject" runat="server" MaxLength="255" /></td>
  </tr>
  <tr>
      <td nowrap="nowrap" width="40%" class="form" style="font-size:12pt" align="right">
          <span style="color:Red;">*</span><asp:Label ID="lbSignDateTime" runat="server" Text="發文時間：" /></td>
      <td nowrap="nowrap" width="60%" class="form" style="font-size:12pt">
          <asp:Label ID="lb_SignDT" width = "100px" runat="server" OnKeyDown="return false" />
            <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" /></td>
  </tr>
  <tr>
      <td nowrap="nowrap" width="40%" class="form" style="font-size:12pt" align="right">
          <span style="color:Red;">*</span><asp:Label ID="lbSecurity_No" runat="server" Text="發文字號：" /></td>
      <td nowrap="nowrap" width="60%" class="form" style="font-size:12pt">
          <asp:TextBox ID="tb_Security_No1" runat="server" MaxLength="4" Width="73px" />字第<asp:TextBox 
              ID="tb_Security_No2" runat="server" MaxLength ="10" Width="94px" />號</td>
  </tr>
  <tr>
      <td nowrap="nowrap" width="40%" class="form" style="font-size:12pt" align="right">
          <span style="color:Red;">*</span><asp:Label ID="lbSecurity_Level" runat="server" Text="機密等級：" /></td>
      <td nowrap="nowrap" width="60%" class="form" style="font-size:12pt">
          <asp:Label ID="lb_Security_Level" runat="server" Visible ="false" Text ="密" />
		  <asp:DropDownList ID="ddl_Security_Level" runat="server" AutoPostBack="True" 
              style="font-size:12pt;" CssClass="form">
              <asp:ListItem Value="2" Selected = "True">密</asp:ListItem>
              <asp:ListItem Value="3">機密</asp:ListItem>
              <asp:ListItem Value="4">極機密</asp:ListItem>
              <asp:ListItem Value="5">絕對機密</asp:ListItem>
          </asp:DropDownList></td>
  </tr>
  <tr>
      <td nowrap="nowrap" width="40%" class="form" style="font-size:12pt" align="right">
          <span style="color:Red;">*</span><asp:Label ID="lbSecurity_Type" runat="server" Text="機密屬性：" /></td>
      <td nowrap="nowrap" width="60%" class="form" style="font-size:12pt">
          <asp:Label ID="lbSecurity_Level_2" runat="server" width="120px" Font-Size="12pt" />
            <asp:DropDownList ID="ddl_Security_Type" runat="server" Visible="False" 
                style="font-size:12pt;" CssClass="form" AutoPostBack="True">
                <asp:ListItem Value="1">國家機密</asp:ListItem>
                <asp:ListItem Value="2">軍事機密</asp:ListItem>
                <asp:ListItem Value="3">國防秘密</asp:ListItem>
                <asp:ListItem Value="4">國家機密亦屬軍事機密</asp:ListItem>
                <asp:ListItem Value="5">國家機密亦屬國防秘密</asp:ListItem>
            </asp:DropDownList></td>
  </tr>
  <tr>
      <td nowrap="nowrap" width="40%" class="form" style="font-size:12pt" align="right">
          <span style="color:Red;">*</span><asp:LinkButton ID="lnkbtn_Security_Range" runat="server" style="text-decoration:none;color:blue;">保密期限/保密條件:</asp:LinkButton></td>
      <td nowrap="nowrap" width="60%" class="form" style="font-size:12pt">
          <asp:Label ID="lb_Security_Range" runat="server" OnKeyDown="return false" /></td>
  </tr>
  <tr>
      <td nowrap="nowrap" width="40%" class="form" style="font-size:12pt" align="right">
          <asp:Label ID="lbProduceUnit" runat="server" Text="產製單位：" /></td>
      <td nowrap="nowrap" width="60%" class="form" style="font-size:12pt">
          <asp:TextBox ID="txtProduceUnit" runat="server"></asp:TextBox></td>
  </tr>
  <tr>
      <td nowrap="nowrap" width="40%" class="form" style="font-size:12pt" align="right">
          <asp:Label ID="lbAgreeTimeOrNumber" runat="server" Text="同意複(影)印時間/文號：" /></td>
      <td nowrap="nowrap" width="60%" class="form" style="font-size:12pt">
          <asp:TextBox ID="txtAgreeTimeOrNumber" runat="server"></asp:TextBox></td>
  </tr>
  <tr>
      <td nowrap="nowrap" width="40%" class="form" style="font-size:12pt" align="right">
          <asp:Label ID="lbAgreeSuperior" runat="server" Text="同意複(影)印權責長官級職姓名：" /></td>
      <td nowrap="nowrap" width="60%" class="form" style="font-size:12pt">
          <asp:TextBox ID="txtAgreeSuperior" runat="server"></asp:TextBox></td>
  </tr>
  <tr>
      <td nowrap="nowrap" width="40%" class="form" style="font-size:12pt" align="right">
          <span style="color:Red;">*</span><asp:Label ID="lbPurpose" runat="server" Text="用途：" /></td>
      <td nowrap="nowrap" width="60%" class="form" style="font-size:12pt">
        <asp:RadioButton ID="RB1" runat="server" GroupName ="Security_Type_RBGroup" Text ="呈閱" AutoPostBack="True" />
        <asp:RadioButton ID="RB2" runat="server" GroupName ="Security_Type_RBGroup" Text ="分會、辦" AutoPostBack="True" />
        <asp:RadioButton ID="RB3" runat="server" GroupName ="Security_Type_RBGroup" Text ="作業用" AutoPostBack="True" />
        <asp:RadioButton ID="RB4" runat="server" GroupName ="Security_Type_RBGroup" Text ="歸檔" AutoPostBack="True" />
        <asp:RadioButton ID="RB5" runat="server" GroupName ="Security_Type_RBGroup" Text ="隨文分發" AutoPostBack="True" />
        <asp:RadioButton ID="RB6" runat="server" GroupName ="Security_Type_RBGroup" Text ="會議分發" AutoPostBack="True" /> <br />   
        <asp:RadioButton ID="RB7" runat="server" GroupName ="Security_Type_RBGroup" Text ="其他" AutoPostBack="True" />    
        <asp:TextBox ID="tb_Purpose_Other" runat="server" MaxLength ="50" Enabled="false" />
      </td>
  </tr>
  <tr>
      <td nowrap="nowrap" width="40%" class="form" style="font-size:12pt" align="right">
          <span style="color:Red;">*</span><asp:Label ID="lbPrinter_Datetime" runat="server" Text="複印時間：" /></td>
      <td nowrap="nowrap" width="60%" class="form" style="font-size:12pt">
          <asp:Label ID="lb_Print_DT" runat="server" OnKeyDown="return false" 
              Width="100px"></asp:Label>
            <asp:ImageButton ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" /></td>
  </tr>
  <tr>
      <td nowrap="nowrap" width="40%" class="form" style="font-size:12pt" align="right">
          <span style="color:Red;">*</span><asp:Label ID="lbPrinter_Num" runat="server" Text="複(影)印機浮水印暗記編號：" /></td>
      <td nowrap="nowrap" width="60%" class="form" style="font-size:12pt">
          <asp:TextBox ID="lb_Printer_Num" runat="server" MaxLength="6" onchange="checkIllegalChar(this);"></asp:TextBox></td>
  </tr>
  <tr>
      <td nowrap="nowrap" width="40%" class="form" style="font-size:12pt" align="right">
          <span style="color:Red;">*</span><asp:Label ID="lbOri_sheet" runat="server" Text="原件張數：" /></td>
      <td nowrap="nowrap" width="60%" class="form" style="font-size:12pt">
          <asp:TextBox ID="lb_Org_Num" runat="server" onchange="sumCalculate(this);"></asp:TextBox></td>
  </tr>
  <tr>
      <td nowrap="nowrap" width="40%" class="form" style="font-size:12pt" align="right">
          <span style="color:Red;">*</span><asp:Label ID="lbCopy_sheet" runat="server" Text="每份複印張數：" /></td>
      <td nowrap="nowrap" width="60%" class="form" style="font-size:12pt">
          <asp:TextBox ID="lb_PrintSet_Cnt" runat="server" onchange="sumCalculate(this);"></asp:TextBox></td>
  </tr>
  <tr>
      <td nowrap="nowrap" width="40%" class="form" style="font-size:12pt" align="right">
          <asp:Label ID="lbTotal_sheet" runat="server" Text="合計複印張數：" /></td>
      <td nowrap="nowrap" width="60%" class="form" style="font-size:12pt">
          <asp:TextBox ID="lb_PrintTotal" runat="server" AutoPostBack="True"></asp:TextBox>
      </td>
  </tr>
  <tr>
      <td nowrap="nowrap" width="40%" class="form" style="font-size:12pt" align="right">
          <asp:Label ID="lbSheet_ID" runat="server" Text="複(影)印張數流水號：" /></td>
      <td nowrap="nowrap" width="60%" class="form" style="font-size:12pt">
          <asp:TextBox ID="lb_sheetID" MaxLength="30" runat="server" onchange="checkIllegalChar(this);"></asp:TextBox></td>
  </tr>
  <tr>
      <td nowrap="nowrap" width="40%" class="form" style="font-size:12pt" align="right" valign="top">
          <asp:Label ID="lbMemo" runat="server" Text="附註：" /></td>
      <td nowrap="nowrap" width="60%" class="form" style="font-size:12pt">
          <textarea ID="tb_Memo" runat="server" rows="5" cols="35" maxlength="10" onchange="checkWords(this);" /></td>
  </tr>
  </table>
</div>
    <center>
    <asp:ImageButton ID="ImgSavaPrint" runat="server" ImageUrl ="~/Image/apply.gif" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <a href="javascript:void(0);" onclick="javascript:window.location.href='MOA08007.aspx';"><img alt="回上頁" src="../../Image/backtop.gif" border="0" /></a>
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

 <div id="Div1" runat="server" style="border-right: lightslategray 2px solid;
            border-top: lightslategray 2px solid; display: block; z-index: 3; left: 194px;
            border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
            position: absolute; top: 824px; height: 150pt; background-color: white" visible="false">
            <asp:Calendar ID="Calendar2" runat="server" BackColor="White" BorderColor="#3366CC"
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
            <asp:Button ID="Button1" runat="server" Text="關閉" Width="220px" /></div>
 <div id="Div_Security" runat="server" style="border-right: lightslategray 2px solid;
            border-top: lightslategray 2px solid; display: block; z-index: 3; left: 550px;
            border-left: lightslategray 2px solid; width: 480px; border-bottom: lightslategray 2px solid;
            position: absolute; top: 350px; height: 115pt; background-color: white" visible="false">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse; border-color:Black; font-family: 標楷體;" width="480px">
	<tr>
		<td style="width:150px; font-size:12pt;" class="form">解密條件：</td>
		<td colspan="3" class="form" style="font-size:12pt">
            <asp:TextBox ID="tb_conditions" runat="server" MaxLength ="30" Width="340px" /></td>
	</tr>
	<tr>
		<td style="width:150px; font-size:12pt;" align="left" class="form">
            <asp:RadioButton ID="RBSecurity1" runat="server" Text ="本文件保密至" GroupName="Security1" AutoPostBack="True" /></td>
        <td colspan="3" align="left" class="form" style="font-size:12pt" >
            &nbsp;西元
        <asp:DropDownList ID="ddl_Year" runat="server" Enabled="false" />年<asp:DropDownList 
                ID="ddl_Month" runat="server" Enabled="false" AutoPostBack="True" CssClass="form" style="font-size:12pt" />
            月<asp:DropDownList ID="ddl_Day" runat="server" Enabled="false" CssClass="form" style="font-size:12pt" />日</td>
	</tr>
	<tr>
		<td style="width:100px; font-size:12pt;" class="form">&nbsp;</td>
		<td align="left" colspan="3" class="form" style="font-size:12pt">
            <asp:RadioButton ID="RB11" runat="server" Text="解除密等" Enabled="false" GroupName="Security2" Checked="True" />
        </td>
	</tr>
	<tr>
		<td style="width:100px; font-size:12pt;" class="form">&nbsp;</td>
		<td colspan ="2" align ="left" style="border-right-color:#FFFFFF; font-size:12pt;" class="form"><asp:RadioButton ID="RB12" 
                Enabled="false" width="116px" runat="server" Text="變更密等為：" 
                GroupName="Security2" />&nbsp;<asp:DropDownList ID="ddl_Security_Level2" 
                 Enabled="false" runat="server" CssClass="form" style="font-size:12pt" /></td>
   
	</tr>
	<tr>
		<td colspan="4" align="left" class="form" style="font-size:12pt">
            <asp:RadioButton ID="RBSecurity2" runat="server" Text="本文件永久保密" Checked="true" GroupName ="Security1" AutoPostBack="True" /></td>
	</tr>
</table>
<table border="0" width="100%"><tr><td align="center">
<asp:Button ID="Security_OK" runat="server" Text="確定" Width="49%" /> <asp:Button ID="Security_Close" runat="server" Text="關閉" Width="49%" />
</td></tr></table>
</div>
    </form>
    
</body>
</html>
