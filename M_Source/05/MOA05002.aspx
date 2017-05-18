<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA05002.aspx.vb" Inherits="Source_05_MOA05002" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>會客洽公查詢</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.7.2.min.js" type="text/javascript"></script>
    <link href="../../css/jquery.datepick.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.10.2.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick-zh-TW.js" type="text/javascript"></script>

    <script type="text/javascript" language="javascript">
        document.oncontextmenu = new Function("return false");

        $(function () {
            $("#nRECDATE1").datepick({ formats: 'yyyy/m/d', defaultDate: $("#nRECDATE1").val(), showTrigger: '#calImg' });
            $("#nRECDATE2").datepick({ formats: 'yyyy/m/d', defaultDate: $("#nRECDATE2").val(), showTrigger: '#calImg' });
        }); 
    </script>
</head>
<body background="../../Image/main_bg.jpg">
    <form id="form1" runat="server">     
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="會客洽公查詢" Width="100%"></asp:Label>
            </td></tr>
        </table>
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
            <tr><td noWrap width="5%">
                <asp:Label ID="Label1" runat="server" Text="部門名稱：" CssClass="form"></asp:Label></td><td width="30%">
                    <asp:DropDownList id="PAUNIT"
                        DataSourceID="SqlDataSource1"
                        DataValueField="ORG_UID"
                        DataTextField="ORG_NAME"
                        runat="server" AutoPostBack="True">
                    </asp:DropDownList></td>
                    <td noWrap width="10%">
                <asp:Label ID="Label2" runat="server" Text="被會人：" CssClass="form"></asp:Label></td><td width="10%">
                    <asp:DropDownList ID="PAIDNO" runat="server" DataSourceID="SqlDataSource3" DataTextField="emp_chinese_name" DataValueField="employee_id">
                    </asp:DropDownList></td><td noWrap width="10%">
                <asp:Label ID="Label5" runat="server" Text="被會人：" CssClass="form" BorderStyle="None"></asp:Label></td><td width="10%">
                <asp:TextBox ID="PANAME" runat="server" Width="60px"></asp:TextBox>
            </td><td noWrap width="5%">
                &nbsp;<asp:Label ID="Label9" runat="server" Text="事由：" CssClass="form"></asp:Label></td><td width="20%">
                <asp:TextBox ID="nREASON" runat="server" Width="120px"></asp:TextBox></td></tr>
            <tr><td noWrap width="5%">
                <asp:Label ID="Label3" runat="server" Text="日期：" CssClass="form"></asp:Label></td><td width="30%">
                <asp:TextBox ID="nRECDATE1" runat="server" Width="60px" OnKeyDown="return false"></asp:TextBox>
                <div style="display: none;">
	                <img id="calImg" src="../../Image/calendar.gif" alt="選擇日期" />
                </div>
                <asp:Label ID="Label4" runat="server" Text="~" CssClass="form"></asp:Label>
                <asp:TextBox ID="nRECDATE2" runat="server" Width="60px" OnKeyDown="return false"></asp:TextBox>
                </td><td noWrap width="10%">
                <asp:Label ID="Label10" runat="server" Text="來賓姓名：" CssClass="form"></asp:Label></td><td width="10%">
                <asp:TextBox ID="nName" runat="server" Width="60px"></asp:TextBox></td><td noWrap width="10%">
                <asp:Label ID="Label8" runat="server" Text="來賓服務單位：" CssClass="form"></asp:Label></td><td width="10%">
                <asp:TextBox ID="nService" runat="server" Width="60px"></asp:TextBox></td><td width="5%">
		    </td>
		    <td width="20%">
                <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" title="查詢" ToolTip="查詢"/>
                <asp:ImageButton ID="ImgClearAll" runat="server" ImageUrl="~/Image/clearall.gif" title="畫面清除" ToolTip="清除"/>
                    <asp:ImageButton ID="imbOutSearch" runat="server" ImageUrl="~/Image/ReadAD.gif" 
                        ToolTip="查詢未匯出資料" />
                    <asp:ImageButton ID="imbExport" runat="server" 
                        ImageUrl="~/Image/mend.gif" ToolTip="匯出資料" Visible="False" onClientClick="ShowProgressBar();"/>
                </td>
		    </tr>
		</table>
	    <asp:GridView ID="GridView1" runat="server" EmptyDataText="查無任何資料" 
            AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" 
            DataKeyNames="EFORMSN" DataSourceID="SqlDataSource2" Width="100%" 
            CellPadding="4" ForeColor="#333333" GridLines="None" 
            EnableModelValidation="True">
            <Columns>
                <asp:BoundField DataField="nRECDATE" HeaderText="會客日期" SortExpression="nRECDATE" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                </asp:BoundField>
                <asp:BoundField DataField="nSTARTTIME" HeaderText="時間" SortExpression="nSTARTTIME" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="nREASON" HeaderText="事由" SortExpression="nREASON" >
                    <HeaderStyle HorizontalAlign="Center" Width="32%" />
                </asp:BoundField>
                <asp:BoundField DataField="PANAME" HeaderText="被會人" SortExpression="PANAME" >
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="nRECROOM" HeaderText="會客室" SortExpression="nRECROOM">
                    <HeaderStyle HorizontalAlign="Center" Width="15%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="狀態" SortExpression="PENDFLAG">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("PENDFLAG") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# FunStatus("PENDFLAG") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="8%" />
                </asp:TemplateField>
                <asp:CommandField ShowSelectButton="True" ButtonType="Image" SelectText="詳細資料" SelectImageUrl="~/Image/List.gif" >
                    <ItemStyle HorizontalAlign="Center"  />
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:CommandField>
                <asp:BoundField DataField="EFORMSN" HeaderText="EFORMSN" ReadOnly="True" 
                    SortExpression="EFORMSN" />
            </Columns>
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="">
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT * FROM [ADMINGROUP] WHERE 1=2">
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT [employee_id], [emp_chinese_name], [ORG_UID] FROM [EMPLOYEE] WHERE 1=2 ORDER BY [emp_chinese_name]">
            <SelectParameters>
            <asp:ControlParameter ControlID="PAUNIT" Name="ORG_NAME" PropertyName="SelectedValue" Type="String" /></SelectParameters>
        </asp:SqlDataSource>

        
<div id="divProgress" style="text-align:center; display: none; position: fixed; top: 50%;  left: 50%;" > 
     <asp:Image ID="imgLoading" runat="server" ImageUrl="~/Image/loading.gif" />          
     <br /> 
     <font color="#1B3563" size="2px">資料處理中</font> 
 </div> 
 <div id="divMaskFrame" style="background-color: #F2F4F7; display: none; left: 0px; 
     position: absolute; top: 0px;"> 
 </div>          

    </form>
    <script type="text/javascript" language="javascript">
    
// 顯示讀取遮罩 
 function ShowProgressBar() { 
      displayProgress(); 
      displayMaskFrame(); 
 } 
 // 隱藏讀取遮罩 
 function HideProgressBar() { 
     var progress = $('#divProgress'); 
     var maskFrame = $("#divMaskFrame"); 
     progress.hide(); 
     maskFrame.hide(); 
 } 
 // 顯示讀取畫面 
 function displayProgress() { 
     var w = $(document).width(); 
     var h = $(window).height(); 
     var progress = $('#divProgress'); 
     progress.css({ "z-index": 999999, "top": (h / 2) - (progress.height() / 2), "left": (w / 2) - (progress.width() / 2) }); 
     progress.show(); 
 } 
 // 顯示遮罩畫面 
 function displayMaskFrame() { 
     var w = $(window).width(); 
     var h = $(document).height(); 
     var maskFrame = $("#divMaskFrame"); 
     maskFrame.css({ "z-index": 999998, "opacity": 0.7, "width": w, "height": h }); 
     maskFrame.show(); 
 } 
    </script>
</body>
</html>
