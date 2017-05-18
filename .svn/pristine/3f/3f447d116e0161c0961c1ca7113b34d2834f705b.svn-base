<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA08012.aspx.vb" Inherits="M_Source_08_MOA08012" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>影印紀錄呈核</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.7.2.min.js" type="text/javascript"></script>
    <script type="text/JavaScript">
        $(document).ready(function () {
            //將資料ID隱藏欄位字串拆解為資料ID字串陣列
            var selectedLog_GuidArray = $('#<%=selectedLog_Guid.ClientID %>').val().split(',');
            if (selectedLog_GuidArray.length > 0) {
                $(":checkbox").each(function () { //分頁轉換用,掃描頁面checbox,若ID於資料ID隱藏欄位字串內,則勾選該checkbox.
                    for (var index = 0; index < selectedLog_GuidArray.length; index++) {
                        if (selectedLog_GuidArray[index] == this.id)
                            this.checked = true;
                    }
                });
            }
        });
        function checkData() {
            var re = /</;
            var beginDate = document.getElementById("<%=txtPrint_DateBegin.ClientID %>").value;
            var endDate = document.getElementById("<%=txtPrint_DateEnd.ClientID %>").value;
            if (re.test(beginDate)) {
                alert("您輸入的複印時間開始日期有包含危險字元！");
                beginDate = "";
                return false;
            }
            if (re.test(endDate)) {
                alert("您輸入的複印時間結束日期有包含危險字元！");
                endDate = "";
                return false;
            }
            if (re.test(document.getElementById("<%=txtFile_Name.ClientID %>").value)) {
                alert("您輸入的複印資料名稱有包含危險字元！");
                document.getElementById("<%=txtFile_Name.ClientID %>").value = "";
                return false;
            }
            if (re.test(document.getElementById("<%=txtUse_For.ClientID %>").value)) {
                alert("您輸入的用途有包含危險字元！");
                document.getElementById("<%=txtUse_For.ClientID %>").value = "";
                return false;
            }
            if (re.test(document.getElementById("<%=txtPrint_Name.ClientID %>").value)) {
                alert("您輸入的姓名有包含危險字元！");
                document.getElementById("<%=txtPrint_Name.ClientID %>").value = "";
                return false;
            }
            if (re.test(document.getElementById("<%=txtPrint_Num.ClientID %>").value)) {
                alert("您輸入的人員帳號流水號有包含危險字元！");
                document.getElementById("<%=txtPrint_Num.ClientID %>").value = "";
                return false;
            }
            if (beginDate != "" && endDate != "") {
                if (new Date(beginDate) > new Date(endDate)) {
                    alert("複印時間結束日期不可大於開始日期！");
                    document.getElementById("<%=txtPrint_Num.ClientID %>").value = "";
                    return false;
                }
            }
            return true;
        }

        function selectAll(obj) { //checkcox全選及取消
            $(":checkbox").each(function () {
                if (this.id != "cbAll") {
                    if (obj.checked) addSelected(this.id)
                    else removeSelected(this.id)

                    this.checked = obj.checked;
                }
            });
        }

        function checkSelected() { //檢查是否有選擇呈核資料
            if ($('#<%=selectedLog_Guid.ClientID %>').val() == "") {
                alert("請勾選欲呈核資料");
                return false;
            }
            return true;
        }

        function checkBoxSelected(cbObj) { //checkbox選擇事件
            if (cbObj.id != "cbAll") {
                if (cbObj.checked) addSelected(cbObj.id); //勾選則增加該筆資料ID至隱藏欄位字串
                else removeSelected(cbObj.id); //反之則於隱藏欄位字串刪除該筆資料ID
            }
        }

        function addSelected(dataId) { //增加資料ID至隱藏欄位字串
            var strSelectedLog_Guid = $('#<%=selectedLog_Guid.ClientID %>').val();
            if (strSelectedLog_Guid == "") //資料ID隱藏欄位字串若為空字串,則直接加入該筆資料ID
                strSelectedLog_Guid = dataId;
            else  //反之則前面多增加逗點,資料ID隱藏欄位字串格式: Id1,Id2,Id3...
                strSelectedLog_Guid = strSelectedLog_Guid + "," + dataId;
            $('#<%=selectedLog_Guid.ClientID %>').val(strSelectedLog_Guid)
        }

        function removeSelected(dataId) { //於隱藏欄位字串刪除資料ID
            var hidFieldObj = $('#<%=selectedLog_Guid.ClientID %>');
            //將資料ID隱藏欄位字串拆解為資料ID字串陣列
            var selectedLog_GuidArray = hidFieldObj.val().split(',');
            if (selectedLog_GuidArray.length > 0) {
                var strTemp = "";
                for (var i = 0; i < selectedLog_GuidArray.length; i++) {
                    if (selectedLog_GuidArray[i] != dataId) { //非刪除目標字串則存入暫存字串
                        if (strTemp == "") strTemp = selectedLog_GuidArray[i];
                        else strTemp = strTemp + "," + selectedLog_GuidArray[i];
                    }
                }
                hidFieldObj.val(strTemp); //暫存字串填入資料ID字串隱藏欄位
            }
        }

        function clearAll() { //清除查詢欄位資料
            $('#<%=txtPrint_DateBegin.ClientID %>').val('');
            $('#<%=txtPrint_DateEnd.ClientID %>').val('');
            $('#<%=txtFile_Name.ClientID %>').val('');
            $('#<%=txtUse_For.ClientID %>').val('');
            $('#<%=txtPrint_Num.ClientID %>').val('');
            $('#<%=txtPrint_Name.ClientID %>').val('');
            $('#<%=ddlORG_UID.ClientID %>').val('');
            $('#<%=ddlSecurity_Status.ClientID %>').val('');
        }
    </script>
    <style type="text/css">
        .style1
        {
            font: 13px Verdana, Arial, Helvetica, sans-serif;
            color: #666666;
            text-decoration: none;
            width: 17%;
        }
    </style>
    </head>
<body>
    <form id="form1" runat="server">
    <asp:HiddenField ID="selectedLog_Guid" runat="server" />
    <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
        <tr><td align="center" class = "toptitle">
        影印紀錄呈核</td></tr>
    </table>
    <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
            <tr>
                <td nowrap="nowrap" width="5%" class="form" align = "left">
                    複印時間：
                </td>
                <td nowrap="nowrap" width="20%" class="form"  align = "left">
                    <asp:TextBox ID="txtPrint_DateBegin" runat="server" Width="80px" 
                        OnKeyDown="return false" ReadOnly="True" ></asp:TextBox>
                    <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />
                    ~
                    <asp:TextBox ID="txtPrint_DateEnd" runat="server" Width="80px" 
                        OnKeyDown="return false" ReadOnly="True"></asp:TextBox>
                    <asp:ImageButton ID="ImgDate2" runat="server" ImageUrl="~/Image/calendar.gif" />
                </td>
                <td nowrap="nowrap" width="8%" class="form"  align = "left">
                    複印資料名稱：
                </td>
                <td nowrap="nowrap" class="style1" align = "left">
                    <asp:TextBox ID="txtFile_Name" runat="server"></asp:TextBox>
                </td>
              
                <td nowrap="nowrap" width="8%" class="form"  align = "left">
                    用途：
                </td>
                <td nowrap="nowrap" width="20%" class="form">
                    <asp:TextBox ID="txtUse_For" runat="server"></asp:TextBox>
                </td>
                </tr>
                <tr>
                 <td nowrap="nowrap" width="6%" class="form"  align = "left">
                    使用單位：
                </td>
                 <td nowrap="nowrap" width="15%" class="form" align = "left">
                     <asp:DropDownList ID="ddlORG_UID" runat="server">
                     </asp:DropDownList>
                </td>
                <td nowrap="nowrap" width="8%" class="form"  align = "left">
                    密等：
                </td>
                <td nowrap="nowrap" class="style1"  align = "left">
                    <asp:DropDownList ID="ddlSecurity_Status" runat="server">
                        <asp:ListItem Value=""></asp:ListItem>
                        <asp:ListItem Value="1">普通</asp:ListItem>
                        <asp:ListItem Value="2">密</asp:ListItem>
                        <asp:ListItem Value="3">機密</asp:ListItem>
                        <asp:ListItem Value="4">極機密</asp:ListItem>
                        <asp:ListItem Value="5">絕對機密</asp:ListItem>
                    </asp:DropDownList>
                </td> 
                <td nowrap="nowrap" width="8%" class="form"  align = "left">
                    流水號：
                </td>
                <td nowrap="nowrap" width="6%" class="form"  align = "left">
                    <asp:TextBox ID="txtPrint_Num" runat="server"></asp:TextBox>
                </td> 
            </tr>
            <tr>
                <td nowrap="nowrap" width="5%" class="form" align = "left">
                    姓名：
                </td>
                <td nowrap="nowrap" width="10%" class="form" align = "left">
                    <asp:TextBox ID="txtPrint_Name" runat="server"></asp:TextBox>
                </td>
                 <td nowrap="nowrap" width="5%" class="form" align = "left">
                </td>
                <td nowrap="nowrap" class="style1" align = "left">
                    &nbsp;</td>
               
                <td class="form" colspan ="2" align="center">
                    <asp:Button ID="btnSubmit" runat="server" Text="呈核" OnClientClick="return checkSelected();" />&nbsp;&nbsp;
                    <asp:Button ID="btnSearch" runat="server" Text="查詢" OnClientClick="return checkData();" />&nbsp;&nbsp;
                    <input type="button" onclick="clearAll();" value="清除查詢條件" />
                </td>
            </tr>
            <tr>
                <td colspan ="8" align ="left"><asp:Label ID="ErrMsg" runat="server" ForeColor='Red' Font-Bold ="true" /></td>
            </tr>
        </table>

        <asp:GridView ID="gvP_08" runat="server" EmptyDataText="查無任何資料" 
            AutoGenerateColumns="False" AllowPaging="True"
         AllowSorting="True" Width="100%" 
            CellPadding="10" ForeColor="#333333" GridLines="None" 
        EnableModelValidation="True" HorizontalAlign="Center">

            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" 
                HorizontalAlign="Center" VerticalAlign="Middle" />
            <Columns>
                <asp:TemplateField HeaderText="&lt;input type='checkbox' id='cbAll' onclick='selectAll(this);' /&gt;">
                    <ItemTemplate>
                        <input type="checkbox" id="<%# Eval("Log_Guid") %>" onclick="checkBoxSelected(this);" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="6.5%"/>
                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:TemplateField>
                <asp:BoundField DataField="File_Name" HeaderText="複印資料名稱" >
                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="8.5%"/>
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="Print_Date" HeaderText="複印時間" >
                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="8.5%"/>
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="ORG_NAME" HeaderText="使用單位" >
                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="8.5%"/>
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="Print_Name" HeaderText="姓名" >
                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="8.5%"/>
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="密等">
                    <ItemTemplate>
                        <%# displaySecurityStatus(Eval("Security_Status"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="8.5%"/>
                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:TemplateField>
                <asp:BoundField DataField="Total_sheet" HeaderText="申請張數" >
                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="8.5%"/>
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="實際張數">
                    <ItemTemplate>
                        <%# displayCopyDetail(Eval("Copy_A3M"),Eval("Copy_A4M"),Eval("Copy_A3C"),Eval("Copy_A4C"),Eval("Scan"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="8.5%"/>
                </asp:TemplateField>
                <asp:BoundField DataField="Use_For" HeaderText="用途" >
                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="8.5%"/>
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="Print_Num" HeaderText="流水號" >
                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="8.5%"/>
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="批核">
                    <ItemTemplate>
                        <%# ShowApproveStatus(Eval("VerifyRequesterID"), Eval("ApprovedByID"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" Width="8.5%"/>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="影印明細">
                    <ItemTemplate>
                        <asp:HiddenField ID="HF_Status" runat="server" Value = '<%# Eval("Status")%>' />
                        <asp:ImageButton ID="ImgDetail" runat="server" CausesValidation="False" CommandArgument ='<%# Eval("Log_Guid") %>' CommandName="Detail" ImageUrl="~/Image/List.gif" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="8.5%"/>
                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:TemplateField>
            </Columns>
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                <EmptyDataRowStyle ForeColor="DarkBlue" HorizontalAlign="Center" />
        </asp:GridView>

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
