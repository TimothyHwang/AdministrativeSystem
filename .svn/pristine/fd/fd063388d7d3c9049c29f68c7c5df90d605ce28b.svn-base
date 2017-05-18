
<OBJECT id="SmfCom" name="SmfCom"
classid="clsid:583EE717-E8D2-4EB2-B59A-82BB5B7D216E" 
type="application/x-oleobject"       
style="display: none"> 
</OBJECT>

<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA08001.aspx.vb" Inherits="M_Source_08_MOA08001" %>

<%@ Register Src="../90/FlowRoute.ascx" TagName="FlowRoute" TagPrefix="uc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>影印掃瞄申請單</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />    
    
    <script src="../../script/jquery-1.7.2.min.js" type="text/javascript"></script>
    <script src="../../script/printCard.js" type="text/javascript"></script>
    <script type="text/javascript" language="javascript">
        document.oncontextmenu = new Function("return false");
    </script>
    <style type="text/css">
        .style2
        {
            width: 68px;
        }
        .style3
        {
            height: 23px;
            width: 82px;
        }
        .style4
        {
            height: 23px;
            width: 68px;
        }
    </style>
    <script type="text/javascript">
        $(document).ready(function () {
            var varFormType = "<%=read_only %>"; //表單開啟型態 空值:新增表單 非空值:瀏覽表單
            var fileName = $('#<%=tbFile_Name.ClientID %>'); //普通申請單主旨、簡由或(資料名稱)
            var useFor = $('#<%=tbUse_For.ClientID %>');     //普通申請單用途
            var orgSheets = $('#<%=txtOriginalSheets.ClientID %>'); //普通申請單原件張數
            var copySheets = $('#<%=txtCopySheets.ClientID %>');    //普通申請單每份複印張數
            var totalSheets = $('#<%=txtTotalSheets.ClientID %>');  //普通申請單合計複印張數
            totalSheets.attr("readonly", "readonly"); //合計複印張數欄位設為readonly
            var ddlSecurityItemDisplay = $('#<%=ddl_SecurityItem.ClientID %>').css("display"); //機密表單顯示狀態
            if (varFormType != "") { //瀏覽表單
                if ($('#<%=ddl_Security_Status.ClientID %>').val() == 1) { //瀏覽普通表單
                    $('#tbContent').show();
                    $('#tbSecurityContent').hide();
                }
                else { //瀏覽機密表單
                    $('#tbContent').hide();
                    $('#tbSecurityContent').show();
                }
            }
            else if ($('#<%=ddl_Security_Status.ClientID %>').val() == 1) { //申請類別普通
                $('#tbContent').show();
                $('#tbSecurityContent').hide();
                fileName.val('');
                useFor.val('');
                orgSheets.val('');
                copySheets.val('');
                totalSheets.val('');
            }
            else if (ddlSecurityItemDisplay == "inline-block" || ddlSecurityItemDisplay == "inline") { //申請類別密以上
                $('#tbContent').hide();
                $('#tbSecurityContent').show();
                fileName.val($('#<%=lbSubject.ClientID %>').text()); //將機密文件申請單之資料填入影印申請單相對欄位內
                useFor.val($('#<%=lbPurpose.ClientID %>').text());
                orgSheets.val($('#<%=lbOri_sheet.ClientID %>').text());
                copySheets.val($('#<%=lbCopy_sheet.ClientID %>').text());
                totalSheets.val($('#<%=lbTotal_sheet.ClientID %>').text());
            }
            else { //申請類別密以上,但無機密文件申請單
                $('#tbContent').hide();
                $('#tbSecurityContent').hide();
                fileName.val('');
                useFor.val('');
                orgSheets.val('');
                copySheets.val('');
                totalSheets.val('');
            }
        });

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

        function checkAndSetTotalSheet() { //計算及顯示合計複印張數欄位
            var orgSheets = $('#<%=txtOriginalSheets.ClientID %>').val();
            var copySheets = $('#<%=txtCopySheets.ClientID %>').val();
            var totalSheets = $('#<%=txtTotalSheets.ClientID %>');
            var re = /^[0-9]+$/;
            if (orgSheets.match(re) != null && copySheets.match(re) != null) {
                totalSheets.val(Number(orgSheets) * Number(copySheets));
            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
     <div>
        <input type="hidden" id="hidCurrentUserID" name="hidCurrentUserID" value="<%= Session("user_id") %>" />
        <table width="740" border="1" cellspacing="0" cellpadding="5" bgcolor="#ffffff" bordercolor="#6699cc" bordercolorlight="#74a3d6" bordercolordark="#000000" style="LEFT: 20px; TOP: 10px">
        <tr>
            <td valign="bottom" bgcolor="#6699cc" bordercolorlight="#66aaaa" bordercolordark="#ffffff" style="height: 33px; width: 727px;"><font color=white><b>&nbsp;影印掃瞄申請單</b></font>
                <asp:Label ID="lbLog_Guid" runat="server" Visible="False"></asp:Label>
            </td>
        </tr>
        
        <tr>
            <td style="width: 727px">
            
            <fieldset id="tableB" style="width: 740px">
               <table style="width: 740px">
                    <tr>
                        <td style="width: 230px; height: 24px;">
                            <asp:Label ID="Label1" runat="server" Text="填表人單位：" Width="80px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_ORG_NAME_1" runat="server" ForeColor="Black" Width="140px" CssClass="form"></asp:Label></td>
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
                            </asp:DropDownList></td>
                        <td style="width: 210px">
                            <asp:Label ID="Label6" runat="server" Text="級職：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_title_name_2" runat="server" ForeColor="Black" Width="140px" CssClass="form"></asp:Label></td>
                    </tr>                
                </table>  
		    </fieldset> 
                        
            <table border="0" style="width: 740px; height: 57px; color:Red" >
                    <tr>
                        <td style="width: 120px; height: 23px;" align="right">*
                            <asp:Label ID="Label7" runat="server" ForeColor="Black" Text="申請類別：" CssClass="form" ></asp:Label></td>
                        <td class="style4" style="width: 120px;">
                            <asp:DropDownList ID="ddl_Security_Status" runat="server" AutoPostBack ="true">
                                <asp:ListItem Value="1">普通</asp:ListItem>
                                <asp:ListItem Value="2">密</asp:ListItem>
                                <asp:ListItem Value="3">機密</asp:ListItem>
                                <asp:ListItem Value="4">極機密</asp:ListItem>
                                <asp:ListItem Value="5">絕對機密</asp:ListItem>
                            </asp:DropDownList>
                        </td>
                        <td class="style3" style="width: 200px;">&nbsp;&nbsp;
                            <asp:Label ID="lb_Security" runat="server" ForeColor="Black" CssClass="form" 
                                text="機密表單：" Visible="False" ></asp:Label></td>
                        <td style="width: 300px; height: 23px;">
                            <asp:DropDownList ID="ddl_SecurityItem" runat="server" Visible="False" />
                            <asp:Label ID="lbSecurity" runat="server" Text="查無此密等申請單" Visible="False" />
                            <asp:HyperLink ID="HLSecurity" runat="server" Visible="False" Target ="_blank">檢視</asp:HyperLink>
                            <asp:Button ID="btnApplyForSecurity" runat="server" Text="申請" Visible="False" />
                       </td>
                    </tr> 
                    <tr>
                        <td style="width: 120px" align="right">
                            <asp:Label ID="Label15" runat="server" ForeColor="Black" Text="申請時間：" CssClass="form" ></asp:Label></td>
                        <td class="style2" colspan="3"  >
                            <asp:Label ID="Lab_time" runat="server" ForeColor="Black" CssClass="form" width = "200" ></asp:Label>
                            <div style="display:none"><asp:Label  ID="hid_readonly" runat="server" Visible ="true" />
                            <asp:Label  ID="hid_status" runat="server" Visible ="true"/>
                            <asp:Label ID="hid_eformsn" runat="server" Visible ="true" /></div>
                        </td>
                    </tr> 
            </table>
            <table border="0" style="width: 740px; height: 57px; color:Red" id="tbContent" >
                <tr>
                    <td style="height: 23px; width: 120px;" align="right">*
                        <asp:Label ID="Label9" runat="server" ForeColor="Black" Text="主旨、簡由<br />或(資料名稱)：" CssClass="form" ></asp:Label></td>
                    <td style="height: 23px; ">
                        <asp:TextBox ID="tbFile_Name" runat="server" MaxLength = "100" Width="583px" /></td>
                </tr>
                <tr>
                    <td style="height: 23px; width: 120px;" align="right" valign="top">*
                        <asp:Label ID="Label8" runat="server" ForeColor="Black" Text="用途：" CssClass="form" ></asp:Label></td>
                    <td style="height: 23px; ">
                        <asp:TextBox ID="tbUse_For" runat="server" MaxLength = "255" Width="583px" Height="100px" TextMode="MultiLine" /></td>
                </tr>
                <tr>
                    <td style="height: 23px; width: 120px;" align="right">*
                        <asp:Label ID="Label14" runat="server" ForeColor="Black" Text="原件張數：" CssClass="form" ></asp:Label></td>
                    <td style="height: 23px; ">
                        <asp:TextBox ID="txtOriginalSheets" runat="server" MaxLength = "3" Width="50px" onchange="sumCalculate(this);" /></td>
                </tr> 
                <tr>
                    <td style="height: 23px; width: 120px;" align="right">*
                        <asp:Label ID="Label16" runat="server" ForeColor="Black" Text="每份複印張數：" CssClass="form" ></asp:Label></td>
                    <td style="height: 23px; ">
                        <asp:TextBox ID="txtCopySheets" runat="server" MaxLength = "3" Width="50px" onchange="sumCalculate(this);" /></td>
                </tr> 
                <tr>
                    <td style="height: 23px; width: 120px;" align="right">
                        <asp:Label ID="Label17" runat="server" ForeColor="Black" Text="合計複印張數：" CssClass="form" ></asp:Label></td>
                    <td style="height: 23px; ">
                        <asp:TextBox ID="txtTotalSheets" runat="server" MaxLength = "3" Width="50px" /></td>
                </tr>
            </table>
            <table border="1" cellpadding="0" style="width: 740px; border:1px solid #6499CC; color:Black; border-collapse:collapse;" id="tbSecurityContent" class="form">
            <tr>
                <td align="center" valign="middle" colspan="3" style="border:1px solid #6499CC; height:20px;">
                    <asp:Label ID="lbTop1unitName" runat="server"></asp:Label>複(影)印資料申請單</td>
            </tr>
            <tr>
                <td align="center" rowspan="5" style="width:3%; height:160px; border:1px solid #6499CC;">原<br /><br />件<br /><br />資<br /><br />料</td>
                <td align="center" style="width:22%; border:1px solid #6499CC; height:40px;">主&nbsp; 旨&nbsp; 、&nbsp; 簡&nbsp; 由<br />或 ( 資 料 名 稱 )</td>
                <td style="width:75%; border:1px solid #6499CC;" align="left">
                    &nbsp;<asp:Label ID="lbSubject" runat="server"></asp:Label></td>
            </tr>
            <tr>
                <td align="center" style="border:1px solid #6499CC; height:20px;">發 文 時 間 、 字 號</td>
                <td align="left" style="border:1px solid #6499CC;">
                    &nbsp;<asp:Label ID="lbSignDateTime" runat="server"></asp:Label>
                    、<asp:Label ID="lbSecurity_No" runat="server"></asp:Label></td>
            </tr>
            <tr>
                <td align="center" colspan="2" style="border:1px solid #6499CC; height:40px;">
                    <table border="0" style="width:100%; border-collapse:collapse; color:Black; height:40px;" class="form">
                    <tr>
                        <td align="center" style="width:7%; border-right:1px solid #6499CC;">機 密<br />等 級</td>
                        <td style="width:14%; border-right:1px solid #6499CC;" align="left">
                            &nbsp;<asp:Label ID="lbSecurity_Level" runat="server"></asp:Label></td>
                        <td align="center" style="width:7%; border-right:1px solid #6499CC;">機 密<br />屬 性</td>
                        <td style="width:29%; border-right:1px solid #6499CC;" align="left">
                            &nbsp;<asp:Label ID="lbSecurity_Type" runat="server"></asp:Label></td>
                        <td align="center" style="width:14%; border-right:1px solid #6499CC;">保 密 期 限/<br />解 密 條 件</td>
                        <td style="width:29%;" align="left">
                            &nbsp;<asp:Label ID="lbSecurity_Range" runat="server"></asp:Label></td>
                    </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center" style="border:1px solid #6499CC; height:20px;">產&nbsp; &nbsp;  製&nbsp; &nbsp;  單&nbsp; &nbsp;  位</td>
                <td align="left" style="border:1px solid #6499CC;">
                    <asp:Label ID="lbProduceUnit" runat="server"></asp:Label>
                </td>
            </tr>
            <tr>
                <td align="center" style="border:1px solid #6499CC; height:40px;">同 意 複 ( 影 ) 印<br />時&nbsp;  間&nbsp;  /&nbsp;  文&nbsp;  號</td>
                <td style="border:1px solid #6499CC;">
                    <table border="0" style="width:100%; border-collapse:collapse; color:Black;" class="form">
                    <tr>
                        <td align="left" style="width:40%; border-right:1px solid #6499CC;">
                            <asp:Label ID="lbAgreeTimeOrNumber" runat="server"></asp:Label></td>
                        <td align="center" style="width:30%; border-right:1px solid #6499CC;">同&nbsp; 意&nbsp; 複&nbsp; (&nbsp;影 )&nbsp; 印<br />權 責 長 官 級 職 姓 名</td>
                        <td align="left" style="width:30%;">
                            <asp:Label ID="lbAgreeSuperior" runat="server"></asp:Label></td>
                    </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center" rowspan="4" style="border:1px solid #6499CC; height:160px;">複<br /><br />印<br /><br />事<br /><br />項</td>
                <td align="center" style="border:1px solid #6499CC; height:40px;">用&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     途</td>
                <td align="left" style="border:1px solid #6499CC;">
                    <asp:Label runat="server" ID="lbPurpose"></asp:Label></td>
            </tr>
            <tr>
                <td align="center" style="border:1px solid #6499CC; height:40px;">複&nbsp;&nbsp;&nbsp;&nbsp; 印&nbsp;&nbsp;&nbsp;&nbsp; 時&nbsp;&nbsp;&nbsp;&nbsp; 間</td>
                <td style="border:1px solid #6499CC;">
                    <table border="0" style="width:100%; height:40px; border-collapse:collapse; color:Black;" class="form">
                    <tr>
                        <td align="left" style="width:33%; border-right:1px solid #6499CC;">
                            &nbsp;<asp:Label ID="lbPrinter_Datetime" runat="server"></asp:Label></td>
                        <td align="center" style="width:34%; border-right:1px solid #6499CC;">浮&nbsp; 水&nbsp; 印&nbsp; 暗&nbsp; 記&nbsp; 編&nbsp; 號</td>
                        <td align="left" style="width:33%">
                            &nbsp;<asp:Label ID="lbPrinter_Num" runat="server"></asp:Label></td>
                    </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="2" style="border:1px solid #6499CC; height:40px;">
                    <table border="0" style="width:100%; border-collapse:collapse; color:Black;" class="form">
                    <tr>
                        <td align="center" style="width:12%; height:40px; border-right:1px solid #6499CC;">原 件 張 數</td>
                        <td align="left" style="width:12%; border-right:1px solid #6499CC;">
                            &nbsp;<asp:Label ID="lbOri_sheet" runat="server"></asp:Label></td>
                        <td align="center" style="width:12%; border-right:1px solid #6499CC;">每&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     張<br />複 印 張 數</td>
                        <td align="left" style="width:12%; border-right:1px solid #6499CC;">
                            &nbsp;<asp:Label ID="lbCopy_sheet" runat="server"></asp:Label></td>
                        <td align="center" style="width:12%; border-right:1px solid #6499CC;">合&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     計<br />複 印 張 數</td>
                        <td align="left" style="width:12%; border-right:1px solid #6499CC;">
                            &nbsp;<asp:Label ID="lbTotal_sheet" runat="server"></asp:Label></td>
                        <td align="center" style="width:16%; border-right:1px solid #6499CC;">複&nbsp; ( 影 )&nbsp; 印<br />張 數 流 水 號</td>
                        <td align="left" style="width:12%">
                            &nbsp;<asp:Label ID="lbSheet_ID" runat="server"></asp:Label></td>
                    </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center" style="border:1px solid #6499CC; height:40px;">附&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 註</td>
                <td align="left" style="border:1px solid #6499CC;">
                    &nbsp;<asp:Label ID="lbMemo" runat="server"></asp:Label></td>
            </tr>
        </table>
        <table style="width: 720px; height: 37px; color:Red" >
            <tr>
                <td align="center" style="height: 29px; width: 720px;">
                    <asp:Button ID="Button1" runat="server" Text="CheckCard" Visible = "false" />&nbsp;
                    <asp:Button ID="btWrite" runat="server" Text="憑據寫入" Visible = "false"   
                                OnClientClick="return smfverifyWriteTicket();" style="height: 21px"/>&nbsp;
                    <asp:Button ID="but_exe" runat="server" Text="送件"  OnClientClick="return checksmf();"/>&nbsp;
                    <asp:Button ID="backBtn" runat="server" Text="駁回" />&nbsp;
                    <asp:Button ID="tranBtn" runat="server" Text="呈轉" />&nbsp;
                    <asp:Button ID="bt_clear" runat="server" Text="清除已申請" Visible = "false"  OnClientClick="return clearsmf();"/>&nbsp;
                </td>
            </tr> 
        </table>           
        <uc1:FlowRoute ID="FlowRoute1" runat="server" />
    </td>
</tr>
</table>
</div>
</form> 
</body>
</html>
