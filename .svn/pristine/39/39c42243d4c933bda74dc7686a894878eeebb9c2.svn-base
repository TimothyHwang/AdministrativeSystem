<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA09001.aspx.vb" Inherits="M_Source_09_MOA09001" %>
<%@ Register Src="../90/FlowRoute.ascx" TagName="FlowRoute" TagPrefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>門禁會議管制申請單</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.7.2.min.js" type="text/javascript"></script>
    <script type="text/javascript" language="javascript">
        document.oncontextmenu = new Function("return false");
        $(document).ready(function () {
            var flagReadOnly = '<%=ViewState("read_only") %>';
            if (flagReadOnly == "") { //若readyonly為空值則為新增表單
                CreateGateList(); //建立下拉選單
                CreateSelectedGate() //建立已選營門顯示資料
            }
            else { //瀏覽或審核表單
                ShowStaticGate() //建立無事件之營門顯示資料
            }
        });

        function CreateGateList() { //建立下拉選單,若已在已選擇之營門資料隱藏欄位內之營門編號,則不加入選單
            var ddlGate = $("#ddlEnteringGate");
            var gateValueArray = $("#" + "<%=hidEnteringGate.ClientID %>").val().split(",");
            ddlGate.empty();
            //if (!funContain(gateValueArray, "1")) ddlGate.append("<option value='1'>博愛營區一號門</option>");
            //if (!funContain(gateValueArray, "2")) ddlGate.append("<option value='2'>博愛營區二號門</option>");
            //if (!funContain(gateValueArray, "3")) ddlGate.append("<option value='3'>博愛營區三號門</option>");
            //if (!funContain(gateValueArray, "4")) ddlGate.append("<option value='4'>博愛營區四號門</option>");
            //if (!funContain(gateValueArray, "5")) ddlGate.append("<option value='5'>採購中心</option>");
            //if (!funContain(gateValueArray, "6")) ddlGate.append("<option value='6'>博一大樓</option>");
            //if (!funContain(gateValueArray, "7")) ddlGate.append("<option value='7'>博二大樓</option>");
            //if (!funContain(gateValueArray, "8")) ddlGate.append("<option value='8'>博一會客室</option>");
            //if (!funContain(gateValueArray, "9")) ddlGate.append("<option value='9'>國防採購室會客室</option>");
            //if (!funContain(gateValueArray, "10")) ddlGate.append("<option value='10'>二號門(軍車進出)</option>");
			if (!funContain(gateValueArray, "11")) ddlGate.append("<option value='11'>武德樓會客室</option>");
			if (!funContain(gateValueArray, "12")) ddlGate.append("<option value='12'>一號營門</option>");
			if (!funContain(gateValueArray, "12")) ddlGate.append("<option value='13'>二號營門</option>");
			if (!funContain(gateValueArray, "13")) ddlGate.append("<option value='14'>三號營門</option>");
			if (!funContain(gateValueArray, "14")) ddlGate.append("<option value='15'>四號營門</option>");
			if (!funContain(gateValueArray, "12")) ddlGate.append("<option value='16'>五號營門</option>");
        }

        function funContain(gateArray,targetNo) { //Contains Function
            for (var i = 0; i < gateArray.length; i++) {
                if (gateArray[i] == targetNo) return true;
            }
            return false;
        }

        function CreateSelectedGate() { //建立已選營門顯示資料
            var gateValueArray = $("#" + "<%=hidEnteringGate.ClientID %>").val().split(",");
            $("#tdGate").html("");
            if (funContain(gateValueArray, "1"))
                $("#tdGate").append("博愛營區一號門<a href='#' onclick='DelGate(1);'>X</a><br />");
            if (funContain(gateValueArray, "2"))
                $("#tdGate").append("博愛營區二號門<a href='#' onclick='DelGate(2);'>X</a><br />");
            if (funContain(gateValueArray, "3"))
                $("#tdGate").append("博愛營區三號門<a href='#' onclick='DelGate(3);'>X</a><br />");
            if (funContain(gateValueArray, "4"))
                $("#tdGate").append("博愛營區四號門<a href='#' onclick='DelGate(4);'>X</a><br />");
            if (funContain(gateValueArray, "5"))
                $("#tdGate").append("採購中心<a href='#' onclick='DelGate(5);'>X</a><br />");
            if (funContain(gateValueArray, "6"))
                $("#tdGate").append("博一大樓<a href='#' onclick='DelGate(6);'>X</a><br />");
            if (funContain(gateValueArray, "7"))
                $("#tdGate").append("博二大樓<a href='#' onclick='DelGate(7);'>X</a><br />");
            if (funContain(gateValueArray, "8"))
                $("#tdGate").append("博一會客室<a href='#' onclick='DelGate(8);'>X</a><br />");
            if (funContain(gateValueArray, "9"))
                $("#tdGate").append("國防採購室會客室<a href='#' onclick='DelGate(9);'>X</a><br />");
            if (funContain(gateValueArray, "10"))
                $("#tdGate").append("二號門(軍車進出)<a href='#' onclick='DelGate(10);'>X</a><br />");
			if (funContain(gateValueArray, "11"))
                $("#tdGate").append("武德樓會客室<a href='#' onclick='DelGate(11);'>X</a><br />");
			if (funContain(gateValueArray, "12"))
                $("#tdGate").append("一號營門<a href='#' onclick='DelGate(12);'>X</a><br />");
			if (funContain(gateValueArray, "13"))
                $("#tdGate").append("二號營門<a href='#' onclick='DelGate(13);'>X</a><br />");
			if (funContain(gateValueArray, "14"))
                $("#tdGate").append("三號營門<a href='#' onclick='DelGate(14);'>X</a><br />");
			if (funContain(gateValueArray, "15"))
                $("#tdGate").append("四號營門<a href='#' onclick='DelGate(15);'>X</a><br />");
			if (funContain(gateValueArray, "16"))
                $("#tdGate").append("五號營門<a href='#' onclick='DelGate(16);'>X</a><br />");
				
        }

        function ShowStaticGate() { //無事件之營門顯示資料
            var gateValueArray = $("#" + "<%=hidEnteringGate.ClientID %>").val().split(",");
            if (funContain(gateValueArray, "1"))
                $("#tdGate").append("博愛營區一號門<br />");
            if (funContain(gateValueArray, "2"))
                $("#tdGate").append("博愛營區二號門<br />");
            if (funContain(gateValueArray, "3"))
                $("#tdGate").append("博愛營區三號門<br />");
            if (funContain(gateValueArray, "4"))
                $("#tdGate").append("博愛營區四號門<br />");
            if (funContain(gateValueArray, "5"))
                $("#tdGate").append("採購中心<br />");
            if (funContain(gateValueArray, "6"))
                $("#tdGate").append("博一大樓<br />");
            if (funContain(gateValueArray, "7"))
                $("#tdGate").append("博二大樓<br />");
            if (funContain(gateValueArray, "8"))
                $("#tdGate").append("博一會客室<br />");
            if (funContain(gateValueArray, "9"))
                $("#tdGate").append("國防採購室會客室<br />");
            if (funContain(gateValueArray, "10"))
                $("#tdGate").append("二號門(軍車進出)<br />");
            if (funContain(gateValueArray, "11"))
                $("#tdGate").append("武德樓會客室<br />");
            if (funContain(gateValueArray, "12"))
                $("#tdGate").append("一號營門<<br />");
            if (funContain(gateValueArray, "13"))
                $("#tdGate").append("二號營門<br />");
            if (funContain(gateValueArray, "14"))
                $("#tdGate").append("三號營門<br />");
			if (funContain(gateValueArray, "15"))
                $("#tdGate").append("四號營門<br />");
			if (funContain(gateValueArray, "16"))
                $("#tdGate").append("五號營門<br />");
        }

        function checkData() {
            var objSubject = $("#" + "<%=txtSubject.ClientID %>");
            if (objSubject.val() == "") {
                alert("開會事由欄位不可空白");
                objSubject.focus();
                return false;
            }
            var objLocation = $("#" + "<%=txtLocation.ClientID %>");
            if (objLocation.val() == "") {
                alert("開會地點欄位不可空白");
                objLocation.focus();
                return false;
            }
            var objModerator = $("#" + "<%=txtModerator.ClientID %>");
            if (objModerator.val() == "") {
                alert("主持人欄位不可空白");
                objModerator.focus();
                return false;
            }
            var objPhoneNumber = $("#" + "<%=txtPhoneNumber.ClientID %>");
            if (objPhoneNumber.val() == "") {
                alert("聯絡人電話欄位不可空白");
                objPhoneNumber.focus();
                return false;
            }
            var objMeetingDate = $("#" + "<%=txtMeetingDate.ClientID %>");
            if (objMeetingDate.val() == "") {
                alert("開會日期(時間)欄位不可空白");
                objMeetingDate.focus();
                return false;
            }
            var objDocumentWord = $("#" + "<%=txtDocumentWord.ClientID %>");
            if (objDocumentWord.val() == "") {
                alert("發文字號欄位不可空白");
                objDocumentWord.focus();
                return false;
            }
            var objDocumentNo = $("#" + "<%=txtDocumentNo.ClientID %>");
            if (objDocumentNo.val() == "") {
                alert("發文字號欄位不可空白");
                objDocumentNo.focus();
                return false;
            }
            var reNum = /^\d+$/;
            if (!reNum.test(objDocumentNo.val())) {
                alert("發文字號欄位只能輸入數字");
                objDocumentNo.focus();
                return false;
            }
            var objEnteringPeopleNumber = $("#" + "<%=txtEnteringPeopleNumber.ClientID %>");
            if (objEnteringPeopleNumber.val() == "") {
                alert("進出人員(部外)欄位不可空白");
                objEnteringPeopleNumber.focus();
                return false;
            }
            if (!reNum.test(objEnteringPeopleNumber.val())) {
                alert("進出人員(部外)欄位只能輸入數字");
                objEnteringPeopleNumber.focus();
                return false;
            }
            objEnteringPeopleNumber.val(Number(objEnteringPeopleNumber.val()));
            var objEnteringGate = $("#" + "<%=hidEnteringGate.ClientID %>");
            if (objEnteringGate.val() == "") {
                alert("請選擇進出營門");
                $("#ddlEnteringGate").focus();
                return false;
            }
        }

        function AddFile() {
            $('#tdFile').append("<div><input id='tempFileId' name='tempFileName' type='file' /><input type='button' id='btnDelFile' value='－' style='height:20px;' onclick='DelFile(this);' /></div>");
        }

        function DelFile(obj) {
            $(obj).parent().remove();
        }

        function AddGate() { //增加選擇進出營門資料
            var gateNo = $("#ddlEnteringGate").val(); //所選之營門資料編號
            //已選擇之進出營門資料,資料字串紀錄格式: 營門編號1,營門編號2,...
            var odjGateValue = $("#" + "<%=hidEnteringGate.ClientID %>");
            if (Number(gateNo) != 0 && !isNaN(Number(gateNo))) {
                if (odjGateValue.val() == "" || odjGateValue.val() == "null") {
                    odjGateValue.val(gateNo);
                }
                else {
                    odjGateValue.val(odjGateValue.val() + "," + gateNo);
                }
                CreateGateList(); //重新建立下拉選單,已選擇的則移出選單
                CreateSelectedGate() //建立已選營門顯示資料
            }
        }

        function DelGate(gateNo){ //移除選擇進出營門資料
            //已選擇之進出營門資料,資料字串紀錄格式: 營門編號1,營門編號2,...
            var odjGateValue = $("#" + "<%=hidEnteringGate.ClientID %>");
            var gateValueArray = odjGateValue.val().split(",");
            odjGateValue.val("");
            for (var i = 0; i < gateValueArray.length; i++) {
                if (gateNo != gateValueArray[i]) {
                    if (odjGateValue.val() == "" || odjGateValue.val() == "null") {
                        odjGateValue.val(gateValueArray[i]);
                    }
                    else {
                        odjGateValue.val(odjGateValue.val() + "," + gateValueArray[i]);
                    }
                }
            }
            CreateGateList(); //重新建立下拉選單,已選擇的則移出選單
            CreateSelectedGate() //建立已選營門顯示資料
        }
    </script>
    <style type="text/css">
        .style1
        {
            width: 80px;
            height: 26px;
        }
        .style2
        {
            width: 150px;
            height: 26px;
        }
        .style3
        {
            width: 60px;
            height: 26px;
        }
        .style4
        {
            width: 160px;
            height: 26px;
        }
        .style5
        {
            width: 130px;
            height: 26px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <table width="750" border="1" cellspacing="0" cellpadding="5" bgcolor="#ffffff" bordercolor="#6699cc" bordercolorlight="#74a3d6" bordercolordark="#000000" style="LEFT: 20px; TOP: 10px">
        <tr>
            <td valign="bottom" bgcolor="#6699cc" bordercolorlight="#66aaaa" bordercolordark="#ffffff"><font color=white><b>&nbsp;門禁會議管制申請單</b></font>
            </td>
        </tr>
        
        <tr>
            <td>            
             <fieldset id="tableB" style="width: 740px">           
		        <table cellspacing="0" cellpadding="0" width="750" align="center" border="0" style="height:30">
			        <tr>
				        <td nowrap="nowrap" class="style1"><asp:Label ID="Label1" runat="server" Text="填表人單位：" CssClass="form" Width="80px" ForeColor="Black"></asp:Label></td>
				        <td nowrap="nowrap" class="style2"><asp:label id="lbFillOutUnit" runat="server" 
                                CssClass="form" Width="139px" ForeColor="Black"></asp:label></td>
				        <td nowrap="nowrap" class="style3"><asp:Label ID="Label2" runat="server" Text="姓名：" CssClass="form" Width="50px" ForeColor="Black"></asp:Label></td>
				        <td nowrap="nowrap" class="style4"><asp:label id="lbFillOutName" runat="server" 
                                CssClass="form" Width="140px" ForeColor="Black"></asp:label></td>
				        <td nowrap="nowrap" class="style3"><asp:Label ID="Label3" runat="server" Text="級職：" Width="50px" CssClass="form" ForeColor="Black"></asp:Label></td>
				        <td nowrap="nowrap" class="style5"><asp:label id="lbFillOutTitle" runat="server" 
                                CssClass="form" ForeColor="Black" Width="130px"></asp:label></td>
			        </tr>
                    <tr>
				        <td nowrap="nowrap" class="style1"><asp:Label ID="Label7" runat="server" Text="申請人單位：" CssClass="form" Width="80px" ForeColor="Black"></asp:Label></td>
				        <td nowrap="nowrap" class="style2"><asp:label id="lbCreatorUnit" runat="server" 
                                CssClass="form" Width="139px" ForeColor="Black"></asp:label></td>
				        <td nowrap="nowrap" class="style3"><asp:Label ID="Label17" runat="server" Text="姓名：" CssClass="form" Width="50px" ForeColor="Black"></asp:Label></td>
				        <td nowrap="nowrap" class="style4">
                            <asp:DropDownList ID="ddlCreatorName" runat="server" AutoPostBack="True">
                            </asp:DropDownList>
                            <asp:Label ID="lbCreatorName" runat="server" CssClass="form" ForeColor="Black" Width="130px"></asp:Label>
                        </td>
				        <td nowrap="nowrap" class="style3"><asp:Label ID="Label19" runat="server" Text="級職：" Width="50px" CssClass="form" ForeColor="Black"></asp:Label></td>
				        <td nowrap="nowrap" class="style5"><asp:label id="lbCreatorTitle" runat="server" 
                                CssClass="form" ForeColor="Black" Width="130px"></asp:label></td>
			        </tr>
		        </table>    			    
		    </fieldset>    		    		
		    <table cellspacing="0" cellpadding="0" width="750" align="center" border="0">   		
			    <tr>
				    <td nowrap="nowrap" class="form" style="height: 33px; width: 10%;"><font color="#ff0033">* 
                        <asp:Label ID="Label12" runat="server" CssClass="form" ForeColor="Black" 
                            Text="開會事由："></asp:Label></font>
                    </td>
				    <td nowrap="nowrap" class="form" style="height: 33px; width: 40%;"><asp:textbox id="txtSubject" runat="server" Width="90%"></asp:textbox></td>
				    <td nowrap="nowrap" class="form" style="height: 33px; width: 10%;">&nbsp;&nbsp;
                        <asp:Label ID="Label13" runat="server" CssClass="form" ForeColor="Black" Text="申請時間："></asp:Label></td>
				    <td nowrap="nowrap" class="form" style="height: 33px; width: 40%;">
                        <asp:label id="lbCreateDate" runat="server" ForeColor="Black" Width="150px"></asp:label></td>
			    </tr>	    
			    <tr>
				    <td nowrap="nowrap" class="form" style="height: 33px; width: 10%;"><font color="#ff0033">* <asp:Label ID="Label11" runat="server" CssClass="form" ForeColor="Black"
                            Text="開會地點："></asp:Label></font></td>
				    <td nowrap="nowrap" class="form" style="height: 33px; width: 40%;"><asp:textbox id="txtLocation" runat="server" Width="90%"></asp:textbox></td>
				    <td nowrap="nowrap" class="form" style="height: 33px; width: 10%;">&nbsp;&nbsp;
                        <asp:Label ID="Label14" runat="server" CssClass="form" ForeColor="Black" Text="承辦單位："></asp:Label></td>
				    <td nowrap="nowrap" class="form" style="height: 33px; width: 40%;">
                         <asp:label id="lbSponsor" runat="server" ForeColor="Black" Width="150px"></asp:label></td>
			    </tr>		 			
			    <tr>
				    <td nowrap="nowrap" class="form" style="height: 33px; width: 10%;"><font color="#ff0033">* <asp:Label ID="Label10" runat="server" CssClass="form" ForeColor="Black"
                            Text="主持人："></asp:Label></font></td>
				    <td nowrap="nowrap" class="form" style="height: 33px; width: 40%;"><asp:textbox id="txtModerator" runat="server" Width="90%"></asp:textbox></td>
				    <td nowrap="nowrap" class="form" style="height: 33px; width: 10%;"><font color="#ff0033">* <asp:Label ID="Label15" runat="server" CssClass="form" ForeColor="Black"
                            Text="聯絡人電話："></asp:Label></font></td>
				    <td nowrap="nowrap" class="form" style="height: 33px; width: 40%;">
                        <asp:textbox id="txtPhoneNumber" runat="server" Width="90%"></asp:textbox></td>
			    </tr>	
			    <tr>
				    <td nowrap="nowrap" class="form" style="height: 33px; width: 10%;"><font color="#ff0033">* </font><asp:Label ID="Label9" runat="server" CssClass="form"
                            ForeColor="Black" Text="開會日期(時間)："></asp:Label></td>
				    <td nowrap="nowrap" class="form" style="height: 33px; width: 40%;">
					    <asp:textbox id="txtMeetingDate" runat="server" OnKeyDown="return false" 
                            Width="100px" ReadOnly="True"></asp:textbox>
                        <asp:ImageButton ID="ImgDate1" runat="server" ImageUrl="~/Image/calendar.gif" />&nbsp;
                        <asp:DropDownList ID="ddlHour" runat="server" AutoPostBack="True">
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
                    &nbsp;時&nbsp;<asp:DropDownList ID="ddlMinute" runat="server">
                            <asp:ListItem>0</asp:ListItem>
                            <asp:ListItem>5</asp:ListItem>
                            <asp:ListItem>10</asp:ListItem>
                            <asp:ListItem>15</asp:ListItem>
                            <asp:ListItem>20</asp:ListItem>
                            <asp:ListItem>25</asp:ListItem>
                            <asp:ListItem>30</asp:ListItem>
                            <asp:ListItem>35</asp:ListItem>
                            <asp:ListItem>40</asp:ListItem>
                            <asp:ListItem>45</asp:ListItem>
                            <asp:ListItem>50</asp:ListItem>
                            <asp:ListItem>55</asp:ListItem>
                        </asp:DropDownList>
                    &nbsp;分
                    </td>                       
                    <td nowrap="nowrap" class="form" style="height: 33px; width: 10%;">
                        <font color="#ff0033">* </font><asp:Label ID="Label16" runat="server" CssClass="form" ForeColor="Black" Text="發文字號："></asp:Label>
                    </td>
                    <td nowrap="nowrap" class="form" style="height: 33px; width: 40%;">
                    <% If String.IsNullOrEmpty(ViewState("read_only").ToString()) Then%>
                        <asp:textbox id="txtDocumentWord" runat="server" Width="30%" MaxLength="4"></asp:textbox>字第
				        <asp:textbox id="txtDocumentNo" runat="server" Width="40%" MaxLength="10"></asp:textbox>號
                    <% Else%>
                        <asp:Label ID="lbDocumentNo" runat="server" Text="Label"></asp:Label>
                    <% End If%>
                    </td>
			    </tr>
                <tr>
				    <td nowrap="nowrap" class="form" style="height: 33px; width: 10%;"><font color="#ff0033">* <asp:Label ID="Label4" runat="server" CssClass="form" ForeColor="Black"
                            Text="進出人員(部外)："></asp:Label></font></td>
				    <td nowrap="nowrap" class="form" style="height: 33px; width: 40%;"><asp:textbox id="txtEnteringPeopleNumber" runat="server" Width="10%"></asp:textbox>&nbsp;員</td>
				    <td nowrap="nowrap" class="form" style="height: 33px; width: 10%;">&nbsp;&nbsp;
                        <% If String.IsNullOrEmpty(ViewState("read_only").ToString()) Then%>
                        <asp:Label ID="Label5" runat="server" CssClass="form" ForeColor="Black" Text="營門選單："></asp:Label></td>
                        <% End If%>
				    <td nowrap="nowrap" class="form" style="height: 33px; width: 40%;">
                        <% If String.IsNullOrEmpty(ViewState("read_only").ToString()) Then%>
                        <select id="ddlEnteringGate" style="width:120px"></select>
                        <input type="button" id="btnAddGate" value="＋" style="height:20px" onclick="AddGate();" /></td>
                        <% End If%>
                        <asp:HiddenField ID="hidEnteringGate" runat="server" Value="" />
			    </tr>
                <tr>
                    <td nowrap="nowrap" class="form" style="width: 10%;" valign="top"><font color="#ff0033">* 
                        <asp:Label ID="Label6" runat="server" CssClass="form" ForeColor="Black"
                            Text="附件："></asp:Label></font>
                    </td>
                    <td class="form" id="tdFile" valign="top">
                        <% If String.IsNullOrEmpty(ViewState("read_only").ToString()) Then%>
                        <input type="file" runat="server" /><input type="button" id="btnAddFile" value="＋" style="height:20px" onclick="AddFile();" />
                        <% Else%>
                        <div id="divFiles" runat="server"></div>
                        <% End If %>
                    </td>
                    <td class="form" valign="top"><font color="#ff0033">* <asp:Label ID="Label8" runat="server" CssClass="form" ForeColor="Black"
                            Text="進出營門："></asp:Label></font></td>
                    <td class="form" id="tdGate" valign="top">
                    </td>
			    </tr>
                <tr>
                    <td></td>
                    <td class="form"><font color="#ff0033">附件須為已奉核之開會通知單</font></td>
                    <td></td>
                    <td></td>
                </tr>
                <% If ViewState("read_only").ToString() = "2" Then%>
			    <tr>
				    <td nowrap="nowrap" class="form" style="height: 27px; width: 10%;" >&nbsp;&nbsp;
                        <asp:Label ID="Label20" runat="server" CssClass="form" ForeColor="Black" Text="批核意見："></asp:Label></td>
				    <td nowrap="nowrap" width="90%" class="form" colspan="3" style="height: 27px" valign="middle">
                       <asp:TextBox ID="txtcomment" runat="server" Height="59px" MaxLength="255" Rows="3"
                                TextMode="MultiLine" Width="550px"></asp:TextBox>
                        <asp:Button ID="But_PHRASE" runat="server" Text="批核片語" /></td>
			    </tr>
                <% End If%>
			    <tr>
				    <td class="form" colspan="4" style="height: 24px">
					    <div align="center"><asp:button id="btnSubmit" runat="server" Text="送件" 
                                onclientclick="return checkData();"></asp:button>
                            <asp:button id="btnApprove" runat="server" Text="核准"></asp:button>
                            <asp:Button ID="btnReject" runat="server" Text="駁回" />
                        </div>
				    </td>
			    </tr>
		    </table>		
            <uc1:flowroute ID="FlowRoute1" runat="server" />            
            </td>
        </tr>
        </table>
        <div id="Div_grid" runat="server" style="border-right: lightslategray 2px solid;
            border-top: lightslategray 2px solid; display: block; z-index: 3; left: 269px;
            border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
            position: absolute; top: 705px; height: 150pt; background-color: white" visible="false">
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
