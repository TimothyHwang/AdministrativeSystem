<%@ Page Language="VB" AutoEventWireup="false" CodeFile="calendar.aspx.vb" Inherits="calendar" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<meta http-equiv="Expires" content="0">
<meta http-equiv="Cache-Control" content="no-cache">
<meta http-equiv="Pragma" content="no-cache">
<base target="_self">
<script language="JavaScript">
function LoadCalendar(dDate){
  var p=window.dialogArguments;  
  <% if return_value="1" then %>
        window.returnValue=dDate;
  <% else %>
      p.document.getElementById("<%=sTextBoxID %>").value=dDate;
  <% end if %>
  window.close();
}
function  setSize(){
  //dialogWidth=10;
  //dialogHeight=10.5;
  calendar.focus();  //讓程式不要focus在下拉選單上
}
function  CalendarChage(){
    calendar.submit();
}
</SCRIPT>
<TITLE>日期</TITLE>
</HEAD>
<BODY onload="setSize();"  style="background-color: #ffffcc;">
<form name ='calendar' action="calendar.aspx" method="get">
<asp:PlaceHolder ID="PlaceHolder1" runat="server" ></asp:PlaceHolder>
<input type="hidden" name="TextBoxId" value="<%=sTextBoxID %>" />
</form>
</BODY>
</HTML>

