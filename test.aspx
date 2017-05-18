<%@ Page Language="VB" AutoEventWireup="false" CodeFile="test.aspx.vb" Inherits="test" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="../../css/jquery.timepick.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.10.2.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.min.js" type="text/javascript"></script>
    <script src="../../script/jquery-ui-timepicker-addon.js" type="text/javascript"></script>
    <script src="../../script/jquery-ui-sliderAccess.JS" type="text/javascript"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <input type="text" name="Txt_nFIXTIME" id="Txt_nFIXTIME" value="" />
        <asp:Button ID="Button1" runat="server" Text="Button" />
    
    </div>
    </form>
    <script language="javascript" type="text/javascript">
        $(function () {
            $("#Txt_nFIXTIME").timepicker();
        }); 
    </script>
</body>
</html>
