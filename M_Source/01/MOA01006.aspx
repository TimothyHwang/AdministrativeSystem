<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA01006.aspx.vb" Inherits="Source_01_MOA01006" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>人員休假狀況</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <link href="../../css/jquery.datepick.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.10.2.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick-zh-TW.js" type="text/javascript"></script>
    <script type="text/javascript" language="javascript">
        $(function () {
            $("#Sdate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#SDate").val(), showTrigger: '#calImg' });
            $("#Edate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#EDate").val(), showTrigger: '#calImg' });
        }); 
    </script>
</head>
<body>
    <form id="form1" runat="server" >   
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center">
                    <asp:Label ID="Label1" runat="server" CssClass="toptitle" Text="人員休假狀況" Width="100%"></asp:Label>
            </td></tr>
        </table>         
	    <table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
		    <tr>
			    <td noWrap style="height: 25px; width: 10%;">
                    <asp:Label ID="Lab2" runat="server" Text="日期：" CssClass="form" ></asp:Label></td>
			    <td noWrap style="height: 25px; width: 90%;">
                    <asp:TextBox ID="Sdate" runat="server" Width="100px" OnKeyDown="return false" ></asp:TextBox>
                    <div style="display: none;">
                        <img id="calImg" src="../../Image/calendar.gif" alt="選擇日期" />
                    </div>
                    <asp:Label ID="Lab3" runat="server" Text="~" ></asp:Label>
                    <asp:TextBox ID="Edate" runat="server" Width="100px" OnKeyDown="return false"></asp:TextBox>
                    <asp:ImageButton ID="ImgSearch" runat="server" ImageUrl="~/Image/search.gif" ToolTip="查詢" />
                    </td>
		    </tr>
	    </table>                            
        
        <asp:PlaceHolder ID="PlaceHolder1" runat="server"></asp:PlaceHolder>         
			
            
    </form>
</body>
</html>
