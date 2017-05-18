<%@ Page Language="VB" AutoEventWireup="false" CodeFile="frameslider.aspx.vb" Inherits="OA_frameslider" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>框架收放</title>
    
    <script language="JavaScript">
    Slider="left"; 
    function goMove(){
      if (Slider=="right") {
         Slider="left";
         parent.mainframecol.cols="205,12,*";
         document.all("btimg").src = "../../Image/left.gif";
      }
      else {
         Slider="right";
         parent.mainframecol.cols="0,12,*";
         document.all("btimg").src = "../../Image/right.gif";
      }
    }

    function movein() {
         Slider="right";
         parent.mainframecol.cols="0,12,*";
         document.all("btimg").src = "../../Image/right.gif";
    }

    </script>
</head>
	<body leftmargin="0" topmargin="0" >
		<table width="100%" height="600" border="0" cellpadding="0" cellspacing="0" background="../../Image/bar-bg.gif" >
			<tr>
				<td valign="middle">
					<img id="btimg" src="../../Image/left.gif" onclick="goMove()" WIDTH="12" HEIGHT="27" style="CURSOR:hand">
				</td>
			</tr>
		</table>
	</body>
</html>
