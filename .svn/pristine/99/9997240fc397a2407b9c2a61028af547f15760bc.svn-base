<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Menu.aspx.vb" Inherits="OA_Menu" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>權限顯示</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>

<body class='menuBody'>
            
<asp:PlaceHolder ID="PlaceHolder1" runat="server"></asp:PlaceHolder>

<!-- tool bar  -->
    <script language="javascript">
    var sOpenItem='';
    document.onmousedown = doMouseDown;
    // variables for up and down icon
    var apk = 0;
    var ik,ih;
    var oInterval;
    var menuContentsPosTop;
    // variables for 隱藏功能表
    var iFrameCols=parseInt(parent.mainframecol.cols.substr(0,parent.mainframecol.cols.indexOf(",")));
    var iCurrentCols;
    function OnGoTo(obj,addr) {
       obj.className='menuLeftNavDown';
       parent.right.location.href=addr;
    }
    function OnGoToTop(obj,addr) {
       obj.className='menuLeftNavDown';
       top.location.href=addr;
    }
    function domenu(objid) {
	    if (objid.id == lastDisplay) 
	 	    return;
    			
//	    if (typeof(menuBottom)=='undefined') {
//	       alert('請稍待功能表載入！');
//	       return;
//	    }   
	    
	    document.all(lastDisplay).style.display='none';
	    document.all(objid).style.display='';
	    lastDisplay=objid;
    }
    function doMouseDown() {
    }
    </script>

<%--<table border="0">
<tr>
<td style="width: 3px">
<div id="menuContents" style="position:absolute">



    <div id=menuBottom style="position:absolute">
    </div>
</div><!-- id=menuContents -->
</td></tr></table>--%>
          
    <SCRIPT>
    var i=1;
    var j;
    var thisFolder;
    var thisFolderTitle;
    while(eval("window.myFolder"+i)) {
	    thisFolder=eval("myFolder"+i);
	    thisFolderTitle=eval("myFolder"+i+"Title");
	    thisFolderImage=eval("myFolder"+i+"Image");
	    document.write("<table class='menuBar' border='0' cellpadding='0' cellspacing='0'>" + "\r\n");
	    document.write("  <tr>" + "\r\n");
	    document.write("    <td valign='middle' align='left' class='menuLeftNavOff' onclick=\"domenu('div" + i + "')\"> " + "&nbsp;" + "<img src='../../image/"+ thisFolderImage +"' width = '24px' height = '24'>" + "&nbsp;" + thisFolderTitle + " </td>" + "\r\n");
	    document.write("  </tr>" + "\r\n");
	    document.write("</table>" + "\r\n");
	    if (i == defaultDiv)
		    document.write("<div id='div" + i + "'>" + "\r\n");		    
	    else
		    document.write("<div id='div" + i + "' style='display:none'>" + "\r\n");
	    document.write("<table class='menuBody' border='0' cellpadding='0' cellspacing='0'>" + "\r\n");

	    j=0;
	    while(thisFolder[j]) {
		    document.write("  <tr class='menuLeftNavOff' onmouseover='this.className=\"menuLeftNavUp\"' onmouseout='this.className=\"menuLeftNavOff\"'>" + "\r\n");
		    document.write("    <td width='10%'>&nbsp;</td>" + "\r\n");
		    document.write("    <td width='90%' nowrap align='left' onclick=\"" + thisFolder[j+1] + "\">" + "\r\n");
		    document.write("      " + thisFolder[j] + "</td>" + "\r\n");
		    document.write("  </tr>" + "\r\n");
		    j+=2;
	    }
	    document.write("  <tr>" + "\r\n");
	    document.write("    <td height='70%' colspan=2>&nbsp;</td>" + "\r\n");
	    document.write("  </tr>" + "\r\n");
	    document.write("</table>" + "\r\n");
	    document.write("</div>" + "\r\n");
	    i++;
    }	
    </SCRIPT>

            
</body>

</html>
