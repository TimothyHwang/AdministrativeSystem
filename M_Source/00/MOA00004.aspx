<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA00004.aspx.vb" Inherits="Source_00_MOA00004" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >

<head runat="server">
    <title>表單流程樹狀圖</title>
    
  <link rel="stylesheet" type="text/css" href="../../Styles.css" />
<style type="text/css">
.cl  {font:9pt; color:#FFFFFF; background:#000099; cursor:hand}
.nr  {font:9pt; color:#000000}
.c2  {font:9pt; color:#FFFFFF; letter-spacing:0.1em; background-color:#707070}
span {color:#000000; cursor:default; text-decoration:none;}
.grouptxt
{
    cursor:hand;
}

</style>
<script language="javascript" type="text/javascript">
function action(TreeID) {
	var pl = "Bplussign.gif";
	var lpl = "Blastplus.gif";
	var Clpl = "Clastplus.gif";
	var mn = "Bminussign.gif";
	var lmn = "Blastminus.gif";
	var Clmn = "Clastminus.gif";

	var plGif = "../../image/Bplussign.gif";
	var lplGif = "../../image/Blastplus.gif";
	var ClplGif = "../../image/Clastplus.gif";
	var mnGif = "../../image/Bminussign.gif";
	var lmnGif = "../../image/Blastminus.gif";
	var ClmnGif = "../../image/Clastminus.gif";
			
	var sopen = "class1open.gif";
	var sclose = "class1close.gif";

	var sopenGif = "../../image/class1open.gif";
	var scloseGif = "../../image/class1close.gif";

	var strNode = TreeID + "cont";
	var strFldImg = TreeID + "img";
	var strFldSign = TreeID + "sign";

	//Sub tree
	if (TreeID=="") return;
	if (!document.all(strNode)) return;
	if (document.all(strNode).style.display=="")
		document.all(strNode).style.display="none";
	else
		document.all(strNode).style.display="";
	
	//Sign
	if (!document.all(strFldSign)) return;
	if (document.all(strFldSign).src.indexOf(pl) > -1) 
		document.all(strFldSign).src = mnGif;
	else if (document.all(strFldSign).src.indexOf(lpl) > -1) 
		document.all(strFldSign).src = lmnGif;
	else if (document.all(strFldSign).src.indexOf(mn) > -1)
		document.all(strFldSign).src = plGif;
	else if (document.all(strFldSign).src.indexOf(Clpl) > -1) 
		document.all(strFldSign).src = ClmnGif;
	else if (document.all(strFldSign).src.indexOf(Clmn) > -1)
		document.all(strFldSign).src = ClplGif;
	else
		document.all(strFldSign).src = lplGif;
}



</script>

</head>
<body onload="ini()" style="margin-top:0px;margin-left:2px; background-position: left top;background-color:#ffffff;" ondragstart="fOnDragStart()" ondragenter="fOnDragOver()" ondragover="fOnDragOver()" ondrop="doMouseUp()" ondragend="fOnDragEnd()">
    <form id="form1" method="post" action="MOA00006.aspx">
    <div>
    &nbsp;        
        <input type="hidden" name="STEPSID">
        <input type="hidden" name="group_id" value="">
        <!-- group id for drag to flow chart ; R1=上一級主管; A1=申請者的上一級主管   uid-3 : group id; uid-0:department id; uid-1:employee;uid-2:group leader-->
        <input type="hidden" name="group_name"><!-- group name for drag to flow chart-->
        <input type="hidden" name="MODE">
        <input type="hidden" name="IX">
        <input type="hidden" name="IY">
        <input type="hidden" name="organization_id" value="<%=organization_id%>">    
        <input type="hidden" name="eformid" value="<%=eformid%>">    
        <input type="hidden" name="eformrole" value="<%=eformrole%>">     
        <input type="hidden" id="frm_chinese_name" name="frm_chinese_name" value="<%=frm_chinese_name%>">    
        <input type="hidden" name="s">
        </div>
    </form>
    
    <div ID="fake" STYLE="border: solid 1px black;font-size:9pt; position:absolute;Z-INDEX:100; display:none; background-color:#ffcccc; text-align:center; color:#000000; width:80;">
</div>

<span style="top:3000px;left:2000;position:absolute;">&nbsp;
</span>
<!-- start left pan -->
<!--<div id="div2" class="GridDiv" style='width:100%;height:150;'>
   <table border="1" cellspacing="0" cellpadding="0">
   <tr> 
    <td> 開始</td>
    </tr>
   </table>
</div>-->


<div id="folder" style="top:0px;left:0;padding:0;position:absolute;width:100%; height:768; background-color: #EEF2FB  ;   font-size:small;background-image: url('../../image/gra.gif'); background-repeat: no-repeat;background-attachment: fixed;clip:rect(10 100% 100% 10);">
   
   
<!-- 組織圖  G-group, 4-group id -->
<div id="inFolder" style="position:absolute;left:3;top:60px;width:100%;clip:rect(0 100% 100% 0);">
	<br>	  
<!-- 列出系統內建使用者 -->
<table border="1" cellspacing="0" cellpadding="0">
<tr><td><img src="../../image/systemusr.gif" border="0" id="systemuserimg" align="absmiddle"  WIDTH="18" HEIGHT="18"></td><td nowrap title="系統內建使用者">&nbsp;系統內建使用者</td></tr>
</table>
<div id="systemusercont" style="width:100%; height:768;">	  
	  <%  	      Dim sqlstrQT As String
	      Dim dr As System.Data.DataRow
	      Dim object_uid As String
	      Dim object_name As String
	      
	      'sqlstrQT = "select object_uid, object_name,display_order from SYSTEMOBJ where object_type='系統內建使用者' order by display_order"
	      sqlstrQT = "select object_uid, object_name,display_order from SYSTEMOBJ  order by display_order"
	     
	      If do_sql.db_sql(sqlstrQT, do_sql.G_conn_string) = True Then
	          For Each dr In do_sql.G_table.Rows
	              object_uid = Trim(dr("object_uid").ToString)
	              object_name = Trim(dr("object_name").ToString)
	              Response.Write("<TABLE border=1 cellspacing=0 cellpadding=0 ><TR><TD><img ")
	              Response.Write("src='../../image/BTsign.gif'")
	              Response.Write(" border=1 height=20></TD><TD><img src='../../image/customer.gif' border=0 align=absmiddle width=16 height=15>&nbsp;</TD><TD nowrap valign=baseline id=Item_" & object_uid & "_4 onmousedown='MouseDoSel()' title='" & object_name & "' style='cursor:hand'>" & object_name & "</td></tr></TABLE>")
	          Next
	      End If
	      	  
	  
	  %>
  <table border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><img src="../../image/BLsign.gif" border="0" height="20"></td>
      <td><img src="../../image/leader.gif" border="0" align="absmiddle" width="16" height="15">&nbsp;</td>
      <td nowrap valign="baseline" id="Item_L_LDR" onmousedown="MouseDoSel()" title="階層主管" style="cursor:hand">階層主管</td>
    </tr></table>
</div>	  
   </div>
  
</div>


<script language="javascript">
var curElement;
var notcorr;
var faketitle;
var PY = 0;
var PX = 0;
var apk = 0;
var onset = 0;
var onset0 = 0;
var objTextRange;
objTextRange = document.body.createTextRange();

//document.onmousedown = doMouseDown;
//window.onresize = doresize;
window.onscroll = doscroll;
function ini(){
      //doresize();
      IX= 0;
      IY= 0;
      window.scrollTo(IX,IY);
}
//function doresize(){
//      down2.style.pixelTop = document.body.offsetHeight - 30;
//}
function doscroll(){
      PY = document.body.scrollTop;
      PX = document.body.scrollLeft;
      folder.style.top  = PY;
      folder.style.left = PX + xt ;
      form1.IX.value = PX;
      form1.IY.value = PY;
}
/*function showDown2() {
    if (inFolder.offsetHeight - apk > (document.body.offsetHeight - 100)){
      down2.style.display = '';
    }
    else{
      down2.style.display = 'none';
    }
} */
function reflow(){
    if (confirm('確定要重新規劃此表單內容?')){
          document.all.MODE.value='NEW';
          document.all.form1.target='FlowRight';
          document.form1.submit();
   }
}
function smallwin() {
    curElement0 = event.srcElement.id
    curElement1 = curElement0.substring(1, curElement0.length)
    var iTopPosition=event.clientY + PY-5;
    if (event.clientY+120 > document.body.offsetHeight) 
        iTopPosition=document.body.offsetHeight +PY - 125;
}
function smallwin0() {
  onset0 = 1;
  smallwin();
}
function dset() {
    onset = 1;
}
function base() {
    form1.STEPSID.value = curElement1;
    form1.action='flw102cf.asp';
    form1.submit();
}
function crit() {
    location.href ='flw108cf.asp?eformid='+document.all.eformid.value+'&stepsid=' + curElement1+'&organization_id='+document.all.organization_id.value+'&frm_chinese_name='+document.all.frm_chinese_name.value
}
/*function doMouseDown() {
      setTimeout('showDown2()',300);
} */
function MouseDoSel(){
	if (!parent.FlowRight.document.all.group_id)
		return;
    curElement = event.srcElement;
    parent.FlowRight.curElement=curElement;
    parent.FlowRight.document.all.group_id.value = curElement.id.substr(curElement.id.indexOf('_')+1);
    parent.FlowRight.document.all.group_name.value = curElement.innerText;
    faketitle = curElement.innerText;
    parent.FlowRight.faketitle=faketitle;
    fake.innerHTML = faketitle;
    parent.FlowRight.fake.innerHTML = faketitle;
    objTextRange.moveToElementText(curElement);
    objTextRange.select();

   //window.event.dataTransfer.setData("Text", "");
}
function fOnDragStart() {
  if (!window.event.dataTransfer) {
      alert("This version of IE doesn't support Drag and Drop!");
      return false;
  }
   doMouseMove();
   window.event.dataTransfer.setData("Text", objTextRange.text);
   event.dataTransfer.effectAllowed="all";
}
function fOnDragOver() {
   event.returnValue = false;
   event.dataTransfer.dropEffect="move";
   doMouseMove();
}
function fOnDragEnd() {
   parent.FlowRight.fake.style.display='none';
   fake.style.display='none';
   curElement = null;
   parent.FlowRight.curElement=null;
   event.dataTransfer.clearData("Text");
}
function doMouseMove() {
   if (curElement!=null){
      parent.FlowRight.fake.style.display='none';
      fake.style.posLeft = window.event.clientX + PX;
      fake.style.posTop = window.event.clientY + PY;
      notcorr = 1
      fake.style.backgroundColor = 'ffcccc'
      fake.innerHTML = faketitle 
      fake.style.display = '';
   }
   //event.returnValue = false;
}
function doMouseUp() {
    fake.style.display = 'none';
    parent.FlowRight.fake.style.display = 'none';
    if (curElement!=null){
        curElement = null;
        parent.FlowRight.curElement=null;
    }
}
/*function addy(){
   ik = apk;
    addygo();
}
function addygo(){
    k = apk;
    inFolder.style.clip = 'rect(' + k  + ' 100% ' + (k+500) + ' 0)';
    inFolder.style.top = (70 - k)
    if ((k < ik + 150) && (k<(inFolder.offsetHeight)-100)){
        apk = (k + 5);
        up2.style.display = '';
        setTimeout('addygo()',10);
    }
    if (apk>=inFolder.offsetHeight-100){
        down2.style.display = 'none';
    }
}
function miny(){
    ih = apk;
    minygo();
}
function minygo(){
    h = apk
    inFolder.style.clip = 'rect(' + h + ' 100% ' + (h+500) + ' 0)';
    inFolder.style.top = (70 - h)
    if ((h > ih - 150) && ( h >= 0)){
        apk = (h - 5);
        down2.style.display = '';
        setTimeout('minygo()',10);
    }
    if (apk<=0){
        up2.style.display = 'none';
    }
} */
var xt = 0 ;
function movein(){
    xt = xt - 5;
    if (xt>-155){
        folder.style.left = xt + PX;
        setTimeout('movein()',1);
    }
    else {
    slider1.style.display = '';
    }
}
function moveout(){
    slider1.style.display = 'none';
    xt = xt + 5 ;
    if (xt< 0){
        folder.style.left = xt + PX;
        setTimeout('moveout()',1);
    }
}
var xs = 0 ;
function movein2(){
    xs = xs - 1;
    if (xs>-18){
        slider1.style.left = xs;
        setTimeout('movein2()',25);
    }
}
function moveout2(){
    xs = xs + 1;
    if (xs <= 0){
        slider1.style.left = xs;
        setTimeout('moveout2()',25);
    }
}
var img1, img2;
img1 = new Image();
img1.src = '../../image/plus.gif';
img2 = new Image();
img2.src = '../../image/minus.gif';
function Changefolder() {
  var targetId, srcElement, targetElement;
  srcElement = window.event.srcElement;
  if (srcElement.className == 'folder') {
     targetId = srcElement.id + 'd';
     targetElement = document.all(targetId);
     if (targetElement.style.display == 'none') {
        targetElement.style.display = '';
        if (srcElement.tagName == 'IMG') {
          srcElement.src = '../../image/minus.gif';
        }
     } else {
        targetElement.style.display = 'none';
        if (srcElement.tagName == 'IMG') {
            srcElement.src = '../../image/plus.gif';
        }
     }
  }
}
folder.onclick = Changefolder;
</script>
</body>
</html>
