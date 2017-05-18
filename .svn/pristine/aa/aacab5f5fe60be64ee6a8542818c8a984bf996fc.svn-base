<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Frame.aspx.vb" Inherits="OA_Frame" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >

<SCRIPT LANGUAGE="JavaScript">
function scroll(seed)
{
    var m1 = "國防部辦公室自動化系統 "
    var m2 = " 最佳顯示解析度建議：1024x768 瀏覽器：IE6.0以上..."
    var msg=m1+m2;
    var out = " ";
    var c = 1;
    if (seed > 100) {
        seed--;
        var cmd="scroll(" + seed + ")";
        timerTwo=window.setTimeout(cmd,100);
    }
    else if (seed <= 100 && seed > 0) {
        for (c=0 ; c < seed ; c++) {
        out+=" ";
    }
        out+=msg;
        seed--;
    var cmd="scroll(" + seed + ")";
    window.status=out;
    timerTwo=window.setTimeout(cmd,100);
    }
    else if (seed <= 0) {
        if (-seed < msg.length) {
            out+=msg.substring(-seed,msg.length);
            seed--;
            var cmd="scroll(" + seed + ")";
            window.status=out;
            timerTwo=window.setTimeout(cmd,100);
        }
    else{
        window.status=" ";
        timerTwo=window.setTimeout("scroll(100)",7);
        }
    }
}
timerONE=window.setTimeout('scroll(100)',50);
</SCRIPT>
<head runat="server">
    <title>行政系統</title>
</head>
	<frameset id="frameset" framespacing="0" border="0" rows="70,*" frameborder="0">
		<frame name="above" src="above.aspx" scrolling="no" noresize>
		<frameset id="mainframecol" cols="205,12,*">
			<frame id="left" name="left" src="Menu.aspx" scrolling="yes">
			<frame id="slider" name="slider" src="frameslider.aspx" scrolling="no" noresize>
			<frame id="right" name="right" src="right.aspx" scrolling="yes">
		</frameset>
		<noframes>
			<body>
				<p>此網頁使用框架,但是您的瀏覽器並不支援.</p>
			</body>
		</noframes>
	</frameset>
</html>
