var initRight=10;
var newWidth;
var nTimer;
function DialogDown(src) {
   document.all("Itemifrm").src=src;
   newWidth=Itemifrmdiv.style.width;
   initRight=10;
   //var clipValue='rect(0,'+newWidth+','+initRight+',0)'
   //Itemifrmdiv.style.clip=clipValue;
   //ItemifrmdivShadow.style.clip=clipValue;
   //原來的 Itemifrmdiv.style.pixelTop=(document.body.offsetWidth - Itemifrmdiv.style.pixelHeight - 8) /2  + document.body.scrollTop;
   Itemifrmdiv.style.pixelTop=((top.document.body.clientHeight - Itemifrmdiv.style.pixelHeight) * 2) / 3  + parent.document.body.scrollTop - 119; // 為x-flow 表單程式修改
   Itemifrmdiv.style.pixelLeft=(document.body.offsetWidth - Itemifrmdiv.style.pixelWidth)  / 2;
   // #pochin modified 2004/05/25 -- begin
   if (document.body.style.pixelTop > Itemifrmdiv.style.pixelTop)
		Itemifrmdiv.style.pixelTop = document.body.style.pixelTop;
   if (document.body.style.pixelLeft > Itemifrmdiv.style.pixelLeft)
		Itemifrmdiv.style.pixelLeft = document.body.style.pixelLeft;
   // #pochin modified 2004/05/25 -- end
   Itemifrmdiv.style.display='';
   ItemifrmdivShadow.style.width=Itemifrmdiv.style.width;
   ItemifrmdivShadow.style.height=Itemifrmdiv.style.height;
   ItemifrmdivShadow.style.pixelTop=Itemifrmdiv.style.pixelTop+8;
   ItemifrmdivShadow.style.pixelLeft=Itemifrmdiv.style.pixelLeft+6;
   ItemifrmdivShadow.style.display='';
   //nTimer=window.setInterval("EnlargeFrame()", 10);
}
//2004/10/07 -- Sam Add：將iframe跟著scroll bar的位置變動
function DialogDown1(src,rowIndex) {
   document.all("Itemifrm").src=src;
   newWidth=Itemifrmdiv.style.width;
   initRight=10;

   Itemifrmdiv.style.pixelTop=((top.document.body.clientHeight - Itemifrmdiv.style.pixelHeight) * 2) / 3  + rowIndex * 30 - 119; // 為x-flow 表單程式修改
   Itemifrmdiv.style.pixelLeft=(document.body.offsetWidth - Itemifrmdiv.style.pixelWidth)  / 2;

   if (document.body.style.pixelTop > Itemifrmdiv.style.pixelTop)
		Itemifrmdiv.style.pixelTop = document.body.style.pixelTop;
   if (document.body.style.pixelLeft > Itemifrmdiv.style.pixelLeft)
		Itemifrmdiv.style.pixelLeft = document.body.style.pixelLeft;

   // #pochin modified 2004/05/25 -- end
   Itemifrmdiv.style.display='';
   ItemifrmdivShadow.style.width=Itemifrmdiv.style.width;
   ItemifrmdivShadow.style.height=Itemifrmdiv.style.height;
   ItemifrmdivShadow.style.pixelTop=Itemifrmdiv.style.pixelTop+8;
   ItemifrmdivShadow.style.pixelLeft=Itemifrmdiv.style.pixelLeft+6;
   ItemifrmdivShadow.style.display='';
}
function EnlargeFrame() {
   initRight+=10;
   var clipValue='rect(0,'+newWidth+','+initRight+',0)'
   Itemifrmdiv.style.clip=clipValue;
   ItemifrmdivShadow.style.clip=clipValue;
   if (initRight > Itemifrmdiv.style.pixelHeight) { 
      window.clearInterval(nTimer);
   }
}
function CloseItemFrameDialog() {
   //var newtopPos=0 - window.screen.width + window.screenLeft;
   //nTimer=window.setInterval("MoveLeftFrameDialog("+escape(newtopPos)+")", 10);
   // add below lines for closing dialog
   Itemifrmdiv.style.display='none';
   ItemifrmdivShadow.style.display='none';
   document.all("Itemifrm").src='';
}
function MoveLeftFrameDialog(newtopPos) {
   Itemifrmdiv.style.posLeft=Itemifrmdiv.style.posLeft-30;
   ItemifrmdivShadow.style.posLeft=Itemifrmdiv.style.posLeft+6;
   if (Itemifrmdiv.style.posLeft < newtopPos) { 
	   bWaitCloseDialog=true;
       window.clearInterval(nTimer);
   }
}
