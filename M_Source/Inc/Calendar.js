    var x=0;
    var y=0;
    function SetPos(e){
        x=e.screenX;
        y=e.screenY;
    }
    function Calendar(d){            
        sPath = "../Inc/calendar.aspx?TextBoxId="+d+"&date="+document.all.item(d).value;
        strFeatures = "dialogWidth=260px;dialogHeight=180px;help=no;status=no;resizable=yes;scroll=no;dialogTop="+y+";dialogLeft="+x;
	    sDate = showModalDialog(sPath,self,strFeatures);
	    
    }
function chg_Fdate(d)
{   var  aryChk;
     
    aryChk = d.split("/") ;    
    if (aryChk[1].length ==1)
      aryChk[1] = "0" + aryChk[1]; 
     
    if (aryChk[2].length ==1)
       aryChk[2] = "0" + aryChk[2]; 
    return aryChk[0] + "/" + aryChk[1] + "/"+ aryChk[2];
}