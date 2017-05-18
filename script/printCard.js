$(document).ready(function () {
    $("#ddl_SecurityItem").change(function () {
        var strHref = "../08/MOA08010.aspx?Security_GuidID=" + $("#ddl_SecurityItem").children('option:selected').val() + "&Security_Level=" + $("#ddl_Security_Status").children('option:selected').val();
        $("#HLSecurity").attr("href", strHref);
    });
});


//-1:查無資料,表示未申請過;0:清除此sn ;1:新申請 ;2:已印但未回登 ;3:印列失敗 ;4:補登完畢 ;5:審核不通過(密等為2~5者才有) 6:申請人取消
function checksmf() {
    try {
        var status = $("#hid_status").html(); //document.getElementById("hid_status").innerText;
        var readonly = $("#hid_readonly").html(); //document.getElementById("hid_readonly").innerText;
        var sstatus = $("#ddl_Security_Status").children('option:selected').val();
        var ret;
        ret = SmfCom.SmfInit();

        if (readonly == "" && ret != 0 && sstatus==1) {
            alert("請確認卡片是否有插入，或再次重插!!");
            return false;
        }
//        if(SmfCom.SmfId != $("#hidCurrentUserID").val()){
//            alert("卡片持有人與系統登入者不相符!!");
//            return false;
//        }
        if (readonly == "" && status == "1") {
            alert("您已申請Ticket，但尚未至影印設備使用，無法重覆申請！\r\n如果你要重新申請，請先按下[清除已申請]按鈕再進行申請作業!!");
            return true;
        }
        if (readonly == "" && status == "2") {
            alert("您已申請Ticket且已使用，但尚未回本系統補上列印資訊，無法重覆申請！");
            return false;
        }
        return true;
    }
    catch (err) {
        alert("請確認PrintCard元件是否有安裝!!");
        return false;
    }
}

function smfWriteTicket(sn) {
    try {
        var ret;
        ret = SmfCom.SmfWriteTicket(1, sn, 1, 1, 1);
        if (ret == 0) {
            alert('已寫入Ticket!!');
            parent.location.reload();
        }
        else {
            updateJson("smfWriteTicket");
//            if (msg == "true") {
//                alert("寫入卡片失敗，請確定讀卡機與卡片已Ready！"); //可在前端加Btn做再次寫入動作
//            }
//            else
//                alert("Json Failed:" + msg + "！")
        }
    }
    catch (err) {
        alert("請確認PrintCard元件是否有安裝!!");
    }
}

function clearsmf() {
    try {
        var ret;
        ret = SmfCom.SmfInit();
        if (ret != 0) {
            alert("請確認卡片是否有插入，或再次重插!!");
            return false;
        }
        ret = SmfCom.SmfCleanTicket();
        if (ret == 0) {           
            return true;
        }
        else {
            alert("清除卡片內Ticket失敗:" + ret);
            return false;
        }
    }
    catch (err) {
        alert("請確認PrintCard元件是否有安裝!!");
        return false;
    }
}

function smfverifyWriteTicket() {
    try {
        var sn = $("#hid_eformsn").html(); // document.getElementById("hid_eformsn").innerText;
        var ret;
        ret = SmfCom.SmfInit();
        if (ret != 0) {
            alert("請確認卡片是否有插入，或再次重插!!");
        }
        else {
            ret = 0;
            ret = SmfCom.SmfWriteTicket(2, sn, 1, 1, 1);
            if (ret == 0) {
                updateJson("smfverifyWriteTicket");
//             if ( msg== "true") {
//                alert("完成申請已寫入卡片，您可持卡進行影印作業！");
//                window.dialogArguments.location = "../00/MOA00010.aspx";
//                window.close();
//             }
//             else
//                alert("Json Failed:" + msg + "！")
            }
            else {
                alert("寫入卡片失敗，請確定讀卡機與卡片已Ready，請重新再試！");
            }
        }
        return false;
    }
    catch (err) {
        alert("請確認PrintCard元件是否有安裝!!");
        return false;
    }
}

//function wchecksmf() {
//    try {
//        var ret;
//        ret = SmfCom.SmfInit();
//        if (ret != 0) {
//            alert("請確認卡片是否有插入，或再次重插!!");
//            return false;
//        }
//        return true;
//    }
//    catch (err) {
//        alert("請確認PrintCard元件是否有安裝!!");
//        return false;
//    }
//}

//function verifyWriteTicket(sn) {
//    try {
//        var ret;
//        ret = SmfCom.SmfWriteTicket(2, sn, 1, 1, 1);
//        if (ret == 0) {
//            //json= "Update P_08 set WriteCard = 1 where EFORMSN='" + str_EFORMSN + "'"
//            alert("完成申請已寫入卡片，您可持卡進行影印作業！");
//            window.dialogArguments.location = "../00/MOA00010.aspx";
//            window.close();
//        }
//        else {
//            alert("寫入卡片失敗，請確定讀卡機與卡片已Ready，請重新再試！");
//        }
//    }
//    catch (err) {
//        alert("請確認PrintCard元件是否有安裝!!");
//    }
//}
function updateJson(action) {
    var sn = $("#hid_eformsn").html();    
    var request = $.ajax({
        url: "MOA08011.ashx",
        type: "POST",
        data: { sn: sn, action: action },
        dataType: "html",
        success: function (Jdata) {
            if ( Jdata== "smfverifyWriteTicket") {
                alert("完成申請已寫入卡片，您可持卡進行影印作業！");
                window.dialogArguments.location = "../00/MOA00011.aspx";
                window.close();
             }

            if (Jdata == "smfWriteTicket") {
                alert("寫入卡片失敗，請確定讀卡機與卡片已Ready！"); //可在前端加Btn做再次寫入動作
            }
        },
        error: function () {
            alert("JSON ERROR!!");
        }
    });    
}