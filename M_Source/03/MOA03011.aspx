<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA03011.aspx.vb" Inherits="M_Source_03_MOA03011" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>派車申請單列印</title>
    <script type="text/javascript">
        window.onload = function () {
            window.print();
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <center style="font-family:標楷體; font-size:large;"><strong>國軍履帶輪型車輛(機車)運輸申請單</strong></center>
    <center>
    <table border="1" cellpadding="0" style="border:1px solid Black; border-collapse:collapse; font-family:標楷體; font-size:small; width:680px;">
    <tr>
        <td align="center" style="width:80%;">
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
            <tr>
                <td align="center" style="width:25%; border-right:1px solid Black; height:20px;">接 受 申 請 時 間</td>
                <td align="center" style="width:40%; border-right:1px solid Black;">
                    <asp:Label ID="lbApplyYear" runat="server" Width="15px"></asp:Label>
                    年 <asp:Label ID="lbApplyMonth" runat="server" Width="10px"></asp:Label>
                    月 <asp:Label ID="lbApplyDay" runat="server" Width="10px"></asp:Label>
                    日 <asp:Label ID="lbApplyHour" runat="server" Width="10px"></asp:Label>
                    時 <asp:Label ID="lbApplyMinute" runat="server" Width="15px"></asp:Label> 分
                </td>
                <td align="center" style="width:20%; border-right:1px solid Black;">單位派車編號</td>
                <td align="center"></td>
            </tr>
            </table>
        </td>
        <td align="center">部&nbsp; 隊&nbsp; 長&nbsp; 批&nbsp; 示</td>
    </tr>
    <tr>
        <td>
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
            <tr>
                <td align="center" style="width:15%; border-right:1px solid Black; height:20px;">任 務 理 由</td>
                <td align="left">
                    &nbsp;<asp:Label ID="lbReason" runat="server"></asp:Label>
                </td>
            </tr>
            </table>
        </td>
        <td rowspan="6">
            <table border="0" style="width:100%; height:99px;">
            <tr><td align="right" style="font-size:x-large;">照派&nbsp;</td></tr>
            <tr><td align="center"><asp:Label ID="lbSuperiorDept2" runat="server"></asp:Label></td></tr>
            <tr><td align="center"><asp:Label ID="lbSuperiorName2" runat="server"></asp:Label></td></tr>
            <tr><td align="center"><asp:Label ID="lbNow3" runat="server"></asp:Label></td></tr>
            </table>
        </td>
    </tr>
    <tr>
        <td align="center" style="height:20px;">任 務 使 用 時 間 自 <asp:Label ID="lbUseStartMonth" runat="server"></asp:Label> 
            月 <asp:Label ID="lbUseStartDay" runat="server"></asp:Label> 
            日 <asp:Label ID="lbUseStartHour" runat="server"></asp:Label> 
            時 <asp:Label ID="lbUseStartMinute" runat="server"></asp:Label> 
            分 起 至 <asp:Label ID="lbUseEndMonth" runat="server"></asp:Label> 
            月 <asp:Label ID="lbUseEndDay" runat="server"></asp:Label> 
            日 <asp:Label ID="lbUseEndHour" runat="server"></asp:Label> 
            時 <asp:Label ID="lbUseEndMinute" runat="server"></asp:Label> 
            分 止</td>
    </tr>
    <tr>
        <td align="center">
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
            <tr>
                <td align="center" style="width:15%; border-right:1px solid Black; height:20px;">人 員 項 目</td>
                <td align="left" style="width:15%; border-right:1px solid Black;">
                    &nbsp;<asp:Label ID="lbItem" runat="server"></asp:Label>&nbsp;員
                </td>
                <td align="center" style="width:30%; border-right:1px solid Black;">軍品種類、重量及體積</td>
                <td align="left">
                    &nbsp;</td>
            </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td align="center">
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
            <tr>
                <td align="center" style="width:23%; border-right:1px solid Black; height:20px;">車 輛 報 到 地 點</td>
                <td align="left" style="width:20%; border-right:1px solid Black;">
                    &nbsp;<asp:Label ID="lbArrivePlace" runat="server"></asp:Label>
                </td>
                <td align="center" style="width:19%; border-right:1px solid Black;">向 何 人 報 到</td>
                <td align="left" style="width:20%; border-right:1px solid Black;">
                    &nbsp;<asp:Label ID="lbnArriveTo" runat="server"></asp:Label>
                </td>
                <td align="center" style="width:11%; border-right:1px solid Black;">車 次 數</td>
                <td align="left"></td>
            </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td align="center">
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
            <tr>
                <td align="center" style="width:55%; border-right:1px solid Black; height:20px;">車 輛 報 到 日 期 <asp:Label 
                        ID="lbArriveMonth" runat="server"></asp:Label> 
                    月 <asp:Label ID="lbArriveDay" runat="server"></asp:Label> 
                    日 <asp:Label ID="lbArriveHour" runat="server"></asp:Label> 
                    時 <asp:Label ID="lbArriveMinute" runat="server"></asp:Label> 分
                </td>
                <td align="center" style="width:20%; border-right:1px solid Black;">使 用 車 輛 數</td>
                <td align="left"></td>
            </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td align="center">
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
            <tr>
                <td align="center" style="width:35%; border-right:1px solid Black; height:20px;">使 用 車 輛 品 名 及 型 式</td>
                <td align="left" style="width:25%; border-right:1px solid Black;"> 
                    &nbsp;<asp:Label ID="lbStyle" runat="server"></asp:Label>
                </td>
                <td align="center" style="width:20%; border-right:1px solid Black;">全 程 里 程 數</td>
                <td align="left"></td>
            </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td align="center">
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%; height:100%;">
            <tr>
                <td align="center" style="width:8%; border-right:1px solid Black; height:20px;">地 點</td>
                <td align="left" style="width:25%; border-right:1px solid Black;"> 
                    &nbsp;<asp:Label ID="lbStartPoint" runat="server"></asp:Label>
                </td>
                <td align="center" style="width:11%; border-right:1px solid Black;">目 的 地</td>
                <td align="left" style="width:25%; border-right:1px solid Black;"> 
                    &nbsp;<asp:Label ID="lbEndPoint" runat="server"></asp:Label>
                </td>
                <td align="center" style="width:15%; border-right:1px solid Black;">使 用 油 料</td>
                <td align="right">加侖</td>
            </tr>
            </table>
        </td>
        <td align="center">調派軍官審核意見</td>
    </tr>
    <tr>
        <td align="center">
            <table border="0" cellpadding="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
            <tr>
                <td align="center" style="height:20px;">預&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 定&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 行&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 程</td>
            </tr>
            <tr>
        <td align="center" style="border-top:1px solid Black;">
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
            <tr>
                <td align="center" style="width:33%; border-right:1px solid Black; height:20px;">從</td>
                <td align="center" style="width:33%; border-right:1px solid Black;">至</td>
                <td align="center">停留時間</td>
            </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td align="center" style="border-top:1px solid Black;">
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
            <tr>
                <td align="left" style="width:33%; border-right:1px solid Black; height:20px;">
                    &nbsp;<asp:Label ID="lbGoLocal1" runat="server"></asp:Label>&nbsp;
                </td>
                <td align="left" style="width:33%; border-right:1px solid Black;">
                    &nbsp;<asp:Label ID="lbEndLocal1" runat="server"></asp:Label>&nbsp;
                </td>
                <td align="left"></td>
            </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td align="center" style="border-top:1px solid Black;">
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
            <tr>
                <td align="left" style="width:33%; border-right:1px solid Black; height:20px;">
                    &nbsp;<asp:Label ID="lbGoLocal2" runat="server"></asp:Label>&nbsp;
                </td>
                <td align="left" style="width:33%; border-right:1px solid Black;">
                    &nbsp;<asp:Label ID="lbEndLocal2" runat="server"></asp:Label>&nbsp;
                </td>
                <td align="left"></td>
            </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td align="center" style="border-top:1px solid Black;">
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
            <tr>
                <td align="left" style="width:33%; border-right:1px solid Black; height:20px;">
                    &nbsp;<asp:Label ID="lbGoLocal3" runat="server"></asp:Label>&nbsp;
                </td>
                <td align="left" style="width:33%; border-right:1px solid Black;">
                    &nbsp;<asp:Label ID="lbEndLocal3" runat="server"></asp:Label>&nbsp;
                </td>
                <td align="left"></td>
            </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td align="center" style="border-top:1px solid Black;">
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
            <tr>
                <td align="left" style="width:33%; border-right:1px solid Black; height:20px;">
                    &nbsp;<asp:Label ID="lbGoLocal4" runat="server"></asp:Label>&nbsp;
                </td>
                <td align="left" style="width:33%; border-right:1px solid Black;">
                    &nbsp;<asp:Label ID="lbEndLocal4" runat="server"></asp:Label>&nbsp;
                </td>
                <td align="left"></td>
            </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td align="center" style="border-top:1px solid Black;">
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
            <tr>
                <td align="left" style="width:33%; border-right:1px solid Black; height:20px;">
                    &nbsp;<asp:Label ID="lbGoLocal5" runat="server"></asp:Label>&nbsp;
                </td>
                <td align="left" style="width:33%; border-right:1px solid Black;">
                    &nbsp;<asp:Label ID="lbEndLocal5" runat="server"></asp:Label>&nbsp;
                </td>
                <td align="left"></td>
            </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td align="left" style="border-top:1px solid Black;">
            <table border="0" style="font-family:標楷體; width:100%;">
            <tr>
                <td style="width:35%; height:40px;">申請單位主管級職姓名(簽章)</td>
                <td>
                    <asp:Label ID="lbSuperiorDept1" runat="server"></asp:Label>
                    &nbsp;<asp:Label ID="lbSuperiorName1" runat="server"></asp:Label>
                    &nbsp;<asp:Label ID="lbNow1" runat="server"></asp:Label>
                </td>
            </tr>
            </table>           
        </td>
    </tr>
    <tr>
        <td align="left" style="border-top:1px solid Black;">
            <table border="0" style="font-family:標楷體; width:100%;">
            <tr>
                <td style="width:32%; height:40px;">申請人單位級職姓名(簽章)</td>
                <td>
                    <asp:Label ID="lbPADept" runat="server"></asp:Label>
                    &nbsp;<asp:Label ID="lbPAName" runat="server"></asp:Label>
                    &nbsp;<asp:Label ID="lbNow2" runat="server"></asp:Label>
                </td>
            </tr>
            </table> 
        </td>
    </tr>
    <tr>
        <td align="center" style="border-top:1px solid Black;">
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
            <tr>
                <td align="center" style="width:28%; border-right:1px solid Black; height:20px;">駕 駛 人 級 職 姓 名</td>
                <td align="center" style="width:40%; border-right:1px solid Black;"></td>
                <td align="center" style="width:9%; border-right:1px solid Black;">車 號</td>
                <td align="center"></td>
            </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td align="center" style="border-top:1px solid Black;">
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%; height:90px;">
            <tr>
                <td align="center">使 用 人 注 意 事 項</td>
            </tr>
            <tr>
                <td align="left">一、負有對駕駛手監督行車安全之責任。</td>
            </tr>
            <tr>
                <td align="left">二、督飭駕駛手作行駛中及停駛時之一級保養。</td>
            </tr>
            <tr>
                <td align="left">三、嚴格要求駕駛手遵守交通規則，並負駕駛手任務中食宿之責。</td>
            </tr>
            </table>
        </td>
    </tr>
            </table>
        </td>
        <td align="center">
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
            <tr>
                <td align="center">
                    <table border="0" style="height:90px; width:100%">
                    <tr><td align="right" style="font-size:x-large;"><asp:Label ID="lbTransfereeOpnion2" runat="server" Text=""></asp:Label>&nbsp;</td></tr>
                    <tr><td align="center"><asp:Label ID="lbTransfereeDept2" runat="server" Text=""></asp:Label></td></tr>
                    <tr><td align="center"><asp:Label ID="lbTransfereeName2" runat="server" Text=""></asp:Label></td></tr>
                    <tr><td align="center"><asp:Label ID="lbTransfereeTime2" runat="server" Text=""></asp:Label></td></tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center" style="border-top:1px dashed Black;">
                <table border="0" style="height:90px; width:100%">
                <tr><td align="right" style="font-size:x-large;"><asp:Label ID="lbTransfereeOpnion1" runat="server" Text=""></asp:Label>&nbsp;</td></tr>
                <tr><td align="center"><asp:Label ID="lbTransfereeDept1" runat="server" Text=""></asp:Label></td></tr>
                <tr><td align="center"><asp:Label ID="lbTransfereeName1" runat="server" Text=""></asp:Label></td></tr>
                <tr><td align="center"><asp:Label ID="lbTransfereeTime1" runat="server" Text=""></asp:Label></td></tr>
                </table>
                </td>
            </tr>
            <tr>
                <td align="center" style="border-top:1px dashed Black;">
                <table border="0" style="height:90px; width:100%">
                <tr><td align="right" style="font-size:x-large;"><asp:Label ID="lbControlOpnion" runat="server" Text=""></asp:Label>&nbsp;</td></tr>
                <tr><td align="center"><asp:Label ID="lbControlDept" runat="server" Text=""></asp:Label></td></tr>
                <tr><td align="center"><asp:Label ID="lbControlName" runat="server" Text=""></asp:Label></td></tr>
                <tr><td align="center"><asp:Label ID="lbControlTime" runat="server" Text=""></asp:Label></td></tr>
                </table>    
                </td>
            </tr>
            <tr>
                <td align="center" style="border-top:1px dashed Black;">
                <table border="0" style="height:90px; width:100%">
                <tr><td align="right" style="font-size:x-large;"><asp:Label ID="Label1" runat="server" Text=""></asp:Label>&nbsp;</td></tr>
                <tr><td align="center"><asp:Label ID="lbDispatchDept" runat="server" Text=""></asp:Label></td></tr>
                <tr><td align="center"><asp:Label ID="lbDispatchName" runat="server" Text=""></asp:Label></td></tr>
                <tr><td align="center"><asp:Label ID="lbDispatchTime" runat="server" Text=""></asp:Label></td></tr>
                </table>   
                </td>
            </tr>
            </table>
        </td>
    </tr>
    </table>
    <br />
    <table border="1" cellpadding="0" style="border:1px solid Black; border-collapse:collapse; font-family:標楷體; font-size:small; width:680px;">
    <tr>
        <td align="center" colspan="4" style="font-size:medium; height:25px;"><strong>行 車 安 全 保 證 責 任 檢 查 表</strong></td>
    </tr>
    <tr>
        <td align="center" style="width:12%; height:20px;">區&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 分</td>
        <td align="center" style="width:45%;">保&nbsp;&nbsp;&nbsp; 證&nbsp;&nbsp;&nbsp; 責&nbsp;&nbsp;&nbsp; 任&nbsp;&nbsp;&nbsp; 要&nbsp;&nbsp;&nbsp; 項</td>
        <td align="center" style="width:20%;">簽&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 名</td>
        <td align="center">年&nbsp;&nbsp; 月&nbsp;&nbsp; 日</td>
    </tr>
    <tr>
        <td align="center">用 車 單 位<br />主 (管) 官</td>
        <td>
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
                <tr><td align="left" style="height:20px;">一、申請車輛確實需要。</td></tr>
                <tr><td align="left" style="border-top:1px solid Black; height:20px;">二、指派車長，當面叮囑。</td></tr>
            </table>
        </td>
        <td></td>
        <td></td>
    </tr>
    <tr>
        <td align="center">調&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 度<br /><br />軍 (士) 官</td>
        <td>
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
                <tr><td align="left" style="height:20px;">一、行車資料，證照要齊全。</td></tr>
                <tr><td align="left" style="border-top:1px solid Black; height:20px;">二、車容、車況均良好。</td></tr>
                <tr><td align="left" style="border-top:1px solid Black; height:20px;">三、駕駛身心健康，情緒正常。</td></tr>
            </table>
        </td>
        <td></td>
        <td></td>
    </tr>
    <tr>
        <td align="center">車&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 輛<br /><br />駕&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 駛</td>
        <td>
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
                <tr><td align="left" style="height:20px;">一、完成車輛行駛前檢查保養。</td></tr>
                <tr><td align="left" style="border-top:1px solid Black; height:20px;">二、清點行車資料、證照。</td></tr>
                <tr><td align="left" style="border-top:1px solid Black; height:20px;">三、服從車長領導，遵守交通規則。</td></tr>
                <tr><td align="left" style="border-top:1px solid Black; height:20px;">四、指導人員，軍品乘載整齊。</td></tr>
            </table>
        </td>
        <td></td>
        <td></td>
    </tr>
    <tr>
        <td align="center">保&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 養<br />軍 (士) 官</td>
        <td>
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
                <tr><td align="left" style="height:20px;">一、實施車輛出場前安全檢查。</td></tr>
                <tr><td align="left" style="border-top:1px solid Black; height:20px;">二、剎車、轉向、燈光、儀表、雨刷均良好。</td></tr>
            </table>
        </td>
        <td></td>
        <td></td>
    </tr>
    <tr>
        <td align="center">車&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 長<br />(副 車 長)</td>
        <td>
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
                <tr><td align="left" style="height:20px;">一、監督駕駛遵守交通規則。</td></tr>
                <tr><td align="left" style="border-top:1px solid Black; height:20px;">二、維護人員、軍品乘載紀律與秩序。</td></tr>
                <tr><td align="left" style="border-top:1px solid Black; height:20px;">三、負責車輛行駛安全。</td></tr>
            </table>
        </td>
        <td></td>
        <td></td>
    </tr>
    <tr>
        <td align="center">營&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 門<br /><br />衛&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 兵</td>
        <td>
            <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
                <tr><td align="left" style="height:20px;">一、檢查申請單確實批准。</td></tr>
                <tr><td align="left" style="border-top:1px solid Black; height:20px;">二、檢查行車安全保證責任表均已簽證。</td></tr>
                <tr><td align="left" style="border-top:1px solid Black; height:20px;">三、檢查車容及乘載情形良好。</td></tr>
            </table>
        </td>
        <td></td>
        <td></td>
    </tr>
    </table>
    </center>
    </div>
    </form>
</body>
</html>

