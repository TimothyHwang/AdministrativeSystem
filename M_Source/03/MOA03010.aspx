<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA03010.aspx.vb" Inherits="M_Source_03_MOA03010" %>

<%@ Register Src="../90/FlowRoute.ascx" TagName="FlowRoute" TagPrefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>派車申請單列印</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
        
        <table width="700" border=1 cellspacing=0 cellpadding=5 bgcolor=#ffffff bordercolor=#6699cc bordercolorlight=#74a3d6 bordercolordark=#000000 style="LEFT: 10px; TOP: 10px">
        <tr>
            <td valign=bottom bgcolor=#6699cc bordercolorlight=#66aaaa bordercolordark=#ffffff style="width: 734px"><font color=white><b>&nbsp;派車申請單</b></font>
            </td>
        </tr>
        
        <tr>
            <td style="width: 690px">           
            
                <table style="width: 690px" rules="all">
                    <tr>
                        <td style="width: 240px">
                            <asp:Label ID="Label1" runat="server" Text="填表人單位：" Width="80px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_PWUNIT" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label>
                        </td>
                        <td style="width: 200px">
                            <asp:Label ID="Label2" runat="server" Text="姓名：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_PWNAME" runat="server" ForeColor="Black" Width="100px" CssClass="form"></asp:Label></td>
                        <td style="width: 230px">
                            <asp:Label ID="Label3" runat="server" Text="級職：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_PWTITLE" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label></td>
                    </tr>
                    <tr>
                        <td style="width: 240px">
                            <asp:Label ID="Label4" runat="server" Text="申請人單位：" Width="80px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_PAUNIT" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label></td>
                        <td style="width: 200px">
                            <asp:Label ID="Label5" runat="server" Text="姓名：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>&nbsp;<asp:Label
                                ID="Lab_PANAME" runat="server" CssClass="form" ForeColor="Black" Width="100px"></asp:Label>
                        </td>
                        <td style="width: 230px">
                            <asp:Label ID="Label6" runat="server" Text="級職：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                            <asp:Label ID="Lab_PATITLE" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label>
                        </td>
                    </tr>                
                </table>
             
                <table style="width: 690px" rules="all">
                    <tr>
                        <td style="width: 110px; height: 12px;">
                        <asp:Label ID="Label15" runat="server" ForeColor="Black" Text="申請時間：" CssClass="form" ></asp:Label></td>
                        <td style="width: 200px; height: 12px;">
                            <asp:Label ID="Lab_nAPPLYTIME" runat="server" ForeColor="Black" CssClass="form" Width="179px" ></asp:Label></td>
                        <td style="width: 100px; height: 12px;">
                            <asp:Label ID="Label7" runat="server" ForeColor="Black" Text="聯絡電話：" CssClass="form" ></asp:Label></td>
                        <td style="height: 12px; width: 210px;" >
                            <asp:Label ID="Lab_nPHONE" runat="server" CssClass="form" ForeColor="Black" Width="179px"></asp:Label></td>
                    </tr>                             

                    <tr>
                        <td style="height: 26px; width: 110px;" >
                            <asp:Label ID="Label12" runat="server" ForeColor="Black" Text="任務理由：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 200px;" >
                            <asp:Label ID="Lab_nREASON" runat="server" CssClass="form" ForeColor="Black" Width="179px"></asp:Label></td>
                        <td style="height: 26px; width: 100px;" >
                            <asp:Label ID="Label20" runat="server" ForeColor="Black" Text="人員項目：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 210px;" >
                            <asp:Label ID="Lab_nITEM" runat="server" CssClass="form" ForeColor="Black" Width="179px"></asp:Label></td>
                    </tr> 
                    <tr>
                        <td style="height: 26px; width: 110px;" >
                            <asp:Label ID="Label8" runat="server" ForeColor="Black" Text="車輛報到地點：" CssClass="form" Width="100px"></asp:Label></td>
                        <td style="height: 26px; width: 200px;" >
                            <asp:Label ID="Lab_nARRIVEPLACE" runat="server" CssClass="form" ForeColor="Black"
                                Width="179px"></asp:Label></td>
                        <td style="height: 26px; width: 100px;" >
                            <asp:Label ID="Label9" runat="server" ForeColor="Black" Text="向何人報到：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 210px;" >
                            <asp:Label ID="Lab_nARRIVETO" runat="server" CssClass="form" ForeColor="Black" Width="179px"></asp:Label></td>
                    </tr>
                    <tr>
                        <td style="height: 23px; width: 110px;">
                            <asp:Label ID="Label16" runat="server" ForeColor="Black" Text="車輛報到日期：" CssClass="form" ></asp:Label></td>
                        <td style="height: 23px; width: 200px;" >
                            <asp:Label ID="Lab_nARRDATE" runat="server" CssClass="form" ForeColor="Black" Width="80px"></asp:Label>
                            <asp:Label ID="Label55" runat="server" CssClass="form" ForeColor="Black" Text="日"
                                Width="20px"></asp:Label><asp:Label ID="Lab_nSTHOUR" runat="server" CssClass="form" ForeColor="Black" Width="20px"></asp:Label>
                            <asp:Label ID="Label17" runat="server" ForeColor="Black" Text="時" CssClass="form" Width="20px"></asp:Label><asp:Label ID="Lab_nEDHOUR" runat="server" CssClass="form"
                                    ForeColor="Black" Width="20px"></asp:Label>
                            <asp:Label ID="Label18" runat="server" ForeColor="Black" Text="分" CssClass="form" Width="20px"></asp:Label></td>
                        <td style="height: 26px; width: 100px;" >
                        <asp:Label ID="Label21" runat="server" ForeColor="Black" Text="車輛品名型式：" CssClass="form" ></asp:Label></td>
                        <td style="height: 26px; width: 210px;" >
                            <asp:Label ID="Lab_nSTYLE" runat="server" CssClass="form" ForeColor="Black" Width="179px"></asp:Label></td>
                    </tr> 
                    <tr>
                        <td style="height: 26px; width: 110px;" >
                            <asp:Label ID="Label13" runat="server" ForeColor="Black" Text="起點：" CssClass="form" Width="80px"></asp:Label></td>
                        <td style="height: 26px; width: 200px;" >
                            <asp:Label ID="Lab_nSTARTPOINT" runat="server" CssClass="form" ForeColor="Black"
                                Width="179px"></asp:Label></td>
                        <td style="height: 26px; width: 100px;" >
                            <asp:Label ID="Label14" runat="server" ForeColor="Black" Text="目的地：" CssClass="form"></asp:Label></td>
                        <td style="height: 26px; width: 210px;" >
                            <asp:Label ID="Lab_nENDPOINT" runat="server" CssClass="form" ForeColor="Black" Width="179px"></asp:Label></td>
                    </tr> 

                    <tr>
                        <td style="width: 110px; height: 23px;">
                        <asp:Label ID="Label19" runat="server" ForeColor="Black" Text="車輛類型：" CssClass="form" ></asp:Label></td>
                        <td style="width: 200px; height: 23px;">
                            <asp:Label ID="Lab_nSTATUS" runat="server" CssClass="form" ForeColor="Black" Width="179px"></asp:Label></td>
                        <td style="width: 100px; height: 23px;"><asp:Label ID="LabCarTitle" runat="server"
                            CssClass="form" ForeColor="Black" Text="車號：" Width="70px"></asp:Label></td>    
                        <td style="width: 210px; height: 23px;">
                        <asp:Label ID="Lab_nCarNum" runat="server" CssClass="form" ForeColor="Black" Visible="False"
                            Width="180px"></asp:Label></td>                  
                    </tr>
                    <tr>
                        <td style="height: 22px; width: 110px;">
                        <asp:Label ID="Label22" runat="server" ForeColor="Black" Text="任務使用時間：" CssClass="form" ></asp:Label></td>
                        <td style="width: 560px;height: 22px;" colspan="3" >
                            <asp:Label ID="Lab_nUSEDATE" runat="server" CssClass="form" ForeColor="Black" Width="100px"></asp:Label>
                            <asp:Label
                            ID="Label29" runat="server" CssClass="form" ForeColor="Black" Text="日" Width="20px"></asp:Label>
                            <asp:Label ID="Lab_nSTUSEHOUR" runat="server" CssClass="form" ForeColor="Black" Width="40px"></asp:Label>
                            <asp:Label ID="Label23" runat="server" ForeColor="Black" Text="時" CssClass="form" Width="20px"></asp:Label>
                            <asp:Label ID="Lab_nSTUSEMIN" runat="server" CssClass="form" ForeColor="Black" Width="40px"></asp:Label>
                            <asp:Label ID="Label24" runat="server" ForeColor="Black" Text="分至" CssClass="form" Width="30px"></asp:Label><asp:Label ID="Lab_nEDUSEDATE" runat="server" CssClass="form"
                                    ForeColor="Black" Width="100px"></asp:Label>
                            <asp:Label
                            ID="Label30" runat="server" CssClass="form" ForeColor="Black" Text="日" Width="20px"></asp:Label>
                            <asp:Label ID="Lab_nEDUSEHOUR" runat="server" CssClass="form" ForeColor="Black" Width="40px"></asp:Label>
                            <asp:Label ID="Label25" runat="server" ForeColor="Black" Text="時" CssClass="form" Width="20px"></asp:Label><asp:Label ID="Lab_nEDUSEMIN" runat="server" CssClass="form"
                                    ForeColor="Black" Width="40px"></asp:Label>
                            <asp:Label ID="Label26" runat="server" ForeColor="Black" Text="止" CssClass="form" Width="20px"></asp:Label></td> 
                    </tr> 
                </table> 
             
                <table style="width: 690px" rules="all">
                    <tr>
                        <td>
                            <asp:GridView ID="GridView1" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="form" DataSourceID="SqlDataSource1">
                                <Columns>
                                    <asp:BoundField DataField="GoLocal" HeaderText="起點" SortExpression="GoLocal" />
                                    <asp:BoundField DataField="EndLocal" HeaderText="終點" SortExpression="EndLocal" />
                                </Columns>
                            </asp:GridView>
                            <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                                SelectCommand="SELECT [EFORMSN], [GoLocal], [EndLocal] FROM [P_0305] WHERE ([EFORMSN] = @EFORMSN)">
                                <SelectParameters>
                                    <asp:QueryStringParameter Name="EFORMSN" QueryStringField="eformsn" Type="String" />
                                </SelectParameters>
                            </asp:SqlDataSource>
                        </td>
                    
                    </tr> 
                </table>                
                    
                <table style="width: 690px" rules="all">
                    <tr>
                        <td style="height: 23px;" align="center" colspan="4">
                            <asp:Label ID="Label10" runat="server" CssClass="form" ForeColor="Black" Text="行車安全保證責任檢查表"
                                Width="350px"></asp:Label></td>    
                    </tr> 
                    <tr>
                        <td style="width: 90px; height: 23px;">
                            <asp:Label ID="Label31" runat="server" CssClass="form" ForeColor="Black" Text="區分"></asp:Label></td>    
                        <td style="width: 270px; height: 23px;">
                            <asp:Label ID="Label27" runat="server" CssClass="form" ForeColor="Black" Text="保證責任要項"></asp:Label></td>     
                        <td style="width: 150px; height: 23px;">
                            <asp:Label ID="Label28" runat="server" CssClass="form" ForeColor="Black" Text="簽名" Width="60px"></asp:Label></td>    
                        <td style="width: 100px; height: 23px;">
                            <asp:Label ID="Label11" runat="server" CssClass="form" ForeColor="Black" Text="年月日" Width="70px"></asp:Label></td>     
                    </tr> 
                    <tr>
                        <td style="width: 90px;">
                            <asp:Label ID="Label43" runat="server" CssClass="form" ForeColor="Black" Text="用車單位主(管)官" Width="55px"></asp:Label></td>    
                        <td style="width: 270px;">
                            <asp:Label ID="Label32" runat="server" CssClass="form" ForeColor="Black" Text="一、申請車輛確實需要。"></asp:Label><br />
                            <asp:Label ID="Label33" runat="server" CssClass="form" ForeColor="Black" Text="二、指派車長，當面叮囑。"></asp:Label></td>     
                        <td style="width: 150px;">
                            <asp:Label ID="Lab_nUSEMASTER" runat="server" CssClass="form" ForeColor="Black" Width="100px"></asp:Label></td>    
                        <td style="width: 100px;">
                            &nbsp;&nbsp;</td>     
                    </tr> 
                    <tr>
                        <td style="width: 90px;">
                            <asp:Label ID="Label50" runat="server" CssClass="form" ForeColor="Black" Text="調度軍(士)官"
                                Width="85px"></asp:Label></td>    
                        <td style="width: 270px;">
                            <asp:Label ID="Label34" runat="server" CssClass="form" ForeColor="Black" Text="一、行車資料，證照要齊全。"></asp:Label><br />
                            <asp:Label ID="Label35" runat="server" CssClass="form" ForeColor="Black" Text="二、車容、車況均良好。"></asp:Label><br />
                            <asp:Label ID="Label36" runat="server" CssClass="form" ForeColor="Black" Text="三、駕駛身心健康，情緒正常。"></asp:Label></td>     
                        <td style="width: 150px;">
                            &nbsp; &nbsp;</td>    
                        <td style="width: 100px;">
                            &nbsp; &nbsp;</td>     
                    </tr> 
                    <tr>
                        <td style="width: 90px;">
                            <asp:Label ID="Label51" runat="server" CssClass="form" ForeColor="Black" Text="車輛駕駛"
                                Width="70px"></asp:Label></td>    
                        <td style="width: 270px;">
                            <asp:Label ID="Label37" runat="server" CssClass="form" ForeColor="Black" Text="一、完成車輛行駛前檢查保養。"></asp:Label><br />
                            <asp:Label ID="Label38" runat="server" CssClass="form" ForeColor="Black" Text="二、清點行車資料、證照。"></asp:Label><br />
                            <asp:Label ID="Label39" runat="server" CssClass="form" ForeColor="Black" Text="三、服從車長領導，遵守交通規格。"></asp:Label><br />
                            <asp:Label ID="Label40" runat="server" CssClass="form" ForeColor="Black" Text="四、指導人員，軍品乘載整齊。"></asp:Label></td>     
                        <td style="width: 150px;">
                            &nbsp; &nbsp;&nbsp;</td>    
                        <td style="width: 100px;">
                            &nbsp; &nbsp;&nbsp;</td>     
                    </tr> 
                    <tr>
                        <td style="width: 90px;">
                            <asp:Label ID="Label52" runat="server" CssClass="form" ForeColor="Black" Text="保養軍(士)官"
                                Width="85px"></asp:Label></td>    
                        <td style="width: 270px;">
                            <asp:Label ID="Label41" runat="server" CssClass="form" ForeColor="Black" Text="一、實施車輛出場前安全檢查。"></asp:Label><br />
                            <asp:Label ID="Label42" runat="server" CssClass="form" ForeColor="Black" Text="二、煞車、轉向、燈光、儀表、雨刷均良好。"></asp:Label></td>     
                        <td style="width: 150px;">
                            &nbsp;&nbsp;</td>    
                        <td style="width: 100px;">
                            &nbsp;&nbsp;</td>     
                    </tr> 
                    <tr>
                        <td style="width: 90px;">
                            <asp:Label ID="Label53" runat="server" CssClass="form" ForeColor="Black" Text="車長(副車長)"
                                Width="85px"></asp:Label></td>    
                        <td style="width: 270px;">
                            <asp:Label ID="Label44" runat="server" CssClass="form" ForeColor="Black" Text="一、監督駕駛遵守交通規格。"></asp:Label><br />
                            <asp:Label ID="Label46" runat="server" CssClass="form" ForeColor="Black" Text="二、維護人員、軍品乘載紀律與秩序。"></asp:Label><br />
                            <asp:Label ID="Label48" runat="server" CssClass="form" ForeColor="Black" Text="三、負責車輛行駛安全。"></asp:Label></td>     
                        <td style="width: 150px;">
                            <asp:Label ID="Lab_nCARMASTER" runat="server" CssClass="form" ForeColor="Black" Width="100px"></asp:Label></td>    
                        <td style="width: 100px;">
                            &nbsp; &nbsp;</td>     
                    </tr> 
                    <tr>
                        <td style="width: 90px;">
                            <asp:Label ID="Label54" runat="server" CssClass="form" ForeColor="Black" Text="營門衛兵"
                                Width="70px"></asp:Label></td>    
                        <td style="width: 270px;">
                            <asp:Label ID="Label45" runat="server" CssClass="form" ForeColor="Black" Text="一、檢查申請單確實批准。"></asp:Label><br />
                            <asp:Label ID="Label47" runat="server" CssClass="form" ForeColor="Black" Text="二、檢查行車安全保證責任表均已簽證。"></asp:Label><br />
                            <asp:Label ID="Label49" runat="server" CssClass="form" ForeColor="Black" Text="三、檢查車容及乘載情形均良好。"></asp:Label></td>     
                        <td style="width: 150px;">
                            &nbsp; &nbsp;</td>    
            
                        <td style="width: 100px;">
                            &nbsp; &nbsp;</td>     
                    </tr> 
                </table>    
                          
                <table style="width: 690px">
                    <tr>
                       <td align="center" colspan=4  style="width: 690px; height: 26px;">                       
                           <input id="Button1" onclick="javascript:window.print();" type="button" value="列印" /></td>
                    </tr> 
                </table>        
                <uc1:FlowRoute ID="FlowRoute1" runat="server" />
            </td>     
        </tr> 
        </table>    
        
    </div>
    </form>
    
</body>
</html>
