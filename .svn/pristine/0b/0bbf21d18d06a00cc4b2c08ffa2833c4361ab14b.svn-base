<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04107.aspx.vb" Inherits="M_Source_04_MOA04107" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>派工詳細資料</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">    
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center" style="height: 24px">
                    <asp:Label ID="Label12" runat="server" CssClass="toptitle" Text="派工詳細資料" Width="100%"></asp:Label>
            </td></tr>
        </table>
        <asp:DetailsView ID="DetailsView1" runat="server" AutoGenerateRows="False" Height="50px" Width="100%" CssClass="form">
            <Fields>
             <asp:BoundField DataField="EFORMSN" HeaderText="表單序號： 　　&nbsp;">
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                 <asp:BoundField DataField="PWUNIT" HeaderText="填表人單位： 　　&nbsp;" >
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                 </asp:BoundField>
                 <asp:BoundField DataField="PWNAME" HeaderText="填表人姓名： 　　&nbsp;" >
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                 </asp:BoundField>
                  <asp:BoundField DataField="PWTITLE" HeaderText="填表人職級： 　　&nbsp;" >
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                 </asp:BoundField>
                 <asp:BoundField DataField="nAPPTIME" HeaderText="申請日期： 　　&nbsp;" >
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                 </asp:BoundField>
                 <asp:BoundField DataField="PAUNIT" HeaderText="申請人單位： 　　&nbsp;" >
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                 <asp:BoundField DataField="PANAME" HeaderText="申請人姓名： 　　&nbsp;" >
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="PATITLE" HeaderText="申請人職級： 　　&nbsp;" >
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                 <asp:BoundField DataField="nPHONE" HeaderText="聯絡電話： 　　&nbsp;" >
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                 <asp:BoundField DataField="nPLACEName" HeaderText="請修地點： 　　&nbsp;" >
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="nFIXITEM" HeaderText="請修事項： 　　&nbsp;" >
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                 </asp:BoundField>
                  <asp:BoundField DataField="nPacthCount" HeaderText="派員人數： 　　&nbsp;" >
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                 </asp:BoundField>
                 <asp:BoundField DataField="nPacthPer" HeaderText="派工人員： 　　&nbsp;" >
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                 </asp:BoundField>
                  <asp:BoundField DataField="House_Name" HeaderText="現勘人員： 　　&nbsp;" >
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="現勘日期： 　　&nbsp;" >
                    <ItemTemplate>
                      <%# CheckDate(Eval("nViewDate"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:TemplateField>
                 <asp:TemplateField HeaderText="故障原因： 　　&nbsp;" >
                    <ItemTemplate>
                      <%# ErrorWord(Eval("nErrCause"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:TemplateField>
                 <asp:BoundField DataField="nNowStatus" HeaderText="目前現況： 　　&nbsp;" >
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="開工日期： 　　&nbsp;" >
                    <ItemTemplate>
                      <%# CheckDate(Eval("nStartDATE"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="完工日期： 　　&nbsp;" >
                    <ItemTemplate>
                      <%# CheckDate(Eval("nFinalDate"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:TemplateField>
                 <asp:BoundField DataField="nFacilityNo" HeaderText="設施編號： 　　&nbsp;" >
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                 </asp:BoundField>
                 <asp:BoundField DataField="nExternal" HeaderText="承辦類別： 　　&nbsp;" >
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                 </asp:BoundField>
              　　<asp:TemplateField HeaderText="流程狀態： 　　&nbsp;" >
                    <ItemTemplate>
                      <%# StatusName(Eval("FlowStatus"))%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:TemplateField>
                <asp:BoundField DataField="nResult" HeaderText="處理結果： 　　&nbsp;" >
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
            </Fields>
            <FieldHeaderStyle Width="30%" />
        </asp:DetailsView>
             
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr>
            <td align="center">
                <a href ="#" onclick ="javascript:history.back();"><asp:Image ID="Image1" runat="server" ImageUrl="~/Image/backtop.gif" border=0 /></a>
            </td>
            </tr>
        </table>
    <asp:Panel ID="Pnl_ItCodeDetail" runat="server" Visible = "false" >
         <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center" style="height: 24px">
               <asp:Label ID="Label1" runat="server" CssClass="toptitle" Text="領料詳細資料" Width="100%"></asp:Label>
            </td>
            </tr>
            <tr>
            <td>
             <asp:GridView ID="gv_itData" runat="server" AutoGenerateColumns="False" CssClass="form"
                     width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;" 
                    BorderStyle="Dotted" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="it_code" HeaderText="用料編號" ReadOnly="True" SortExpression="it_code" />
                <asp:BoundField DataField="it_name" HeaderText="用料品項" SortExpression="it_name" />
                <asp:BoundField DataField="cnt" HeaderText="用料/已領數量" SortExpression="cnt" />
                <asp:BoundField DataField="it_unit" HeaderText="用料單位" SortExpression="it_unit" />
            </Columns>
                 <RowStyle VerticalAlign="Middle" Font-Size="Smaller" ForeColor="#666666" 
                     HorizontalAlign="Center" />
        </asp:GridView></td>
            </tr>
        </table>
          </asp:Panel>
    </form>
</body>
</html>