<%@ Page Language="VB" EnableEventValidation = "false" AutoEventWireup="false" CodeFile="MOA04104.aspx.vb" Inherits="M_Source_04_MOA04104" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>房屋水電修繕單項明細資料</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
      <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr><td align="center" style="height: 24px">
                    <asp:Label ID="lbTypeName" runat="server" CssClass="toptitle" Width="100%" />
            </td></tr>
        </table>
    <div>
        
        <center>
            <asp:GridView ID="gvData" runat="server" AllowPaging="True" 
            AllowSorting="True" AutoGenerateColumns="False" visible = "False"
        CellPadding="3" EmptyDataText="查無任何資料" GridLines="Horizontal" 
            Width="100%" CssClass ="form" BackColor="White" 
            BorderColor="#E7E7FF" BorderStyle="None" BorderWidth="1px" 
          EnableModelValidation="True">
            <Columns>
               <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:LinkButton ID="LBItCodeDetail" runat="server" CausesValidation="False" CommandArgument=<%#Eval("P_Num") %> CommandName="GetDetail" text="明細" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="5%" />
                </asp:TemplateField> 
               <asp:BoundField  DataField="EFORMSN" HeaderText="報修單號">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="10%" />
                </asp:BoundField>
                <asp:BoundField  DataField="PAUNIT" HeaderText="申請單位">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="7%" />
                </asp:BoundField>
                <asp:BoundField  DataField="PANAME" HeaderText="申請人姓名">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="5%" />
                </asp:BoundField>
                <asp:BoundField  DataField="nAPPTIME" HeaderText="申請時間">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="8%" />
                </asp:BoundField>
                     <asp:BoundField  DataField="nFIXITEM" HeaderText="請修事項">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="12%" />
                </asp:BoundField>
                 <asp:BoundField  DataField="location" HeaderText="請修地點">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="10%" />
                </asp:BoundField>
                <asp:BoundField  DataField="nCause" HeaderText="原因分析">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="15%" />
                </asp:BoundField>
                <asp:BoundField  DataField="nFinalDate" HeaderText="完工時間">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="8%" />
                </asp:BoundField>
                 <asp:BoundField  DataField="nResult" HeaderText="處理結果">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="20%" />
                </asp:BoundField>
            </Columns>
                <PagerStyle HorizontalAlign="Center" />
        </asp:GridView>
            <asp:Label ID="lbEmptyData" runat="server" Font-Bold="True" ForeColor="Red" />  
            <asp:ImageButton ID="Img_Export" runat="server" ImageUrl="~/Image/ExportFile.gif" Enabled="False" />
            <asp:ImageButton ID="ibtnPrevious" runat="server" ImageUrl="~/Image/backtop.gif" />
        </center>
    </div>
    </form>
</body>
</html>
