<%@ Page Language="VB" EnableEventValidation = "false" AutoEventWireup="true" CodeFile="MOA08009.aspx.vb" Inherits="M_Source_08_MOA08009" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>影印記錄統計-單項明細資料</title>
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
            <asp:GridView ID="gvPrinterData" runat="server" AllowPaging="True" 
            AllowSorting="True" AutoGenerateColumns="False" visible = "False"
        CellPadding="3" EmptyDataText="查無任何資料" GridLines="Horizontal" 
            Width="100%" CssClass ="form" BackColor="White" 
            BorderColor="#E7E7FF" BorderStyle="None" BorderWidth="1px" 
          EnableModelValidation="True">
            <Columns>
               <asp:TemplateField ShowHeader="False"> 
                    <ItemTemplate>
                        <asp:HiddenField ID="HFPrintLogStatus" runat="server" Value = '<%# Eval("status")%>'  />
                        <asp:ImageButton ID="ImgPrinterDetail" ImageUrl ="~/Image/List.gif" runat="server" CausesValidation="False" CommandArgument=<%#Eval("Log_Guid") %> CommandName="GetDetail" Visible ="false" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="5%" />
                </asp:TemplateField> 
                <asp:BoundField  DataField="Log_Guid" HeaderText="流水號" Visible="false">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="10%" />
                </asp:BoundField>
                <asp:BoundField  DataField="ORG_NAME" HeaderText="部門名稱">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="15%" />
                </asp:BoundField>
                <asp:BoundField  DataField="emp_chinese_name" HeaderText="送印人姓名">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="10%" />
                </asp:BoundField>
                <asp:BoundField  DataField="Printer_Name" HeaderText="影印機名稱">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="8%" />
                </asp:BoundField>
                     <asp:BoundField  DataField="PAIDNO" HeaderText="人員帳號">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="12%" />
                </asp:BoundField>
                 <asp:BoundField  DataField="PrintLogDate" HeaderText="列印日期時間">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="15%" />
                </asp:BoundField>
                <asp:BoundField  DataField="security_statusName" HeaderText="保密區分">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="8%" />
                </asp:BoundField>
                <asp:BoundField  DataField="UPdate_Date" HeaderText="登記/修改時間">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="15%" />
                </asp:BoundField>
            </Columns>
            <PagerStyle HorizontalAlign="Center" />
        </asp:GridView>

            <asp:GridView ID="gvPrintHistory" runat="server" AllowPaging="True" 
            AllowSorting="True" AutoGenerateColumns="False" visible = "False"
        CellPadding="3" EmptyDataText="查無任何資料" GridLines="Horizontal" 
            Width="100%" CssClass ="form" BackColor="White" 
            BorderColor="#E7E7FF" BorderStyle="None" BorderWidth="1px" 
          EnableModelValidation="True" PageSize="15">
            <Columns>
                <asp:BoundField  DataField="History_ID" HeaderText="流水號" Visible="false">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="10%" />
                </asp:BoundField>
                 <asp:TemplateField HeaderText ="影印記錄編號" > 
                    <ItemTemplate>
                        <asp:Label ID="lbPrintLog_ID" runat="server" Text= '<%# Eval("PrintLog_ID")%>' />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="7%" />
                </asp:TemplateField> 
                <asp:BoundField  DataField="employee_id" HeaderText="人員帳號">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="12%" />
                </asp:BoundField>
                <asp:BoundField  DataField="ORG_NAME" HeaderText="部門名稱">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="15%" />
                </asp:BoundField>
                <asp:BoundField  DataField="emp_chinese_name" HeaderText="歷程人員">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="10%" />
                </asp:BoundField>
                 <asp:BoundField  DataField="movementName" HeaderText="操作動作">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="5%" />
                </asp:BoundField>
                <asp:BoundField  DataField="History_Date" HeaderText="歷程動作時間">
                    <ItemStyle HorizontalAlign="Center" VerticalAlign ="Top" Width="15%" />
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
