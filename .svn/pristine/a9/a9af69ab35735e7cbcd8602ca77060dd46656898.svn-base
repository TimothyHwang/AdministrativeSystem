<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04005.aspx.vb" Inherits="Source_04_MOA04005" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
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
        <asp:DetailsView ID="DetailsView1" runat="server" AutoGenerateRows="False"
            DataSourceID="SqlDataSource1" Height="50px" Width="100%">
            <Fields>
                <asp:BoundField DataField="nFIXDATE" HeaderText="維修日期" SortExpression="nFIXDATE" >
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nFinalDate" HeaderText="確認完工" SortExpression="nFinalDate" >
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nHandleDate" HeaderText="交辦日期" SortExpression="nHandleDate" >
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nCheckDate" HeaderText="勘查日期" SortExpression="nCheckDate" >
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nDoleDate" HeaderText="發料日期" SortExpression="nDoleDate" >
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="nFinishDate" HeaderText="完工日期" SortExpression="nFinishDate" >
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
            </Fields>
            <FieldHeaderStyle Width="30%" />
        </asp:DetailsView>
             
        <table  border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr>
            <td align="center">
                <asp:ImageButton ID="BackBtn" runat="server" ImageUrl="~/Image/backtop.gif" AlternateText="回上頁" />
            </td>
            </tr>
        </table>
        
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT * FROM P_04 LEFT OUTER JOIN P_0401 ON P_04.EFORMSN = P_0401.nEFORMSN WHERE P_04.P_NUM =@P_NUM">
            <SelectParameters>
                <asp:QueryStringParameter Name="p_num" QueryStringField="P_Num" />
            </SelectParameters>
        </asp:SqlDataSource>
    
    </form>
</body>
</html>
