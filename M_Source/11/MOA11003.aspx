<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA11003.aspx.vb" Inherits="M_Source._11.M_Source_11_MOA11003" %>

<%@ Register src="../90/FlowRoute.ascx" tagname="FlowRoute" tagprefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">

        tr
        {
            padding: 5px;
        }
        td
        {
            padding: 5px;
        }
        .form
{
	font: 13px Verdana, Arial, Helvetica, sans-serif;
	color: #666666;
	text-decoration: none;
}

        .style1
        {
            width: 120px;
            height: 26px;
        }
        .style2
        {
            width: 550px;
            height: 26px;
        }
        th
        {
            padding: 5px;
        }
        </style>
</head>
<body lang="javascript" onload="return window_onload()">
    <form id="form1" runat="server">
    <div>
        <table width="750" border="1" cellspacing="0" cellpadding="5" bgcolor="#ffffff" bordercolor="#6699cc"
            bordercolorlight="#74a3d6" bordercolordark="#000000" style="left: 20px; top: 10px">
            <tr>
                <td valign="bottom" bgcolor="#6699cc" bordercolorlight="#66aaaa" bordercolordark="#ffffff">
                    <font color="white"><b>&nbsp;&nbsp;資訊設備修繕申請單 </b></font>
                </td>
            </tr>
            <tr>
                <td>
                    <fieldset id="tableB" style="width: 750px">
                        <table style="width: 750px; height: 57px">
                            <tr>
                                <td style="width: 240px">
                                    <asp:Label ID="Label1" runat="server" Text="填表人單位：" Width="80px" ForeColor="Black"
                                        CssClass="form"></asp:Label>
                                    <asp:Label ID="Lab_PWUNIT" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label>
                                    <td style="width: 220px">
                                        <asp:Label ID="Label2" runat="server" Text="姓名：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                                        <asp:Label ID="Lab_PWNAME" runat="server" ForeColor="Black" Width="106px" CssClass="form"></asp:Label>
                                    </td>
                                    <td style="width: 250px">
                                        <asp:Label ID="Label3" runat="server" Text="級職：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                                        <asp:Label ID="Lab_PWTITLE" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label>
                                    </td>
                            </tr>
                            <tr>
                                <td style="width: 240px">
                                    <asp:Label ID="Label4" runat="server" Text="申請人單位：" Width="80px" ForeColor="Black"
                                        CssClass="form"></asp:Label>
                                    <asp:Label ID="Lab_PAUNIT" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label>
                                </td>
                                <td style="width: 220px">
                                    <asp:Label ID="Label5" runat="server" Text="姓名：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>&nbsp;
                                        <asp:Label ID="Lab_PANAME" runat="server" ForeColor="Black" Width="106px" 
                                        CssClass="form"></asp:Label>
                                    <td style="width: 250px">
                                        <asp:Label ID="Label6" runat="server" Text="級職：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                                        <asp:Label ID="Lab_PATITLE" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label>
                                    </td>
                            </tr>
                        </table>
                    </fieldset>
                    <table border="0" style="width: 750px; height: 57px; color: Red">
                        <tr>
                            <td style="width: 120px; height: 23px;">
                                <asp:Label ID="Label7" runat="server" ForeColor="Black" Text="申請時間：" CssClass="form"></asp:Label>
                            </td>
                            <td style="width: 550px; height: 23px;">
                                <asp:Label ID="lblApplyTime" runat="server" ForeColor="Black" CssClass="form" Width="310px"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 120px">
                                &nbsp;<asp:Label ID="Label9" runat="server" ForeColor="Black" Text="電話：" CssClass="form"></asp:Label>
                            </td>
                            <td style="height: 26px; width: 550px;">
                                <asp:Label ID="lblPhone" runat="server" ForeColor="Black" CssClass="form" 
                                    Width="310px"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="style1">
                                <asp:Label ID="Label12" runat="server" ForeColor="Black" Text="請修地點：" CssClass="form"></asp:Label>
                            </td>
                            <td class="style2">
                                <asp:Label ID="lblBuilding" runat="server" ForeColor="Black" CssClass="form" 
                                    Width="310px"></asp:Label>
                                <asp:Label ID="lblLevel" runat="server" ForeColor="Black" CssClass="form" 
                                    Width="310px"></asp:Label>
                                <asp:Label ID="lblRoom" runat="server" ForeColor="Black" CssClass="form" 
                                    Width="310px"></asp:Label>
                                <asp:Label ID="Label29" runat="server" ForeColor="Black" Text="室"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td style="height: 26px; width: 120px;">
                                <asp:Label ID="Label8" runat="server" ForeColor="Black" Text="種類：" CssClass="form"></asp:Label>
                            </td>
                            <td style="height: 26px; width: 550px;">
                                <asp:Label ID="lblRepairMainKind" runat="server" ForeColor="Black" 
                                    CssClass="form" Width="310px"></asp:Label>
                                <asp:Label ID="lblProblemKind" runat="server" ForeColor="Black" CssClass="form" 
                                    Width="310px"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td style="height: 26px; width: 120px;">
                                <asp:Label ID="Label10" runat="server" ForeColor="Black" Text="請修事項：" CssClass="form"></asp:Label>
                            </td>
                            <td style="height: 26px; width: 550px;">
                                <asp:Label ID="lblDESCRIPTION" runat="server" ForeColor="Black" CssClass="form" 
                                    Width="310px"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td style="height: 26px; width: 120px;">
                                <asp:Label ID="Label30" runat="server" ForeColor="Black" Text="維修人員：" 
                                    CssClass="form"></asp:Label>
                            </td>
                            <td>
                                <asp:DropDownList ID="ddlRepairMan" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td style="height: 26px; width: 120px;">
                                <asp:Label ID="Label28" runat="server" ForeColor="Black" Text="批核意見：" CssClass="form"></asp:Label><br />
                            </td>
                            <td>
                                <asp:TextBox ID="txtcomment" runat="server" Height="59px" MaxLength="255" Rows="3"
                                    TextMode="MultiLine" Width="529px"></asp:TextBox>
                                <asp:Button ID="But_PHRASE" runat="server" Text="批核片語" />
                            </td>
                        </tr>
                    </table>
                    <table border="0" style="width: 750px; height: 57px">
                        <tr>
                            <td align="center">
                                <asp:Button ID="But_exe" runat="server" Text="送件" />
                                <asp:Button ID="But_prt" runat="server" Text="列印" Visible="False" />&nbsp;<asp:Button ID="backBtn"
                                    runat="server" Text="駁回" />&nbsp;<asp:Button ID="tranBtn" runat="server" 
                                    Text="呈轉" Visible="False" />
                            </td>
                        </tr>
                    </table>
                    <uc1:FlowRoute ID="FlowRoute1" runat="server" />
                </td>
            </tr>
        </table>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand=""></asp:SqlDataSource>
    </div>
    <div id="Div_grid10" runat="server" style="position: absolute; z-index: 3; background-color: white;
        width: 300pt; height: 80pt; left: 259px; top: 1097px; display: block;" 
        visible="false">
        <asp:GridView ID="GridView10" runat="server" CssClass="form" Width="100%" Height="50px"
            DataSourceID="SqlDataSource10" PageSize="5" AutoGenerateColumns="False" AllowPaging="True"
            BorderColor="Lime" BorderWidth="2px">
            <Columns>
                <asp:BoundField DataField="comment" HeaderText="批核片語">
                    <HeaderStyle HorizontalAlign="Center" Width="90%" BackColor="#80FF80" CssClass="form" />
                </asp:BoundField>
                <asp:BoundField DataField="Phrase_num" HeaderText="Phrase_num" Visible="False">
                    <ItemStyle Wrap="False" />
                </asp:BoundField>
                <asp:CommandField ShowSelectButton="True">
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:CommandField>
            </Columns>
            <RowStyle Height="10px" />
        </asp:GridView>
        &nbsp;
        <asp:Button ID="Btn_PHclose" runat="server" Text="關閉" Width="389px" />
        <asp:SqlDataSource ID="SqlDataSource10" runat="server" SelectCommand="SELECT Phrase_num, employee_id, comment FROM PHRASE WHERE [employee_id] = @employee_id ORDER BY Phrase_num"
            ConnectionString="<%$ ConnectionStrings:ConnectionString %>">
            <SelectParameters>
                <asp:SessionParameter Name="employee_id" SessionField="user_id" Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
    </div>
    </form>

</body>
</html>
