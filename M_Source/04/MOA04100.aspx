<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04100.aspx.vb" Inherits="Source_04_MOA04100" %>

<%@ Register Src="../90/FlowRoute.ascx" TagName="FlowRoute" TagPrefix="uc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>房屋水電修繕申請單</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
<%--    <script type="text/javascript" language="javascript">
        document.oncontextmenu = new Function("return false");
    </script>--%>
    <style type="text/css">
        .style1
        {
            height: 26px;
            width: 110px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table width="750" border="1" cellspacing="0" cellpadding="5" bgcolor="#ffffff" bordercolor="#6699cc"
            bordercolorlight="#74a3d6" bordercolordark="#000000" style="left: 20px; top: 10px">
            <tr>
                <td valign="bottom" bgcolor="#6699cc" bordercolorlight="#66aaaa" bordercolordark="#ffffff"
                    style="width: 782px">
                    <font color="white"><b>&nbsp;房屋水電修繕申請單</b></font>
                </td>
            </tr>
            <tr>
                <td style="width: 782px">
                    <fieldset id="tableB" style="width: 750px">
                        <table style="width: 750px; height: 57px">
                        <tr>
                        <td colspan ="2"><asp:Label ID="lbMsg" runat="server" ForeColor = "Red" Font-Size="Small" />
                        </td>
                        </tr>
                            <tr>
                                <td style="width: 240px">
                                    <asp:Label ID="Label1" runat="server" Text="填表人單位：" Width="80px" ForeColor="Black"
                                        CssClass="form"></asp:Label>
                                    <asp:Label ID="Lab_PWUNIT" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label>
                                </td>
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
                                    <asp:DropDownList ID="DrDown_PANAME" runat="server" Width="143px" AutoPostBack="True"
                                        DataSourceID="SqlDataSource1" DataTextField="emp_chinese_name" DataValueField="employee_id">
                                    </asp:DropDownList>
                                    <td style="width: 250px">
                                        <asp:Label ID="Label6" runat="server" Text="級職：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                                        <asp:Label ID="Lab_PATITLE" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label>
                                    </td>
                            </tr>
                        </table>
                    </fieldset>
                    <table border="0" style="width: 750px; height: 57px; color: Red">
                        <tr>
                            <td style="width: 100px; height: 23px;">
                                &nbsp;&nbsp;
                                <asp:Label ID="Label23" runat="server" ForeColor="Black" Text="報修單號：" CssClass="form"></asp:Label>
                            </td>
                            <td style="height: 23px;" colspan="3">
                                <asp:Label ID="lbleformsn" runat="server" ForeColor="Black" CssClass="form" Width="310px"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 100px; height: 23px;">
                                *
                                <asp:Label ID="Label7" runat="server" ForeColor="Black" Text="申請時間：" CssClass="form"></asp:Label>
                            </td>
                            <td style="height: 23px;" colspan="3">
                                <asp:Label ID="AppDate" runat="server" ForeColor="Black" CssClass="form" Width="310px"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 100px">
                                *
                                <asp:Label ID="Label9" runat="server" ForeColor="Black" Text="聯絡電話：" CssClass="form"></asp:Label>
                            </td>
                            <td style="height: 26px;" colspan="3">
                                <asp:TextBox ID="Txt_nPHONE" runat="server" MaxLength="25" Width="160px"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td style="height: 26px; width: 100px;">
                                *
                                <asp:Label ID="Label12" runat="server" ForeColor="Black" Text="請修地點：" CssClass="form"></asp:Label>
                            </td>
                            <td style="height: 26px;" colspan="3">
                                <asp:DropDownList ID="DDLBD" runat="server" AutoPostBack="True" DataSourceID="SqlDataSourceEle1"
                                    DataTextField="bd_name" DataValueField="bd_code">
                                </asp:DropDownList>
                                <asp:DropDownList ID="DDLFL" runat="server" AutoPostBack="True" DataSourceID="SqlDataSourceEle2"
                                    DataTextField="fl_name" DataValueField="fl_code">
                                </asp:DropDownList>
                                <asp:DropDownList ID="DDLRNUM" runat="server" AutoPostBack="True" DataSourceID=""
                                    DataTextField="ShowName" DataValueField="rnum_code">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td style="height: 26px; width: 100px;">
                                *
                                <asp:Label ID="Label10" runat="server" ForeColor="Black" Text="請修事項：" CssClass="form"></asp:Label>
                            </td>
                            <td colspan="3" style="height: 26px">
                                <asp:TextBox ID="Txt_nFIXITEM" runat="server" MaxLength="2000" Width="535px" Height="80px"
                                    TextMode="MultiLine"></asp:TextBox>
                            </td>
                        </tr>
                        <% If read_only <> "" And stepChk >= 2 Then%>
                        <asp:Panel ID="Pl_2w" runat="server" Visible = "false">
                        <tr>
                            <td style="height: 26px; width: 100px;">
                                *
                                <asp:Label ID="Label21" runat="server" CssClass="form" ForeColor="Black" Text="現勘人員："></asp:Label>
                            </td>
                            <td style="height: 26px; width: 230px;">
                                <asp:DropDownList ID="DDViewPer" runat="server" />
                              
                                
                            </td>
                            <td class="style1">
                                *
                                <asp:Label ID="Label22" runat="server" CssClass="form" ForeColor="Black" Text="現勘日期："></asp:Label>
                            </td>
                            <td style="width: 220px; height: 26px">
                                <asp:TextBox ID="Txt_nViewDate" runat="server" EnableTheming="False" onkeydown="return false"
                                    Width="89px"></asp:TextBox>
                                <asp:ImageButton ID="ImgBtn_View" runat="server" ImageUrl="~/Image/calendar.gif" />
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 100px">
                                *
                                <asp:Label ID="Label17" runat="server" CssClass="form" ForeColor="Black" Text="原因分析："></asp:Label>
                            </td>
                            <td colspan="3">
                                <asp:TextBox ID="Txt_nCause" runat="server" MaxLength="255" Width="535px"></asp:TextBox>&nbsp;
                            </td>
                        </tr>
                        <tr>
                            <td style="height: 26px; width: 100px;">
                                *
                                <asp:Label ID="Label13" runat="server" ForeColor="Black" Text="派員人數：" CssClass="form"></asp:Label>
                            </td>
                            <td style="height: 26px; width: 230px;">
                                <asp:TextBox ID="Txt_nPacthCount" runat="server" MaxLength="4" Width="85px"></asp:TextBox>
                            </td>
                            <td class="style1">
                                *
                                <asp:Label ID="Label14" runat="server" ForeColor="Black" Text="派工人員：" CssClass="form"></asp:Label>
                            </td>
                            <td style="width: 220px; height: 26px">
                                <asp:TextBox ID="Txt_nPacthPer" runat="server" MaxLength="100" Width="150px"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td style="height: 26px; width: 100px;">
                                *
                                <asp:Label ID="Label16" runat="server" ForeColor="Black" Text="設施(備)編號：" CssClass="form"></asp:Label>
                            </td>
                            <td style="height: 26px; width: 230px;">
                                <asp:TextBox ID="Txt_nFacilityNo" runat="server" MaxLength="50" Width="140px" ReadOnly="True"></asp:TextBox>
                                <asp:Button ID="Button2" runat="server" Text="編號選擇" />
                            </td>
                            <td class="style1">
                                *
                                <asp:Label ID="Label11" runat="server" CssClass="form" ForeColor="Black" Text="預計開工日期："></asp:Label>
                            </td>
                            <td style="width: 220px; height: 26px">
                                <asp:TextBox ID="Txt_nStartDATE" runat="server" EnableTheming="False" onkeydown="return false"
                                    Width="89px"></asp:TextBox>
                                <asp:ImageButton ID="ImgBtn_SDate" runat="server" ImageUrl="~/Image/calendar.gif" />
                            </td>
                        </tr>
                        <tr>
                            <td style="height: 26px; width: 100px;">
                                *
                                <asp:Label ID="Label18" runat="server" CssClass="form" ForeColor="Black" Text="故障原因："></asp:Label>
                            </td>
                            <td style="height: 26px; width: 230px;">
                                <asp:DropDownList ID="DDL_nErrCause" runat="server">
                                    <asp:ListItem Value="1">人為因素</asp:ListItem>
                                    <asp:ListItem Value="2">自然因素</asp:ListItem>
                                    <asp:ListItem Value="3">維護查報</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                            <td class="style1">
                                *
                                <asp:Label ID="Label15" runat="server" CssClass="form" ForeColor="Black" Text="承辦類別："></asp:Label>
                            </td>
                            <td style="width: 220px; height: 26px">
                                <asp:RadioButtonList ID="rdo_nExternal" runat="server" ForeColor="Black" RepeatDirection="Horizontal"
                                    Width="150px" CssClass="form">
                                    <asp:ListItem Selected="True">內派</asp:ListItem>
                                    <asp:ListItem>外包</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                       </asp:Panel>
                        <% End If%>
                        <% If read_only <> "" And stepChk = 3 Then%>
                        <tr>
                            <td style="height: 26px;" colspan="4">
                                <asp:Button ID="btn_get" runat="server" Text="領料申請" />
                                <br />
                                <div style="padding-left: 10px; padding-top: 2px">
                                    <asp:GridView ID="GridView11" runat="server" AutoGenerateColumns="False" CssClass="form"
                                        DataKeyNames="it_code" DataSourceID="SqlDataSource5">
                                        <Columns>
                                            <asp:BoundField DataField="it_code" HeaderText="用料編號" ReadOnly="True" SortExpression="it_code" />
                                            <asp:BoundField DataField="it_name" HeaderText="用料品項" SortExpression="it_name" />
                                            <asp:TemplateField HeaderText="用料數量">
                                                <ItemTemplate>
                                                    <asp:Label ID="lblIt_App" runat="server" />
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                            <asp:BoundField DataField="it_unit" HeaderText="用料單位" SortExpression="it_unit" />
                                            <asp:TemplateField HeaderText="已領取量">
                                                <ItemTemplate>
                                                    <asp:Label ID="lblIt_Use" runat="server" />
                                                    <asp:HiddenField ID="Hidit_code" runat="server" Value='<%# Bind("it_code") %>' />
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                        </Columns>
                                    </asp:GridView>
                                </div>
                                <asp:SqlDataSource ID="SqlDataSource5" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                                    SelectCommand=""></asp:SqlDataSource>
                                <br />
                                <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>" />
                            </td>
                        </tr>
                        <asp:Panel ID="Pl_3w" runat="server" Visible = "false">
                        <tr>
                            <td style="height: 26px; width: 100px;">
                                *
                                <asp:Label ID="Label19" runat="server" CssClass="form" ForeColor="Black" Text="目前現況："></asp:Label>
                            </td>
                            <td style="height: 26px; width: 230px;">
                                <asp:TextBox ID="Txt_nNowStatus" runat="server" MaxLength="100" Width="150px"></asp:TextBox>
                            </td>
                            <td class="style1">
                                *
                                <asp:Label ID="Label20" runat="server" CssClass="form" ForeColor="Black" Text="完工日期："></asp:Label>
                            </td>
                            <td style="width: 220px; height: 26px">
                                <asp:TextBox ID="Txt_nFINALDATE" runat="server" EnableTheming="False" onkeydown="return false"
                                    Width="89px"></asp:TextBox>
                                <asp:ImageButton ID="ImgBtn_EDate" runat="server" ImageUrl="~/Image/calendar.gif" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                *
                                <asp:Label ID="Label8" runat="server" CssClass="form" ForeColor="Black" Text="處理結果："></asp:Label>
                            </td>
                            <td colspan="3">
                                <asp:TextBox ID="Txt_nResult" runat="server" TextMode="MultiLine" Rows="3" Width="525px"></asp:TextBox>
                            </td>
                        </tr>
                        </asp:Panel>
                        <% End If%>
                        <% If confirmFlag = "1" Then%>
                        <tr>
                            <td style="height: 26px; width: 100px;">
                                &nbsp;<asp:Label ID="Label28" runat="server" ForeColor="Black" Text="批核意見：" CssClass="form"></asp:Label>
                            </td>
                            <td colspan="3">
                                <asp:TextBox ID="txtcomment" runat="server" Height="59px" MaxLength="255" Rows="3"
                                    TextMode="MultiLine" Width="536px"></asp:TextBox>
                                <asp:Button ID="But_PHRASE" runat="server" Text="批核片語" />
                                <div id="Div_grid10" runat="server" style="position: relative; z-index: 3;
                                        background-color: white;" visible="false">
                                    <div style="position: absolute;border: 2 solid;
                                        background-color: white; width: 300pt; left:0px;top:10px; display: block;">
                                        <asp:GridView ID="GridView10" runat="server" CssClass="form" Width="100%" DataSourceID="SqlDataSource10"
                                            PageSize="5" AutoGenerateColumns="False" AllowPaging="True" BorderColor="#698BCC"
                                            BorderWidth="2px">
                                            <Columns>
                                                <asp:BoundField DataField="comment" HeaderText="批核片語">
                                                    <HeaderStyle HorizontalAlign="Center" Width="90%" BackColor="#B5C6E6" CssClass="form" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Phrase_num" HeaderText="Phrase_num" Visible="False">
                                                    <ItemStyle Wrap="False" />
                                                </asp:BoundField>
                                                <asp:CommandField ShowSelectButton="True">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%" BackColor="#B5C6E6" />
                                                </asp:CommandField>
                                            </Columns>
                                            <RowStyle Height="10px" />
                                        </asp:GridView>
                                        <asp:Button ID="Btn_PHclose" runat="server" Text="關閉" />
                                        <asp:SqlDataSource ID="SqlDataSource10" runat="server" SelectCommand="SELECT Phrase_num, employee_id, comment FROM PHRASE WHERE [employee_id] = @employee_id ORDER BY Phrase_num"
                                            ConnectionString="<%$ ConnectionStrings:ConnectionString %>">
                                            <SelectParameters>
                                                <asp:SessionParameter Name="employee_id" SessionField="user_id" Type="String" />
                                            </SelectParameters>
                                        </asp:SqlDataSource>
                                    </div>
                                </div>
                            </td>
                        </tr>
                        <% End If%>
                    </table>
                    <table border="0" style="width: 750px; height: 57px">
                        <tr>
                            <td align="center">
                                <asp:Button ID="prtP1btn" runat="server" Text="列印報修單" Visible="false" />
                                <asp:Button ID="prtP2btn" runat="server" Text="列印派工單" Visible="false" />
                                <asp:Button ID="SaveBtn" runat="server" Text="暫存" Visible="false" />
                                <asp:Button ID="send" runat="server" Text="送件" />
                                <asp:Button ID="EndBtn" runat="server" Text="結束" Visible="false" />
                                <asp:Button ID="backBtn" runat="server" Text="駁回" />&nbsp;
                            </td>
                        </tr>
                    </table>
                    <uc1:FlowRoute ID="FlowRoute1" runat="server" />
                </td>
            </tr>
        </table>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT * FROM [V_EmpInfo] WHERE ([ORG_UID] = @ORG_UID) ORDER BY [emp_chinese_name]">
            <SelectParameters>
                <asp:SessionParameter Name="ORG_UID" SessionField="org_uid" Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand=""></asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSourceEle1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT * FROM [P_0404] ORDER BY [no]"></asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSourceEle2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT * FROM [P_0406] ORDER BY [fl_no]"></asp:SqlDataSource>
      
        <asp:SqlDataSource ID="SqlDataSourceEle4" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT * FROM [P_0413] ORDER BY [no]"></asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSourceEle5" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT * FROM [P_0403] ORDER BY [no]"></asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSourceEle6" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT p_0405.it_code,it_name,p_0405.it_code+'-'+it_name as it_ShowName FROM [P_0405],[P_0407] WHERE p_0405.it_code = p_0407.it_code and (([bd_code] = @bd_code) AND ([bg_code] = @bg_code) AND ([fl_code] = @fl_code) AND ([wa_code] = @wa_code) AND ([rnum_code] = @rnum_code)) UNION select '','',''">
            <SelectParameters>
                <asp:ControlParameter ControlID="DDLBD" Name="bd_code" PropertyName="SelectedValue"
                    Type="String" />
                <asp:ControlParameter ControlID="DDLBG" Name="bg_code" PropertyName="SelectedValue"
                    Type="String" />
                <asp:ControlParameter ControlID="DDLFL" Name="fl_code" PropertyName="SelectedValue"
                    Type="String" />
                <asp:ControlParameter ControlID="DDLWA" Name="wa_code" PropertyName="SelectedValue"
                    Type="String" />
                <asp:ControlParameter ControlID="DDLRNUM" Name="rnum_code" PropertyName="SelectedValue"
                    Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSourceEle7" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT '' AS [element_code] FROM [P_0405]&#13;&#10;UNION&#13;&#10;SELECT DISTINCT [element_code] FROM [P_0405] WHERE (([bd_code] = @bd_code) AND ([bg_code] = @bg_code) AND ([fl_code] = @fl_code) AND ([wa_code] = @wa_code) AND ([rnum_code] = @rnum_code)) AND ([it_code] = @it_code)">
            <SelectParameters>
                <asp:ControlParameter ControlID="DDLBD" Name="bd_code" PropertyName="SelectedValue"
                    Type="String" />
                <asp:ControlParameter ControlID="DDLBG" Name="bg_code" PropertyName="SelectedValue"
                    Type="String" />
                <asp:ControlParameter ControlID="DDLFL" Name="fl_code" PropertyName="SelectedValue"
                    Type="String" />
                <asp:ControlParameter ControlID="DDLWA" Name="wa_code" PropertyName="SelectedValue"
                    Type="String" />
                <asp:ControlParameter ControlID="DDLRNUM" Name="rnum_code" PropertyName="SelectedValue"
                    Type="String" />
                <asp:ControlParameter ControlID="DDLIT" Name="it_code" PropertyName="SelectedValue"
                    Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlDataSourceEle8" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT * FROM [P_0416]"></asp:SqlDataSource>
    </div>
    <div id="Div_nFacilityNo" runat="server" style="border-right: lightslategray 2px solid;
        border-top: lightslategray 2px solid; display: block; z-index: 3; left: 290px;
        border-left: lightslategray 2px solid; border-bottom: lightslategray 2px solid;
        position: absolute; top: 420px; background-color: #F2F8FD" visible="false">
        <div style="margin: 10px">
            <span>請依序選擇：</span><br />
            <asp:DropDownList ID="DDLWA" runat="server" AutoPostBack="True" DataSourceID="SqlDataSourceEle4"
                DataTextField="wa_name" DataValueField="wa_code">
            </asp:DropDownList>
            <br />
            <asp:DropDownList ID="DDLBG" runat="server" AutoPostBack="True" DataSourceID="SqlDataSourceEle5"
                DataTextField="bg_name" DataValueField="bg_code">
            </asp:DropDownList>
            <br />
            <asp:DropDownList ID="DDLIT" runat="server" AutoPostBack="True" DataSourceID="SqlDataSourceEle6"
                DataTextField="it_ShowName" DataValueField="it_code">
            </asp:DropDownList>
            <br />
            <asp:DropDownList ID="DDLelement" runat="server" AutoPostBack="True" DataSourceID="SqlDataSourceEle7"
                DataTextField="element_code" DataValueField="element_code">
            </asp:DropDownList>
            <br />
            <asp:Button ID="btn_Close" runat="server" Text="關閉" Width="75px" />
        </div>
    </div>
    <div id="DivView_grid" runat="server" style="border-right: lightslategray 2px solid;
        border-top: lightslategray 2px solid; display: block; z-index: 3; left: 103px;
        border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
        position: absolute; top: 823px; height: 150pt; background-color: white" visible="false">
        <asp:Calendar ID="Calendar1" runat="server" BackColor="White" BorderColor="#3366CC"
            BorderWidth="1px" Caption="" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana"
            Font-Size="8pt" ForeColor="#003399" Height="200px" ShowGridLines="True" Width="220px">
            <SelectedDayStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
            <TodayDayStyle BackColor="#99CCCC" ForeColor="White" />
            <SelectorStyle BackColor="#99CCCC" ForeColor="#336666" />
            <WeekendDayStyle BackColor="#CCCCFF" />
            <OtherMonthDayStyle ForeColor="#999999" />
            <NextPrevStyle Font-Size="8pt" ForeColor="#CCCCFF" />
            <DayHeaderStyle BackColor="#99CCCC" ForeColor="#336666" Height="1px" />
            <TitleStyle BackColor="#003399" BorderColor="#3366CC" BorderWidth="1px" Font-Bold="True"
                Font-Size="10pt" ForeColor="#CCCCFF" Height="25px" />
        </asp:Calendar>
        <asp:Button ID="btnClose1" runat="server" Text="關閉" Width="220px" />
    </div>
    <div id="DivSDate_grid" runat="server" style="border-right: lightslategray 2px solid;
        border-top: lightslategray 2px solid; display: block; z-index: 3; left: 103px;
        border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
        position: absolute; top: 823px; height: 150pt; background-color: white" visible="false">
        <asp:Calendar ID="Calendar2" runat="server" BackColor="White" BorderColor="#3366CC"
            BorderWidth="1px" Caption="" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana"
            Font-Size="8pt" ForeColor="#003399" Height="200px" ShowGridLines="True" Width="220px">
            <SelectedDayStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
            <TodayDayStyle BackColor="#99CCCC" ForeColor="White" />
            <SelectorStyle BackColor="#99CCCC" ForeColor="#336666" />
            <WeekendDayStyle BackColor="#CCCCFF" />
            <OtherMonthDayStyle ForeColor="#999999" />
            <NextPrevStyle Font-Size="8pt" ForeColor="#CCCCFF" />
            <DayHeaderStyle BackColor="#99CCCC" ForeColor="#336666" Height="1px" />
            <TitleStyle BackColor="#003399" BorderColor="#3366CC" BorderWidth="1px" Font-Bold="True"
                Font-Size="10pt" ForeColor="#CCCCFF" Height="25px" />
        </asp:Calendar>
        <asp:Button ID="btnClose2" runat="server" Text="關閉" Width="220px" />
    </div>
    <div id="DivEDate_grid" runat="server" style="border-right: lightslategray 2px solid;
        border-top: lightslategray 2px solid; display: block; z-index: 3; left: 103px;
        border-left: lightslategray 2px solid; width: 165pt; border-bottom: lightslategray 2px solid;
        position: absolute; top: 823px; height: 150pt; background-color: white" visible="false">
        <asp:Calendar ID="Calendar3" runat="server" BackColor="White" BorderColor="#3366CC"
            BorderWidth="1px" Caption="" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana"
            Font-Size="8pt" ForeColor="#003399" Height="200px" ShowGridLines="True" Width="220px">
            <SelectedDayStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
            <TodayDayStyle BackColor="#99CCCC" ForeColor="White" />
            <SelectorStyle BackColor="#99CCCC" ForeColor="#336666" />
            <WeekendDayStyle BackColor="#CCCCFF" />
            <OtherMonthDayStyle ForeColor="#999999" />
            <NextPrevStyle Font-Size="8pt" ForeColor="#CCCCFF" />
            <DayHeaderStyle BackColor="#99CCCC" ForeColor="#336666" Height="1px" />
            <TitleStyle BackColor="#003399" BorderColor="#3366CC" BorderWidth="1px" Font-Bold="True"
                Font-Size="10pt" ForeColor="#CCCCFF" Height="25px" />
        </asp:Calendar>
        <asp:Button ID="btnClose3" runat="server" Text="關閉" Width="220px" />
    </div>
    </form>
</body>
</html>
