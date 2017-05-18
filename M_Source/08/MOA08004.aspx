<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA08004.aspx.vb" Inherits="M_Source_08_MOA08004" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>影印記錄詳細記錄</title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
</head>
<body>
    <form id="form1" runat="server">
      <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center">
                    <asp:Label ID="lblTitle" runat="server" CssClass="toptitle" Width="100%" Text ="單一影印記錄詳細記錄"></asp:Label>
                </td>
            </tr>
        </table>
        <asp:DetailsView ID="DV_PrintLog" runat="server" AutoGenerateRows="False" DefaultMode="ReadOnly" 
         DataKeyNames="Log_Guid" Width="100%" CssClass="form">
            <Fields>
                <asp:TemplateField ShowHeader="False">
                   <ItemTemplate> 
                            <table cellspacing="0" cellpadding="0" width="90%" align="center" border="0" style="text-align: left;">
                                    <tr>
                                        <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                                            <asp:Label ID="Label3" runat="server" CssClass="form" Text="流水號:"></asp:Label>&nbsp;
                                        </td>
                                        <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                                           <asp:Label ID="lb_Log_Guid" runat="server" CssClass="form" Width="100%" Text='<%# Bind("Log_Guid") %>' />
                                        </td>
                                        <td class="form" style="height: 25px; width: 12%; white-space: nowrap;">
                                            <asp:Label ID="Label4" runat="server" CssClass="form" Text="列表資料時間:" />
                                        </td>
                                        <td class="form" style="height: 25px; white-space: nowrap;">
                                            <asp:Label ID="lb_Print_Date" runat="server" CssClass="form" Text='<%# Bind("Print_Date") %>' />&nbsp;
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                                            <asp:Label ID="Label7" runat="server" CssClass="form" Text="複印文件資料名稱:"></asp:Label>&nbsp;
                                        </td>
                                        <td width="50%" class="form" colspan = "3" style="height: 25px; padding-right: 27px;">
                                            <asp:Label ID="lb_File_Name" runat="server" Width="100%" Text='<%# Bind("File_Name") %>' />
                                        </td>
                                        </tr>
                                    <tr>
                                        <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                                            <asp:Label ID="Label11" runat="server" CssClass="form" Text="使用單位:"></asp:Label>
                                        </td>
                                        <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                                            <asp:Label ID="lb_ORG_NAME" runat="server" CssClass="form" Text='<%# Bind("ORG_Name") %>' />
                                        </td> 
                                        <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                                            <asp:Label ID="Label5" runat="server" CssClass="form" Text="列印份數:"></asp:Label>
                                        </td>
                                       <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                                            <asp:label ID="lb_PrintSet_Cnt" runat="server" Text="1" />
                                        </td>
                                      </tr>
                                    <tr>
                                         <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                                            <asp:Label ID="Label1" runat="server" CssClass="form" Text="保密區分:"></asp:Label>
                                        </td>
                                         <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                                            <asp:Label ID="lb_Security_Status" runat="server" CssClass="form" Text='<%# ShowSecurity_Status(Eval("Security_Status"),Eval("Security_Guid"))%>' />&nbsp;
                                         </td>
                                         <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                                            <asp:Label ID="Label6" runat="server" CssClass="form" Text="作廢張數:"></asp:Label>&nbsp;
                                        </td>
                                        <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                                            <asp:Label ID="lb_Useless" runat="server" CssClass="form" Text='<%# Bind("Useless") %>'/>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                                            <asp:Label ID="Label9" runat="server" CssClass="form" Text="原資料張數:"></asp:Label>&nbsp;
                                        </td>
                                        <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                                            <%# ShowPrint(Eval("Copy_A3C"), Eval("Copy_A4C"), Eval("Copy_A3M"), Eval("Copy_A4M"), Eval("Scan"))%>
                                        </td>
                                        <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                                            <asp:Label ID="Label10" runat="server" CssClass="form" Text="總列印張數:"></asp:Label>&nbsp;
                                        </td>
                                        <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                                             <asp:Label ID="lb_PrintTotalCnt" runat="server" CssClass="form" Text='<%# Bind("PrintTotalCnt") %>' />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                                            <asp:Label ID="Label8" runat="server" CssClass="form" Text="送印人姓名:"></asp:Label>&nbsp;
                                        </td>
                                        <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                                            <asp:Label ID="lb_Print_Name" runat="server" CssClass="form" Text='<%# Bind("Print_Name") %>' />
                                        </td>
                                        <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                                            <asp:Label ID="Label12" runat="server" CssClass="form" Text="送印人級職:" />&nbsp;
                                        </td>
                                        <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                                            <asp:Label ID="lb_TU_ID_Name" runat="server" Text='<%# Bind("TU_Name") %>' />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="style1" valign="middle" style="white-space: nowrap;">
                                            <asp:Label ID="Label2" runat="server" CssClass="form" Text="用途:"></asp:Label>&nbsp;
                                        </td>
                                        <td colspan="3" class="style2" style="padding-right: 27px;">
                    
                                            <asp:TextBox ID="tb_Use_For" runat="server" MaxLength="255" 
                                                TextMode="MultiLine" Width="100%" Height="60px" Enabled = "false" Text='<%# Bind("Use_For") %>' />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="form" valign="middle" style="height: 80px; white-space: nowrap;">
                                            <asp:Label ID="Label13" runat="server" CssClass="form" Text="備註:"></asp:Label>&nbsp;
                                        </td>
                                        <td colspan="3" class="form" style="width: 99%; height: 40px; padding-right: 27px;">
                                            <asp:TextBox ID="tb_memo" runat="server" MaxLength="255" TextMode="MultiLine" Width="100%"
                                             Enabled="false" Height = "50px" Text='<%# Bind("memo") %>' />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td ><asp:Label ID="Label14" runat="server" CssClass="form" text = "登記流水號：" /></td>
                                        <td ><asp:Label ID="lb_Print_Num" runat="server" CssClass="form" Text='<%# Bind("Print_Num") %>' /></td>
                                        <td colspan = "2" style="padding-right: 22px;" >
                                            &nbsp;
                                        </td>
                                    </tr>
                                </table>
                  </ItemTemplate> 
                </asp:TemplateField>
                <asp:CommandField />
            </Fields>
        </asp:DetailsView>
        <br />
        <center> <asp:ImageButton ID="ibtnPrevious" runat="server" AlternateText="回上頁" ImageUrl="~/Image/backtop.gif" /></center>
    </form>
</body>
</html>
