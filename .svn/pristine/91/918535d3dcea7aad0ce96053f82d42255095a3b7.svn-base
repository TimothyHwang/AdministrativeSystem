<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA08002.aspx.vb" Inherits="M_Source_08_MOA08002" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <%--<title>影印使用登記編輯</title>--%>
    <title>影印使用申請編輯</title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
    <script type="text/JavaScript">
         function Check() {
             var re = /</;
             if (re.test(document.getElementById("tb_Useless").value)) {
                 alert("您輸入的作廢張數有包含危險字元！");
                 document.getElementById("tb_Useless").value = "";
                 return false;
             }
            
             if (re.test(document.getElementById("tb_memo").value)) {
                 alert("您輸入的附註有包含危險字元！");
                 document.getElementById("tb_memo").value= "";
                 return false;
             }
             
             return true; 
         }
</script>
    <style type="text/css">
        .style1
        {
            width: 1%;
            height: 99px;
        }
        .style2
        {
            width: 99%;
            height: 99px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center">
                    <asp:Label ID="lblTitle" runat="server" CssClass="toptitle" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>
        <div style="width: 100; text-align: center; padding-top: 10px;">
        <table cellspacing="0" cellpadding="0" width="90%" align="center" border="0" style="text-align: left;">
                <tr>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label3" runat="server" CssClass="form" Text="申請時間：" 
                            Font-Size="Medium"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                       <asp:Label ID="lb_Log_Guid" runat="server" CssClass="form" Width="100%" 
                            visible= "False" Font-Size="Medium"  />
                       <asp:Label ID="lb_Org_Num" runat="server" CssClass="form" visible= "False" 
                            Font-Size="Medium" />
                       <asp:Label ID="lb_Log_Time" runat="server" CssClass="form" Font-Size="Medium"/>
                    </td>
                    <td class="form" style="height: 25px; width: 12%; white-space: nowrap;">
                        <asp:Label ID="Label4" runat="server" CssClass="form" Text="複印時間：" 
                            Font-Size="Medium"></asp:Label>&nbsp;
                    </td>
                    <td class="form" style="height: 25px; white-space: nowrap;">
                        <asp:Label ID="lb_Print_Date" runat="server" CssClass="form" 
                            Font-Size="Medium" />&nbsp;
                    </td>
                </tr>
                <tr>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label7" runat="server" CssClass="form" Text="複印資料名稱：" 
                            Font-Size="Medium"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:Label ID="lb_File_Name" runat="server" Font-Size="Medium"></asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="Label11" runat="server" CssClass="form" Text="使用單位：" 
                            Font-Size="Medium"></asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lb_ORG_NAME" runat="server" CssClass="form" Font-Size="Medium" />
                    </td>
                </tr>
                <tr>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap; font-size:medium;">
                        申請單原件張數：</td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:Label ID="lbOri_sheet" 
                                        runat="server" CssClass="form" Font-Size="Medium"></asp:Label></td> 
                    <td class="form" style="height: 25px; white-space: nowrap; font-size:medium;">
                        申請單每份複印張數：</td>
                    <td class="form" style="height: 25px; white-space: nowrap;">
                        <asp:Label ID="lbCopy_sheet" runat="server" CssClass="form" Font-Size="Medium"></asp:Label></td>
                </tr>
                <tr>
                    <td class="form" style="height: 25px; white-space: nowrap; font-size:medium;">
                        申請單合計複印張數：</td>
                    <td class="form" style="height: 25px; white-space: nowrap;">
                        <asp:Label ID="lbTotal_sheet" runat="server" CssClass="form" Font-Size="Medium"></asp:Label></td>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label5" runat="server" CssClass="form" Text="列印份數：" 
                            Visible="False" Font-Size="Medium"></asp:Label>
                    </td>
                   <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        
                    </td>
                </tr> 
                <tr>
                      <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label1" runat="server" CssClass="form" Text="保密區分：" 
                              Font-Size="Medium"></asp:Label>
                    </td>
                     <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:DropDownList ID="ddl_Security_Status" runat="server" Enabled = "false" 
                             Font-Size="Medium">
                            <asp:ListItem Value="1">普通</asp:ListItem>
                            <asp:ListItem Value="2">密</asp:ListItem>
                            <asp:ListItem Value="3">機密</asp:ListItem>
                            <asp:ListItem Value="4">極機密</asp:ListItem>
                            <asp:ListItem Value="5">絕對機密</asp:ListItem>
                        </asp:DropDownList>
                        <asp:HyperLink ID="HLSecurity" runat="server" Visible="False" Target ="_blank" 
                             Font-Size="Medium">檢視</asp:HyperLink>
                     </td>
                     <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label6" runat="server" CssClass="form" Text="作廢張數：" 
                             Font-Size="Medium"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:TextBox ID="tb_Useless" runat="server" MaxLength = "3" Text = "0" 
                            AutoPostBack="True" Font-Size="Medium" Width="65px"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label9" runat="server" CssClass="form" Text="資料張數：" 
                            Font-Size="Medium"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                         <asp:Label ID="lb_Print1" runat="server" MaxLength="" Font-Size="Medium" />
                          <asp:Label ID="lb_Print3" runat="server" MaxLength="" Font-Size="Medium" />
                          <%  If (lb_Print2.Text.Length > 0) Then%>
                                  <br />
                              <% End If %>
                          <asp:Label ID="lb_Print2" runat="server" MaxLength="" Font-Size="Medium" />
                    </td>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label10" runat="server" CssClass="form" Text="總張數：" 
                            Font-Size="Medium"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                         <asp:Label ID="lb_PrintTotalCnt" runat="server" CssClass="form" 
                             Font-Size="Medium"  />
                    </td>
                </tr>
                <tr>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label8" runat="server" CssClass="form" Text="送印人姓名：" 
                            Font-Size="Medium"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:Label ID="lb_Print_Name" runat="server" CssClass="form" 
                            Font-Size="Medium" />
                        (<asp:Label ID="lb_PAIDNO" runat="server" Font-Size="Medium" />)
                    </td>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label12" runat="server" CssClass="form" Text="送印人級職：" 
                            Font-Size="Medium" />&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:Label ID="lb_TU_ID_Name" runat="server" Font-Size="Medium" />                        
                    </td>
                </tr>
                <tr>
                    <td class="style1" valign="middle" style="white-space: nowrap;">
                        <asp:Label ID="Label2" runat="server" CssClass="form" Text="用途：" 
                            Font-Size="Medium"></asp:Label>&nbsp;
                    </td>
                    <td colspan="3" class="style2" style="padding-right: 27px;">
                        <asp:TextBox ID="tb_Use_For" runat="server" MaxLength="255" 
                            TextMode="MultiLine" Width="100%" Height="60px" Enabled="False" 
                            Font-Size="Medium"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="form" valign="middle" style="height: 50px; white-space: nowrap;">
                        <asp:Label ID="Label13" runat="server" CssClass="form" Text="備註：" 
                            Font-Size="Medium"></asp:Label>&nbsp;
                    </td>
                    <td colspan="3" class="form" style="width: 99%; padding-right: 27px;" height="80px">
                        <asp:TextBox ID="tb_memo" runat="server" MaxLength="255" TextMode="MultiLine" 
                            Width="100%"  Height="60px" Font-Size="Medium" />
                    </td>
                </tr>
                <tr>
                    <td ><asp:Label ID="Label14" runat="server" CssClass="form" text = "登記流水號：" 
                            Font-Size="Medium" /></td>
                    <td ><asp:Label ID="lb_Print_Num" runat="server" CssClass="form" 
                            Font-Size="Medium"  /></td>
                    <td></td>
                    <td></td>
                </tr>
                <tr><td colspan="4" class="form" style="height: 25px;">
                <div style="border-top:1px dashed #666666;height: 1px;overflow:hidden;"></div>
                </td></tr>
                <tr>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap; font-size:medium;">
                        呈核申請人：</td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:Label ID="lbVerifyRequester" 
                                        runat="server" CssClass="form" Font-Size="Medium"></asp:Label></td> 
                    <td class="form" style="height: 25px; white-space: nowrap; font-size:medium;">
                        批示：</td>
                    <td class="form" style="height: 25px; white-space: nowrap;">
                        <asp:Label ID="lbApprovedBy" runat="server" CssClass="form" Font-Size="Medium"></asp:Label></td>
                </tr>
                <tr>
                    <td colspan = "4" align ="left" ><asp:Label ID="lbMsg" runat="server" Font-Bold="True" Font-Size="Small" ForeColor="Red" /></td>
                </tr>
                <tr>
                    <td></td>
                    <td></td>
                    <td colspan = "2" style="padding-right: 22px;"  >
                        <div style="width: 100%; padding-top: 5px; text-align: right;">
                            <asp:ImageButton ID="ibtnModify" runat="server" CommandName="Modify" ImageUrl="~/Image/mend.gif" ToolTip="修改" OnClientClick ="return Check();" />
                            <asp:ImageButton ID="ibtnPrevious" runat="server" AlternateText="回上頁" ImageUrl="~/Image/backtop.gif" onclientclick="form1.reset();"/>
                        </div>
                    </td>
                </tr>
            </table>
      
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT ORG_UID, ORG_NAME FROM ADMINGROUP ORDER BY ORG_NAME">
        </asp:SqlDataSource>
        </div>
    </form>
</body>
</html>
