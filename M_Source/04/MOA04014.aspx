<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04014.aspx.vb" Inherits="Source_04_MOA04014" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>設備物料分類管理-新增</title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
    <script type="text/javascript" language="javascript" src="../Inc/inc_common.js"></script>
</head>
<body>
    <form id="form1" runat="server">
  <div>
        <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center">
                    <asp:Label ID="Label1" runat="server" CssClass="toptitle" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>
     <div style="width: 100; text-align: center; padding-top: 10px;">
            <table cellspacing="0" cellpadding="0" width="60%" align="center" border="0" style="text-align: left;">
                <tr>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label2" runat="server" CssClass="form" Text="物料編號:"></asp:Label>&nbsp;
                    </td>
                    <td colspan="3">&nbsp;</td>
                </tr>
                <tr>
                    <td colspan="4" style="padding-right: 27px; padding-bottom: 8px;">
                        <div style="border: solid 1px #808080; padding: 4px 4px 4px 0px; white-space: nowrap; width: 100%;">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                                    <td style="width: 33%; padding: 4px 2px 0px 4px;">
                                        <asp:DropDownList ID="ddl_it_code_L1" runat="server" AutoPostBack="True" DataTextField="it_name" DataValueField="it_code" Width="100%"></asp:DropDownList>
                                    </td>
                                    <td style="width: 33%; padding: 4px 2px 0px 2px;">
                                        <asp:DropDownList ID="ddl_it_code_L2" runat="server" AutoPostBack="True" DataTextField="it_name" DataValueField="it_code" Width="100%"></asp:DropDownList>
                                    </td>
                                    <td style="width: 33%; padding: 4px 0px 0px 2px;">
                                        <asp:DropDownList ID="ddl_it_code_L3" runat="server" AutoPostBack="True" DataTextField="it_name" DataValueField="it_code" Width="100%"></asp:DropDownList>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="form" style="height: 25px; white-space: nowrap; padding-left: 4px; padding-right: 2px; padding-top: 4px;">
                                        <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                            <tr>
                                                <td style="width: 50%; padding-right: 6px;">
                                                    <asp:TextBox ID="tb_it_code_front" runat="server" MaxLength="3" Width="100%"></asp:TextBox>
                                                </td>
                                                <td style="width: 4px;">&nbsp;</td>
                                                <td style="width: 50%; padding-right: 6px;">
                                                    <asp:TextBox ID="tb_it_code_back" runat="server" MaxLength="3" Width="100%"></asp:TextBox>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td colspan="2">&nbsp;</td>
                                </tr>
                            </table>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label3" runat="server" CssClass="form" Text="物料名稱:"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:TextBox ID="tb_it_name" runat="server" MaxLength="255" Width="100%"></asp:TextBox>
                    </td>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label4" runat="server" CssClass="form" Text="物料規格:"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:TextBox ID="tb_it_spec" runat="server" MaxLength="255" Width="100%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label5" runat="server" CssClass="form" Text="物料價格:"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:TextBox ID="tb_it_cost" runat="server" MaxLength = "4" Width="100%"></asp:TextBox>
                    </td>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label6" runat="server" CssClass="form" Text="物料公司:"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:TextBox ID="tb_manufacturer" runat="server" MaxLength="255" Width="100%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label8" runat="server" CssClass="form" Text="物料圖檔一:" />
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 20px;">
                        <table><tr><td>
                            <A HREF="#" onClick="window.open('<%= sfile_a_path %>', '物料圖檔一', config='height=200,width=200,menubar=no,toolbar=no,location=no,status=no,resizable=yes')">
                            <asp:Image ID="img_file_a" runat="server" Visible = "false" Width ="100px" Height ="100px" /></A> 
                         </td></tr></table>
                        <asp:MultiView ID="mv_file_a" runat="server">
                            <asp:View ID="v_file_a_upload" runat="server">
                                <asp:FileUpload ID="fu_file_a" runat="server" Width="100%" />
                            </asp:View>
                            <asp:View ID="v_file_a_modify" runat="server">
                                <asp:LinkButton ID="lkb_file_a" runat="server" CausesValidation="False" ForeColor="Blue"></asp:LinkButton>
                                <asp:Button ID="btn_del_file_a" runat="server" Height="20px" Text="刪除" Width="40px" />
                            </asp:View>
                        </asp:MultiView>
                    </td>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label9" runat="server" CssClass="form" Text="物料圖檔二:"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 20px;">
                        <table><tr><td>
                         <A HREF="#" onClick="window.open('<%= sfile_b_path %>', '物料圖檔二', config='height=200,width=200,menubar=no,toolbar=no,location=no,status=no,resizable=yes')">
                            <asp:Image ID="img_file_b" runat="server" Visible = "false" Width ="100px" Height ="100px" /></A>
                         </td></tr></table>
                        <asp:MultiView ID="mv_file_b" runat="server">
                            <asp:View ID="v_file_b_upload" runat="server">
                                <asp:FileUpload ID="fu_file_b" runat="server" Width="100%" />
                            </asp:View>
                            <asp:View ID="v_file_b_modify" runat="server">
                                <asp:LinkButton ID="lkb_file_b" runat="server" CausesValidation="False" ForeColor="Blue"></asp:LinkButton>
                                <asp:Button ID="btn_del_file_b" runat="server" Height="20px" Text="刪除" Width="40px" />
                            </asp:View>
                        </asp:MultiView>
                    </td>
                </tr>
                <tr>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label10" runat="server" CssClass="form" Text="安全數量:"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:TextBox ID="tb_snum" runat="server" Width="100%" MaxLength ="3" ></asp:TextBox>
                    </td>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label11" runat="server" CssClass="form" Text="有效期:"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:TextBox ID="tb_expired_y"  MaxLength = "4" runat="server" Width="100%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label12" runat="server" CssClass="form" Text="物料類別:"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:TextBox ID="tb_it_sort" runat="server" MaxLength="10" Width="100%"></asp:TextBox>
                    </td>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label13" runat="server" CssClass="form" Text="物料單位:"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:TextBox ID="tb_it_unit" runat="server" MaxLength="10" Width="100%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="form" valign="middle" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label7" runat="server" CssClass="form" Text="備註:"></asp:Label>&nbsp;
                    </td>
                    <td colspan="3" class="form" style="width: 99%; height: 25px; padding-right: 27px;">
                        <asp:TextBox ID="tb_memo" runat="server" MaxLength="255" TextMode="SingleLine" Width="100%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td colspan="3">&nbsp;</td>
                    <td style="padding-right: 21px;">
                        <div style="width: 100%; padding-top: 5px; text-align: right;">
                            <asp:ImageButton ID="ibtnAdd" runat="server" CommandName="Add" ImageUrl="~/Image/add.gif" ToolTip="新增" />
                            <asp:ImageButton ID="ibtnModify" runat="server" CommandName="Modify" ImageUrl="~/Image/update.gif" ToolTip="修改" />
                            <asp:ImageButton ID="ibtnPrevious" runat="server" AlternateText="回上頁" ImageUrl="~/Image/backtop.gif" />
                        </div>
                    </td>
                </tr>
            </table>
         
       </div>
    </div>
    <asp:SqlDataSource ID="sdsItCodeFilter" runat="server"
        ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
        SelectCommand="Select it_name = Case IsNull(it_spec, '') When '' Then it_name Else (it_name + ' - ' + it_spec) End, it_code ,file_a,file_b
                       From P_0407 Where it_code Like @it_code And (1 = 1) Order by it_name;">
        <SelectParameters>
            <asp:Parameter Name="it_code" DbType="String" Size="6" />
        </SelectParameters>
    </asp:SqlDataSource>
</form>
</body>
</html>