<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04011.aspx.vb" Inherits="Source_04_MOA04011" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>設備識別編碼管理-新增</title>
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
                <asp:MultiView ID="mvMultiInputParameters" runat="server">
                    <asp:View ID="vCreateNew" runat="server">
                        <tr>
                            <td nowrap class="form" style="height: 25px; width: 10%;">
                                <asp:Label ID="Label12" runat="server" CssClass="form" Text="資料新增筆數:" Width="77px"></asp:Label>
                            </td>
                            <td nowrap width="15%" class="form" style="height: 25px;">
                                <asp:TextBox ID="tb_records_count" runat="server" Width="92%" Text="1"></asp:TextBox>
                            </td>
                            <td colspan="2">&nbsp;</td>
                        </tr>
                    </asp:View>
                    <asp:View ID="vModification" runat="server">
                        <tr>
                            <td nowrap class="form" style="height: 25px; width: 10%;">
                                <asp:Label ID="Label3" runat="server" CssClass="form" Text="設備編號:" Width="77px"></asp:Label>
                            </td>
                            <td nowrap width="15%" class="form" style="height: 25px;">
                                <asp:Label ID="lbl_element_no" runat="server" Width="90%"></asp:Label>
                            </td>
                            <td colspan="2">&nbsp;</td>
                        </tr>
                    </asp:View>
                </asp:MultiView>
                <tr>
                    <td nowrap class="form" style="height: 25px; width: 10%;">
                        <asp:Label ID="Label4" runat="server" CssClass="form" Text="建物編號:" Width="77px"></asp:Label>
                    </td>
                    <td nowrap width="15%" class="form" style="height: 25px;">
                        <asp:DropDownList ID="ddl_bd_code" runat="server" AutoPostBack="true" DataTextField="bd_name" DataValueField="bd_code" Width="95%"></asp:DropDownList>
                    </td>
                    <td nowrap class="form" style="height: 25px; width: 10%;">
                        <asp:Label ID="Label5" runat="server" CssClass="form" Text="樓層編號:" Width="77px"></asp:Label>
                    </td>
                    <td nowrap width="15%" class="form" style="height: 25px;">
                        <asp:DropDownList ID="ddl_fl_code" runat="server" AutoPostBack="true" DataTextField="fl_name" DataValueField="fl_code" Width="95%"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td nowrap class="form" style="height: 25px; width: 10%;">
                        <asp:Label ID="Label6" runat="server" CssClass="form" Text="房間(區域)編號:" Width="77px"></asp:Label>
                    </td>
                    <td nowrap width="15%" class="form" style="height: 25px;">
                        <asp:DropDownList ID="ddl_rnum_code" runat="server" DataTextField="rnum_name" DataValueField="rnum_code" Width="95%"></asp:DropDownList>
                    </td>
                    <td nowrap class="form" style="height: 25px; width: 10%;">
                        <asp:Label ID="Label7" runat="server" CssClass="form" Text="牆柱編號:" Width="77px"></asp:Label>
                    </td>
                    <td nowrap width="15%" class="form" style="height: 25px;">
                        <asp:DropDownList ID="ddl_wa_code" runat="server" DataTextField="wa_name" DataValueField="wa_code" Width="95%"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td nowrap class="form" style="height: 25px; width: 10%;">
                        <asp:Label ID="Label8" runat="server" CssClass="form" Text="預算編號:" Width="77px"></asp:Label>
                    </td>
                    <td nowrap width="15%" class="form" style="height: 25px;">
                        <asp:DropDownList ID="ddl_bg_code" runat="server" DataTextField="bg_name" DataValueField="bg_code" Width="95%"></asp:DropDownList>
                    </td>
                    <td colspan="2">&nbsp;</td>
                </tr>
                <tr>
                    <td nowrap class="form" style="height: 25px; width: 10%;">
                        <asp:Label ID="Label9" runat="server" CssClass="form" Text="物料編號:" Width="77px"></asp:Label>
                    </td>
                    <td nowrap width="15%" class="form" style="height: 25px;">
                        <asp:TextBox ID="tb_it_code" runat="server" Width="92%" MaxLength="6" ReadOnly="true" Enabled="false"></asp:TextBox>
                        <input type="hidden" id="hid_it_code" name="hid_it_code" />
                    </td>
                    <td colspan="2">&nbsp;</td>
                </tr>
                <tr>
                    <td colspan="4" style="padding-top: 4px;">
                        <div style="border: solid 1px #808080; padding: 8px 0px 8px 8px; white-space: nowrap; width: 97%;">
                            <asp:DropDownList ID="ddl_it_code_L1" runat="server" AutoPostBack="True" DataTextField="it_name" DataValueField="it_code" Width="32%"></asp:DropDownList>
                            <asp:DropDownList ID="ddl_it_code_L2" runat="server" AutoPostBack="True" DataTextField="it_name" DataValueField="it_code" Width="32%"></asp:DropDownList>
                            <asp:DropDownList ID="ddl_it_code_L3" runat="server" AutoPostBack="True" DataTextField="it_name" DataValueField="it_code" Width="32%"></asp:DropDownList>
                            <div style="margin-top: 4px;">
                                <asp:ListBox ID="list_it_code" runat="server" DataTextField="it_name" DataValueField="it_code" Rows="10" Width="98%">
                                </asp:ListBox>
                            </div>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td colspan="3">&nbsp;</td>
                    <td>
                        <div style="width: 95%; padding-top: 5px; text-align: right;">
                            <asp:ImageButton ID="ibtnAdd" runat="server" CommandName="Add" ImageUrl="~/Image/add.gif" ToolTip="新增" />
                            <asp:ImageButton ID="ibtnModify" runat="server" CommandName="Modify" ImageUrl="~/Image/update.gif" ToolTip="修改" />
                            <asp:ImageButton ID="ibtnPrevious" runat="server" AlternateText="回上頁" ImageUrl="~/Image/backtop.gif" />
                        </div>
                    </td>
                </tr>
            </table>
        </div>
    </div>
    </form>
</body>
</html>