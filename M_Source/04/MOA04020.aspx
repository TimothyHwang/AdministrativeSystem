<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA04020.aspx.vb" Inherits="Source_04_MOA04020" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link type="text/css" rel="Stylesheet" href="../../Styles.css" />
    <script type="text/javascript" language="javascript" src="../Inc/inc_common.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table border="0" width="100%" style="z-index: 101; left: 104px; top: 33px;">
            <tr>
                <td align="center">
                    <asp:Label ID="lblTitle" runat="server" CssClass="toptitle" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>
        <div style="width: 100; text-align: center; padding-top: 10px;">
            <table cellspacing="0" cellpadding="0" width="70%" align="center" border="0" style="text-align: left;">
                <asp:MultiView ID="mvMultiInputParameters" runat="server">
                    <asp:View ID="vCreateNew" runat="server">
                        <tr>
                            <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                                <asp:Label ID="Label1" runat="server" CssClass="form" Text="資料新增筆數:"></asp:Label>&nbsp;
                            </td>
                            <td style="white-space: nowrap;">
                                <asp:TextBox ID="tb_records_count" runat="server" Columns="23" Text="1"></asp:TextBox>
                            </td>
                            <td colspan="2">&nbsp;</td>
                        </tr>
                    </asp:View>
                    <asp:View ID="vModification" runat="server"></asp:View>
                </asp:MultiView>
                <tr>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label2" runat="server" CssClass="form" Text="物品編號:"></asp:Label>&nbsp;
                    </td>
                    <td colspan="3" style="white-space: nowrap;">
                        <span style="padding-right: 2px;">
                            <asp:TextBox ID="tb_shcode_first" runat="server" Columns="6" Enabled="false" MaxLength="6" ReadOnly="true"></asp:TextBox>
                        </span>
                        <span style="padding-right: 1px;">
                            <asp:TextBox ID="tb_shcode_middle" runat="server" Columns="8" Enabled="false" MaxLength="8" ReadOnly="true"></asp:TextBox>
                        </span>
                        <asp:TextBox ID="tb_shcode_last" runat="server" Columns="4" Enabled="false" MaxLength="4" ReadOnly="true"></asp:TextBox>
                        <asp:TextBox ID="tb_schcode_lastest" runat="server" Columns="4" Enabled="false" MaxLength="5" ReadOnly="true" Visible="false" />
                    </td>
                </tr>
                <tr>
                    <td colspan="4" style="padding-right: 28px; padding-top: 4px; padding-bottom: 8px;">
                        <div style="border: solid 1px #808080; padding: 6px 0px 6px 6px; white-space: nowrap; width: 100%;">
                            <asp:DropDownList ID="ddl_it_code_L1" runat="server" AutoPostBack="True" DataTextField="it_name" DataValueField="it_code" Width="32%"></asp:DropDownList>
                            <asp:DropDownList ID="ddl_it_code_L2" runat="server" AutoPostBack="True" DataTextField="it_name" DataValueField="it_code" Width="33%"></asp:DropDownList>
                            <asp:DropDownList ID="ddl_it_code_L3" runat="server" AutoPostBack="True" DataTextField="it_name" DataValueField="it_code" Width="33%"></asp:DropDownList>
                            <div style="margin-top: 4px;">
                                <asp:ListBox ID="list_it_code" runat="server" AutoPostBack="true" DataTextField="it_name" DataValueField="it_code" Rows="5" Width="99%"></asp:ListBox>
                            </div>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label3" runat="server" CssClass="form" Text="物料價格:"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:TextBox ID="tb_shcost" runat="server" Width="100%"></asp:TextBox>
                    </td>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label4" runat="server" CssClass="form" Text="廠商名稱:"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:TextBox ID="tb_company" runat="server" MaxLength="255" Width="100%"></asp:TextBox>
                    </td>
                </tr>
                
                <tr>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label7" runat="server" CssClass="form" Text="有效期:"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:TextBox ID="tb_expired_y" runat="server" Width="100%"></asp:TextBox>
                    </td>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label11" runat="server" CssClass="form" Text="廠商產品貨號:"></asp:Label>
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:TextBox ID="tb_company_num" runat="server" MaxLength="50" Width="100%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label9" runat="server" CssClass="form" Text="採購編號:"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:TextBox ID="tb_buy_num" runat="server" MaxLength="20" Width="100%"></asp:TextBox>
                    </td>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label10" runat="server" CssClass="form" Text="儲庫編號:"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:DropDownList ID="ddl_seat_num" runat="server" DataSourceID = "sdsItSeatNum" DataTextField ="seat_name" DataValueField = "seat_num" >
                        </asp:DropDownList>
                    </td>
                </tr>
                 <tr>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label8" runat="server" CssClass="form" Text="物料圖檔一:"></asp:Label>&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:Image ID="it_image1" runat="server" Width = "100px"  Visible = "false" />
                        <asp:Label ID="lb_image1" runat="server" Text="本物料未上傳圖檔一之圖片" Visible = "false" />
                    </td>
                    <td class="form" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label12" runat="server" CssClass="form" Text="物料圖檔二:" />&nbsp;
                    </td>
                    <td width="50%" class="form" style="height: 25px; padding-right: 27px;">
                        <asp:Image ID="it_image2" runat="server" Width = "100px"  Visible = "false" />
                        <asp:Label ID="lb_image2" runat="server" Text="本物料未上傳圖檔二之圖片" Visible = "false" />
                    </td>
                </tr>
                <tr>
                    <td class="form" valign="middle" style="height: 25px; width: 1%; white-space: nowrap;">
                        <asp:Label ID="Label13" runat="server" CssClass="form" Text="備註:"></asp:Label>&nbsp;
                    </td>
                    <td colspan="3" class="form" style="width: 99%; height: 25px; padding-right: 27px;">
                        <asp:TextBox ID="tb_memo" runat="server" MaxLength="255" TextMode="SingleLine" Width="100%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td colspan="3">&nbsp;</td>
                    <td style="padding-right: 22px;">
                        <div style="width: 100%; padding-top: 5px; text-align: right;">
                            <asp:ImageButton ID="ibtnAdd" runat="server" CommandName="Add" ImageUrl="~/Image/add.gif" ToolTip="新增" />
                            <asp:ImageButton ID="ibtnModify" runat="server" CommandName="Modify" ImageUrl="~/Image/update.gif" ToolTip="修改" />
                            <asp:ImageButton ID="ibtnPrevious" runat="server" AlternateText="回上頁" ImageUrl="~/Image/backtop.gif" onclientclick="form1.reset();"/>
                        </div>
                    </td>
                </tr>
            </table>
        </div>
    </div>
    <asp:SqlDataSource ID="sdsItCodeFilter" runat="server"

        ConnectionString="<%$ ConnectionStrings:ConnectionString %>"

        SelectCommand="Select it_name = Case IsNull(it_spec, '') When '' Then it_name Else (it_name) End, 
                        (it_code+','+ convert(varchar,isnull(it_cost,'-1')) + ',' + convert(varchar,isnull(expired_y,'-1')) + ',' + isnull(file_a,'') + ',' + isnull(file_b,'')) as it_code ,it_sort From 
                       P_0407 Where it_code Like IsNull(@it_code, '') And it_code &lt;&gt; IsNull(@it_code_ignore, '') Order by it_name;">

        <SelectParameters>
            <asp:Parameter Name="it_code" DbType="String" Size="6" />
            <asp:Parameter Name="it_code_ignore" DbType="String" Size="6" />
        </SelectParameters>

    </asp:SqlDataSource>

    <asp:SqlDataSource ID="sdsItSeatNum" runat="server"
        ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
        SelectCommand="select seat_num+'_'+seat_name as seat_name,seat_num from P_0417 with(nolock) order by seat_num ">
    </asp:SqlDataSource>
    <asp:GridView ID="GridView1" runat="server">
    </asp:GridView>
    </form>
</body>
</html>