<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA10003.aspx.vb" Inherits="M_Source_10_MOA10003" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>維護主官負責人</title>
    <link href='<%#ResolveUrl("~/Styles.css") %>' rel="stylesheet" type="text/css" />
    <script src='<%#ResolveUrl("~/Script/jquery-1.10.2.min.js") %>' type="text/javascript"></script>
    <style type="text/css">  
        .treetd
        {
            width: 250px;
            vertical-align: top;
        }
        .valignTop {
            vertical-align: top;
        }
        .maintable {
            width: 100%;            
        }
        .maintable .caption {
            vertical-align: bottom;
            background-color: #6699cc;
            border-color: #66aaaa;            
            color: white;
            text-align: center;
        }
         
    </style>
    <script language="javascript" type="text/javascript">
        function public_GetParentByTagName(element, tagName) {
            var parent;
            var upperTagName;

            parent = element.parentNode;
            upperTagName = tagName.toUpperCase();
            while (parent && (parent.tagName.toUpperCase() != upperTagName)) {
                parent = parent.parentNode ? parent.parentNode : parent.parentElement;
            }
            return parent;
        }

        function setParentChecked(objNode) {
            var objParentDiv;
            var objID;
            var objParentCheckBox;

            objParentDiv = public_GetParentByTagName(objNode, "div");
            if (objParentDiv == null || objParentDiv == "undefined") {
                return;
            }
            objID = objParentDiv.getAttribute("ID");
            objID = objID.substring(0, objID.indexOf("Nodes"));
            objID = objID + "CheckBox";
            objParentCheckBox = document.getElementById(objID);
            if (objParentCheckBox == null || objParentCheckBox == "undefined") {
                return;
            }
            if (objParentCheckBox.tagName != "INPUT" && objParentCheckBox.type == "checkbox")
                return;

            objParentCheckBox.checked = true;
            setParentChecked(objParentCheckBox);
        }

        function setChildUnChecked(divID) {
            var objchild;
            var count;
            var tempObj;

            objchild = divID.children;
            count = objchild.length;
            for (var i = 0; i < objchild.length; i++) {
                tempObj = objchild[i];
                if (tempObj.tagName == "INPUT" && tempObj.type == "checkbox") {
                    tempObj.checked = false;
                }
                setChildUnChecked(tempObj);
            }
        }

        function setChildChecked(divID) {
            var objchild;
            var count;
            var tempObj;

            objchild = divID.children;
            count = objchild.length;
            for (var i = 0; i < objchild.length; i++) {
                tempObj = objchild[i];

                if (tempObj.tagName == "INPUT" && tempObj.type == "checkbox") {
                    tempObj.checked = true;
                }
                setChildChecked(tempObj);
            }
        }

        function CheckEvent() {
            var objNode;
            var objID;
            var objParentDiv;

            objNode = event.srcElement;
            if (objNode.tagName != "INPUT" || objNode.type != "checkbox")
                return;

            if (objNode.checked == true) {
                setParentChecked(objNode);
                objID = objNode.getAttribute("ID");
                objID = objID.substring(0, objID.indexOf("CheckBox"));
                objParentDiv = document.getElementById(objID + "Nodes");
                if (objParentDiv == null || objParentDiv == "undefined") {
                    return;
                }
                setChildChecked(objParentDiv);
            } else {
                objID = objNode.getAttribute("ID");
                objID = objID.substring(0, objID.indexOf("CheckBox"));
                objParentDiv = document.getElementById(objID + "Nodes");
                if (objParentDiv == null || objParentDiv == "undefined") {
                    return;
                }
                setChildUnChecked(objParentDiv);
            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <asp:ScriptManager ID="sm" runat="server">
        </asp:ScriptManager>
        <table class="maintable">
            <tr>
                <td class="caption" colspan="2"><b>維護主官負責人</b></td>
            </tr>
            <tr>
                <td class="treetd">
                    <div style="background-color:#6499CD;text-align: center;color: #FFFFFF;">
                        <asp:Label ID="lblManager" runat="server" Text="Label" Width="100%"></asp:Label>                                    
                    </div>
                    <div id="MemberTree">
                        <asp:TreeView ID="tvEmployee" runat="server" ExpandDepth="1">
                            </asp:TreeView>
                    </div>
                </td>
                <td style="vertical-align: top">
                    <div id="MemberGridView">
                        <asp:UpdatePanel ID="UPMembers" runat="server" >
                            <ContentTemplate>
                                 <asp:GridView ID="gvMembers" runat="server" 
                                     AllowSorting="True" AutoGenerateColumns="False" CellPadding="4" 
                                     DataKeyNames="empuid" DataSourceID="SqlDataSourceMembers" 
                                     EnableModelValidation="True" ForeColor="#333333" GridLines="None" 
                                     PageSize="20">
                                     <AlternatingRowStyle BackColor="White" />
                                     <Columns>
                                         <asp:TemplateField HeaderText="選取">
                                             <ItemTemplate>
                                                 <asp:CheckBox ID="chkSelect" runat="server" />
                                             </ItemTemplate>
                                         </asp:TemplateField>
                                         <asp:BoundField DataField="empuid" HeaderText="編號" InsertVisible="False" 
                                             ReadOnly="True" SortExpression="empuid" />
                                         <asp:BoundField DataField="employee_id" HeaderText="帳號" 
                                             SortExpression="employee_id" />
                                         <asp:BoundField DataField="emp_chinese_name" HeaderText="姓名" 
                                             SortExpression="emp_chinese_name" />
                                     </Columns>
                                     <EditRowStyle BackColor="#2461BF" />
                                     <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                     <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                     <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                                     <RowStyle BackColor="#EFF3FB" />
                                     <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                                 </asp:GridView>
                                <asp:Button ID="btnEmployeeModify" runat="server" Text="確定" Visible="False" />
                                <br />
                                <asp:SqlDataSource ID="SqlDataSourceMembers" runat="server" 
                                    ConnectionString="<%$ ConnectionStrings:ConnectionString %>"                         
                                    SelectCommand="SELECT * FROM [EMPLOYEE] WHERE ([ORG_UID] = @ORG_UID) AND LEAVE='Y'">
                                    <SelectParameters>
                                        <asp:ControlParameter ControlID="tvEmployee" Name="ORG_UID" 
                                            PropertyName="SelectedValue" Type="String" />
                                    </SelectParameters>
                                </asp:SqlDataSource>
                            </ContentTemplate> 
                            <Triggers>
                                <asp:AsyncPostBackTrigger ControlID="tvEmployee" EventName="SelectedNodeChanged" />
                            </Triggers>
                        </asp:UpdatePanel>
                    </div>
                </td>            
            </tr>
        </table>
    </form>
</body>
</html>
