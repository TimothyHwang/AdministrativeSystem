<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA09005.aspx.vb" Inherits="M_Source_09_MOA09005" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>選擇欲呈核之門禁會議管制群組成員</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">    
        <table border="0" width="100%" style="Z-INDEX: 101; LEFT: 104px; TOP: 33px;">
            <tr>
            <td align="center">
                    <asp:Label ID="LabTitle" runat="server" CssClass="toptitle" Text="選擇欲呈核之門禁會議管制群組成員" Width="100%"></asp:Label>
            </td>
            </tr>
        </table>
        <table width="100%" border="3" bordercolor="#ccddee" align="center">
            <tr>
            <td align="center" class="CellClass" style="width: 25%">
                <asp:Label ID="LabName" runat="server" Text="門禁會議管制群組成員：" ></asp:Label>
            </td>
            <td style="width: 75%">
                <asp:DropDownList ID="DDLUser" runat="server" DataSourceID="SqlDataSource1" DataTextField="emp_chinese_name"
                    DataValueField="employee_id">
                </asp:DropDownList>
            </td>
            </tr>
            <tr>
            <td colspan="2" align="center">
                <asp:Button ID="btnSend" runat="server" Text="送件" />
                <asp:HiddenField ID="hidEFormID" runat="server" />
                <asp:HiddenField ID="hidEFormSN" runat="server" />
                <asp:HiddenField ID="hidFillOutBy" runat="server" />
                <asp:HiddenField ID="hidCreateBy" runat="server" />
                <asp:HiddenField ID="hidSubject" runat="server" />
                <asp:HiddenField ID="hidLocation" runat="server" />
                <asp:HiddenField ID="hidSponsor" runat="server" />
                <asp:HiddenField ID="hidModerator" runat="server" />
                <asp:HiddenField ID="hidPhoneNumber" runat="server" />
                <asp:HiddenField ID="hidMeetingDate" runat="server" />
                <asp:HiddenField ID="hidHour" runat="server" />
                <asp:HiddenField ID="hidMinute" runat="server" />
                <asp:HiddenField ID="hidDocumentNo" runat="server" />
                <asp:HiddenField ID="hidEnteringPeopleNumber" runat="server" />
                <asp:HiddenField ID="hidEnteringGate" runat="server" />
                </td>
            </tr>
        </table>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand="SELECT SU.employee_id,E.emp_chinese_name FROM SYSTEMOBJUSE AS SU JOIN SYSTEMOBJ AS S ON SU.object_uid = S.object_uid JOIN EMPLOYEE AS E ON SU.employee_id = E.employee_id WHERE S.object_name = N'門禁會議管制群組'">
        </asp:SqlDataSource>
    </form>
</body>
</html>
