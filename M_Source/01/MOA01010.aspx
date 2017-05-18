<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA01010.aspx.vb" Inherits="Source_01_MOA01010" %>
<%@ Register TagPrefix="rsweb" Namespace="Microsoft.Reporting.WebForms" Assembly="Microsoft.ReportViewer.WebForms, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>文職人員加班統計報表</title>
    <style type="text/css">
        body
        {
            height: 100%;
            margin: 0px 0px 0px 0px;
            padding: 0px 0px 0px 0px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="height: 640px; text-align: center;">
            <tr>
                <td valign="top" style="white-space: nowrap;">
                    <rsweb:ReportViewer ID="rvSignRecords" runat="server" Height="600px" Width="100%" ShowExportControls="False"></rsweb:ReportViewer>
                </td>
            </tr>
        </table>
        <asp:SqlDataSource ID="sdsSignRecords" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"></asp:SqlDataSource>
    </div>
    </form>
</body>
</html>