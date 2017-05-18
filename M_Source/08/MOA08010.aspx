<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA08010.aspx.vb" Inherits="M_Source_08_MOA08010" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>機密資訊複(影)印資料原始文件</title>
    </head>
<body>
    <form id="form1" runat="server">
    <div>
    <br /><br /><br />
    <center style="font-family:標楷體; font-size:x-large;"><strong>機密資訊複影印申請單</strong></center>
    <center><br />
    <table cellpadding="0" style="border:1px solid Black; border-collapse:collapse; font-family:標楷體; font-size:small; width:650px;">
         <tr>
             <td align="center" valign="middle" colspan="3" style="border:1px solid Black; height:35px; font-size:large;">
                 <asp:Label ID="lbTop1unitName1" runat="server"></asp:Label>
                 複(影)印資料申請單</td>
         </tr>
         <tr>
             <td align="center" rowspan="5" style="border:1px solid Black; width:5%; height:110px;">原<br /><br />件<br /><br />資<br /><br />料</td>
             <td align="center" style="border:1px solid Black; width:20%; height:27px;">主旨、簡由<br />或(資料名稱)</td>
             <td style="border:1px solid Black; width:75%" align="left">
                 &nbsp;<asp:Label ID="lbSubject1" runat="server"></asp:Label>
             </td>
         </tr>
         <tr>
             <td align="center" style="border:1px solid Black; height:27px;">發文時間、字號</td>
             <td align="left" style="border:1px solid Black;">
                 &nbsp;<asp:Label ID="lbSignDateTime1" runat="server"></asp:Label>
                 &nbsp;<asp:Label ID="lbSecurity_No1" runat="server"></asp:Label>
             </td>
         </tr>
         <tr>
             <td align="center" colspan="2" style="border:1px solid Black;">
                 <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
                     <tr>
                         <td align="center" style="width:7%; height:27px;">機密<br />等級</td>
                         <td style="border-left:1px solid Black; width:16%"; align="left">
                             &nbsp;<asp:Label ID="lbSecurity_Level1" runat="server"></asp:Label>
                         </td>
                         <td align="center" style="border-left:1px solid Black; width:7%">機密<br />屬性</td>
                         <td style="border-left:1px solid Black; width:28%" align="left">
                             &nbsp;<asp:Label ID="lbSecurity_Type1" runat="server"></asp:Label>
                         </td>
                         <td align="center" style="border-left:1px solid Black; width:14%">保密期限/<br />解密條件</td>
                         <td style="border-left:1px solid Black; width:28%" align="left">
                             &nbsp;<asp:Label ID="lbSecurity_Range1" runat="server"></asp:Label>
                         </td>
                     </tr>
                 </table>
             </td>
         </tr>
         <tr>
             <td align="center" style="border:1px solid Black; height:27px;">產製單位</td>
             <td align="left" style="border:1px solid Black;">
                 &nbsp;<asp:Label ID="lbProduceUnit1" runat="server"></asp:Label>
             </td>
         </tr>
         <tr>
             <td align="center" style="border:1px solid Black; height:27px;">同意複(影)印<br />時間/文號</td>
             <td style="border:1px solid Black;">
                 <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
                     <tr>
                         <td align="left" style="width:40%">
                             &nbsp;<asp:Label ID="lbAgreeTimeOrNumber1" runat="server"></asp:Label>
                         </td>
                         <td align="center" style="border-left:1px solid Black; width:30%">同意複(影)印<br />權責長官級職姓名</td>
                         <td align="left" style="border-left:1px solid Black; width:30%">
                             &nbsp;<asp:Label ID="lbAgreeSuperior1" runat="server"></asp:Label>
                         </td>
                     </tr>
                 </table>
             </td>
         </tr>
         <tr>
             <td align="center" style="border:1px solid Black; height:110px;" rowspan="4">複<br /><br />印<br /><br />事<br /><br />項</td>
             <td align="center" style="border:1px solid Black; height:27px;">用途</td>
             <td align="left" style="border:1px solid Black;">
                 <asp:RadioButtonList ID="rblPurpose1" runat="server" 
                     RepeatDirection="Horizontal">
                     <asp:ListItem Enabled="False" Value="1">呈閱/</asp:ListItem>
                     <asp:ListItem Enabled="False" Value="2">分會、辦/</asp:ListItem>
                     <asp:ListItem Enabled="False" Value="3">作業用/</asp:ListItem>
                     <asp:ListItem Enabled="False" Value="4">歸檔/</asp:ListItem>
                     <asp:ListItem Enabled="False" Value="5">隨文分發/</asp:ListItem>
                     <asp:ListItem Enabled="False" Value="6">會議分發/</asp:ListItem>
                 </asp:RadioButtonList>
                 <asp:RadioButton ID="rbPurposeElse1" runat="server" Enabled="False" 
                     Text="其他：" />
                 <asp:Label ID="lbPurpose_Other1" runat="server"></asp:Label>
             </td>
         </tr>
         <tr>
             <td colspan="2" style="border:1px solid Black;">
                 <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%;">
                     <tr>
                         <td align="center" style="width:25%; height:27px;">複印時間</td>
                         <td align="left" style="border-left:1px solid Black; width:25%">
                             &nbsp;<asp:Label ID="lbPrinter_Datetime1" runat="server"></asp:Label>
                         </td>
                         <td align="center" style="border-left:1px solid Black; width:25%">浮水印暗記編號</td>
                         <td align="left" style="border-left:1px solid Black; width:25%">
                             &nbsp;<asp:Label ID="lbPrinter_Num1" runat="server"></asp:Label>
                         </td>
                     </tr>
                 </table>
             </td>
         </tr>
         <tr>
             <td colspan="2" style="border:1px solid Black;">
                 <table border="0" style="border-collapse:collapse; font-family:標楷體; width:100%; height:27px;">
                     <tr>
                         <td align="center" style="width:12%">原件張數</td>
                         <td align="left" style="border-left:1px solid Black; width:12%">
                             &nbsp;<asp:Label ID="lbOri_sheet1" runat="server"></asp:Label>
                         </td>
                         <td align="center" style="border-left:1px solid Black; width:12%">每張<br />複印張數</td>
                         <td align="left" style="border-left:1px solid Black; width:12%">
                             &nbsp;<asp:Label ID="lbCopy_sheet1" runat="server"></asp:Label>
                         </td>
                         <td align="center" style="border-left:1px solid Black; width:12%">合計<br />複印張數</td>
                         <td align="left" style="border-left:1px solid Black; width:12%">
                             &nbsp;<asp:Label ID="lbTotal_sheet1" runat="server"></asp:Label>
                         </td>
                         <td align="center" style="border-left:1px solid Black; width:16%">複(影)印<br />張數流水號</td>
                         <td align="left" style="border-left:1px solid Black; width:12%">
                             &nbsp;<asp:Label ID="lbSheet_ID1" runat="server"></asp:Label>
                         </td>
                     </tr>
                 </table>
             </td>
         </tr>
         <tr>
             <td align="center" style="border:1px solid Black; height:27px;">附註</td>
             <td align="left" style="border:1px solid Black;">
                 &nbsp;<asp:Label ID="lbMemo1" runat="server"></asp:Label>
             </td>
         </tr>
    </table>
    </center>
    </div>
    </form>
</body>
</html>
