<%@ Page AutoEventWireup="false" CodeFile="MOA11001.aspx.vb" Inherits="M_Source._11.M_Source_11_MOA11001" Language="VB" %>

<%@ Register Src="../90/FlowRoute.ascx" TagName="FlowRoute" TagPrefix="uc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>資訊設備維修申請單</title>
    <link href="../../Styles.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        .table
        {
            border: 1px solid black;
            background-color: #ffffff;
            border-spacing: 0px;
            border-color: #74a3d6 #000000 #000000 #74a3d6;
            margin: 0px auto;
            left: 20px;
            top: 10px;
            width: 750px;
        }
        th, tr, td
        {
            padding: 5px;
        }
        .caption
        {
            vertical-align: text-bottom;
            background-color: #6699cc;
            border-color: #66aaaa #ffffff #ffffff #66aaaa;
            color: #ffffff;
        }
        .hide
        {
            display: none;
        }
        .style1
        {
            width: 120px;
            height: 26px;
        }
        .style2
        {
            width: 550px;
            height: 26px;
        }
    </style>
    <link href="../../css/jquery.datepick.css" rel="stylesheet" type="text/css" />
    <script src="../../script/jquery-1.10.2.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.min.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick.js" type="text/javascript"></script>
    <script src="../../script/jquery.datepick-zh-TW.js" type="text/javascript"></script>
    <script type="text/javascript" language="javascript">
        $(function () {
            $("#txtCallDate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#txtCallDate").val(), showTrigger: '#calImg' });
            $("#txtArriveDate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#txtArriveDate").val(), showTrigger: '#calImg' });
            $("#txtFinalDate").datepick({ formats: 'yyyy/m/d', defaultDate: $("#txtFinalDate").val(), showTrigger: '#calImg' });
        });        
    </script>
</head>
<body lang="javascript" onload="return window_onload()">
    <form id="form1" runat="server" enctype="multipart/form-data">
    <div>
        <table width="750" border="1" cellspacing="0" cellpadding="5" bgcolor="#ffffff" bordercolor="#6699cc"
            bordercolorlight="#74a3d6" bordercolordark="#000000" style="left: 20px; top: 10px">
            <tr>
                <td valign="bottom" bgcolor="#6699cc" bordercolorlight="#66aaaa" bordercolordark="#ffffff">
                    <font color="white"><b>&nbsp;資訊設備修繕申請單</b></font>
                </td>
            </tr>
            <tr>
                <td>
                    <fieldset id="tableB" style="width: 750px">
                        <table style="width: 750px; height: 57px">
                            <tr>
                                <td style="width: 240px">
                                    <asp:Label ID="Label1" runat="server" Text="填表人單位：" Width="80px" ForeColor="Black"
                                        CssClass="form"></asp:Label>
                                    <asp:Label ID="Lab_PWUNIT" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label>
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
                    <asp:DropDownList id="ddlPAUNIT"
                        DataSourceID="SqlDataSource7"
                        DataValueField="ORG_UID"
                        DataTextField="ORG_NAME"
                        runat="server" AutoPostBack="True">
                    </asp:DropDownList>
                                    <asp:SqlDataSource ID="SqlDataSource7" runat="server" 
                                        ConnectionString="<%$ ConnectionStrings:ConnectionString %>" 
                                        SelectCommand="SELECT * FROM [ADMINGROUP]"></asp:SqlDataSource>
                                </td>
                                <td style="width: 220px">
                                    <asp:Label ID="Label5" runat="server" Text="姓名：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>&nbsp;
                                    <asp:DropDownList ID="DrDown_PANAME" runat="server" Width="143px" 
                                        AutoPostBack="True" DataSourceID="SqlDataSource12" 
                                        DataTextField="emp_chinese_name" DataValueField="employee_id">
                                    </asp:DropDownList>
                                    <asp:SqlDataSource ID="SqlDataSource12" runat="server" 
                                        ConnectionString="<%$ ConnectionStrings:ConnectionString %>" 
                                        SelectCommand="SELECT [employee_id], [emp_chinese_name], [ORG_UID] FROM [EMPLOYEE] WHERE ([ORG_UID] = @ORG_UID) ORDER BY [emp_chinese_name]">
                                        <SelectParameters>
                                            <asp:ControlParameter ControlID="ddlPAUNIT" Name="ORG_UID" 
                                                PropertyName="SelectedValue" />
                                        </SelectParameters>
                                    </asp:SqlDataSource>
                                    <td style="width: 250px">
                                        <asp:Label ID="Label6" runat="server" Text="級職：" Width="50px" ForeColor="Black" CssClass="form"></asp:Label>
                                        <asp:Label ID="Lab_PATITLE" runat="server" ForeColor="Black" Width="150px" CssClass="form"></asp:Label>
                                    </td>
                            </tr>
                        </table>
                    </fieldset>
                    <table border="0" style="width: 750px; height: 57px; color: Red">
                        <tr>
                            <td style="width: 120px; height: 23px;">
                                <asp:Label ID="Label7" runat="server" ForeColor="Black" Text="申請時間：" CssClass="form"></asp:Label>
                            </td>
                            <td style="width: 550px; height: 23px;">
                                <asp:Label ID="lblApplyTime" runat="server" ForeColor="Black" CssClass="form" Width="310px"></asp:Label>
                                <asp:HiddenField ID="hdApplyTime" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 120px">
                                *
                                <asp:Label ID="Label9" runat="server" ForeColor="Black" Text="電話：" CssClass="form"></asp:Label>
                            </td>
                            <td style="height: 26px; width: 550px;">
                                <asp:TextBox ID="txtPhone" runat="server" MaxLength="25" Width="160px"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="style1">
                                *
                                <asp:Label ID="Label12" runat="server" ForeColor="Black" Text="請修地點：" CssClass="form"></asp:Label>
                            </td>
                            <td class="style2">
                                <asp:DropDownList ID="ddlBuilding" runat="server" DataSourceID="SqlDataSource4" DataTextField="Kind_Name"
                                    DataValueField="Kind_Num" AutoPostBack="True">
                                </asp:DropDownList>
                                <asp:SqlDataSource ID="SqlDataSource4" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"                                                                        
                                    SelectCommand="">
                                    <SelectParameters>
                                        <asp:Parameter DefaultValue="12" Name="Kind_Num" Type="Int32" />
                                    </SelectParameters>
                                </asp:SqlDataSource>
                                <asp:DropDownList ID="ddlLevel" runat="server" DataSourceID="SqlDataSource5" DataTextField="State_Name"
                                    DataValueField="State_Num">
                                </asp:DropDownList>
                                <asp:SqlDataSource ID="SqlDataSource5" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                                    
                                    
                                    SelectCommand="SELECT * FROM [SYSKIND] WHERE ([Kind_Num] = @Kind_Num) ORDER BY [State_order]">
                                    <SelectParameters>
                                        <asp:ControlParameter ControlID="ddlBuilding" Name="Kind_Num" 
                                            PropertyName="SelectedValue" />
                                    </SelectParameters>
                                </asp:SqlDataSource>
                                <asp:TextBox ID="txtRoom" runat="server"></asp:TextBox>
                                <asp:Label ID="Label29" runat="server" ForeColor="Black" Text="室"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td style="height: 26px; width: 120px;">
                                *
                                <asp:Label ID="Label8" runat="server" ForeColor="Black" Text="種類：" CssClass="form"></asp:Label>
                            </td>
                            <td style="height: 26px; width: 550px;">
                                <asp:DropDownList ID="ddlRepairMainKind" runat="server" AutoPostBack="True" DataSourceID="SqlDataSource2"
                                    DataTextField="Kind_Name" DataValueField="Kind_Num">
                                </asp:DropDownList>
                                <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                                    SelectCommand="">
                                    <SelectParameters>
                                        <asp:Parameter DefaultValue="7" Name="Kind_Num" Type="Int32" />
                                    </SelectParameters>
                                </asp:SqlDataSource>
                                <asp:DropDownList ID="ddlProblemKind" runat="server" DataSourceID="SqlDataSource3"
                                    DataTextField="State_Name" DataValueField="Kind_SysID">
                                </asp:DropDownList>
                                <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                                    SelectCommand="SELECT * FROM [SYSKIND] WHERE ([Kind_Num] = @Kind_Num)">
                                    <SelectParameters>
                                        <asp:ControlParameter ControlID="ddlRepairMainKind" Name="Kind_Num" PropertyName="SelectedValue"
                                            Type="Int32" />
                                    </SelectParameters>
                                </asp:SqlDataSource>
                                <asp:Panel ID="pnlNetWorkProb" runat="server" Visible="False">
                                    <asp:Label ID="Label30" runat="server" ForeColor="Black" 
    Text="IP位址：" CssClass="form"></asp:Label>
                                    <asp:TextBox ID="txtIPAddr" runat="server" MaxLength="15" Width="96px" 
                                        CssClass="form"></asp:TextBox>
                                    <asp:Label ID="Label31" runat="server" ForeColor="Black" Text="MAC位址：" 
                                        CssClass="form"></asp:Label>
                                    <asp:TextBox ID="txtMACAddr" runat="server" CssClass="form" MaxLength="16"></asp:TextBox>
                                    <asp:Label ID="Label32" runat="server" ForeColor="Black" Text="插座編號：" 
                                        CssClass="form"></asp:Label>
                                    <asp:TextBox ID="txtPlugNo" runat="server" Width="43px" CssClass="form"></asp:TextBox>
                                </asp:Panel>
                            </td>
                        </tr>
                        <tr>
                            <td style="height: 26px; width: 120px;">
                                *
                                <asp:Label ID="Label10" runat="server" ForeColor="Black" Text="請修事項：" CssClass="form"></asp:Label>
                            </td>
                            <td style="height: 26px; width: 550px;">
                                <asp:TextBox ID="txtDESCRIPTION" runat="server" MaxLength="100" Width="500px" Height="80px"
                                    TextMode="MultiLine"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td style="height: 26px; width: 120px;">
                                &nbsp;&nbsp;&nbsp;
                                <asp:Label ID="Label33" runat="server" ForeColor="Black" Text="附件上傳：" 
                                    CssClass="form"></asp:Label>
                            </td>
                            <td style="height: 26px; width: 550px;">
                                
                                <asp:Panel ID="pnlApplyUploadFile" runat="server">
                                    <asp:FileUpload ID="FileUpload1" runat="server" />
                                    <asp:Label ID="Label34" runat="server" CssClass="form" ForeColor="Red" 
                                        Text="附件大小不可超過2MB"></asp:Label>                                                                     
                                </asp:Panel>
                                   <asp:Panel ID="pnlViewUploadFile" runat="server" Visible="False">
                                       <asp:HyperLink ID="lnkUploadFile" runat="server">HyperLink</asp:HyperLink>
                                    </asp:Panel>
                            </td>
                        </tr>
                        <tr <% =Show(read_only,flowAdmin,"FixMan") %>>
                            <td style="height: 26px; width: 120px;">
                                *
                                <asp:Label ID="Label11" runat="server" ForeColor="Black" Text="維修人員：" 
                                    CssClass="form"></asp:Label>
                            </td>
                            <td style="height: 26px; width: 550px;">
                                <div style="display: none;">
                                    <img id="calImg" src="../../Image/calendar.gif" alt="選擇日期" />
                                </div>
                                <asp:DropDownList ID="ddlRepairMan" runat="server" Enabled="False" 
                                    DataSourceID="SqlDataSource6" DataTextField="emp_chinese_name" 
                                    DataValueField="employee_id" Visible="False">
                                    
                                </asp:DropDownList>
                                <asp:Label ID="lblRepairMan" runat="server" ForeColor="Black" 
                                    CssClass="form"></asp:Label>
                                <asp:HiddenField ID="hdnRepairMan" runat="server" />
                                <asp:SqlDataSource ID="SqlDataSource6" runat="server" 
                                    ConnectionString="<%$ ConnectionStrings:ConnectionString %>" 
                                    
                                    SelectCommand="select a.object_uid,a.object_name,b.employee_id,b.object_num,c.emp_chinese_name from systemobj a  left join systemobjuse b on a.object_uid=b.object_uid left join employee c on b.employee_id = c.employee_id where a.object_name='資訊維修單位'">
                                </asp:SqlDataSource>
                            </td>
                        </tr>
                        <tr <% =Show(read_only,flowAdmin,"CallTime") %>>
                            <td style="height: 26px; width: 120px;">
                                *
                                <asp:Label ID="Label35" runat="server" ForeColor="Black" Text="叫修時間：" 
                                    CssClass="form"></asp:Label>
                            </td>
                            <td>
                <asp:TextBox ID="txtCallDate" runat="server" OnKeyDown="return false" Width="100px"></asp:TextBox>
                                <asp:DropDownList ID="ddlCallTimeHour" runat="server">
                                    <asp:ListItem>1</asp:ListItem>
                                    <asp:ListItem>2</asp:ListItem>
                                    <asp:ListItem>3</asp:ListItem>
                                    <asp:ListItem>4</asp:ListItem>
                                    <asp:ListItem>5</asp:ListItem>
                                    <asp:ListItem>6</asp:ListItem>
                                    <asp:ListItem>7</asp:ListItem>
                                    <asp:ListItem>8</asp:ListItem>
                                    <asp:ListItem>9</asp:ListItem>
                                    <asp:ListItem>10</asp:ListItem>
                                    <asp:ListItem>11</asp:ListItem>
                                    <asp:ListItem>12</asp:ListItem>
                                    <asp:ListItem>13</asp:ListItem>
                                    <asp:ListItem>14</asp:ListItem>
                                    <asp:ListItem>15</asp:ListItem>
                                    <asp:ListItem>16</asp:ListItem>
                                    <asp:ListItem>17</asp:ListItem>
                                    <asp:ListItem>18</asp:ListItem>
                                    <asp:ListItem>19</asp:ListItem>
                                    <asp:ListItem>20</asp:ListItem>
                                    <asp:ListItem>21</asp:ListItem>
                                    <asp:ListItem>22</asp:ListItem>
                                    <asp:ListItem>23</asp:ListItem>
                                    <asp:ListItem>24</asp:ListItem>
                                </asp:DropDownList>
                                <asp:Label ID="Label39" runat="server" ForeColor="Black" Text="時"></asp:Label>
                                <asp:DropDownList ID="ddlCallTimeMin" runat="server">
                                    <asp:ListItem>00</asp:ListItem>
                                    <asp:ListItem>01</asp:ListItem>
                                    <asp:ListItem>02</asp:ListItem>
                                    <asp:ListItem>03</asp:ListItem>
                                    <asp:ListItem>04</asp:ListItem>
                                    <asp:ListItem>05</asp:ListItem>
                                    <asp:ListItem>06</asp:ListItem>
                                    <asp:ListItem>07</asp:ListItem>
                                    <asp:ListItem>08</asp:ListItem>
                                    <asp:ListItem>09</asp:ListItem>
                                    <asp:ListItem>10</asp:ListItem>
                                    <asp:ListItem>11</asp:ListItem>
                                    <asp:ListItem>12</asp:ListItem>
                                    <asp:ListItem>13</asp:ListItem>
                                    <asp:ListItem>14</asp:ListItem>
                                    <asp:ListItem>15</asp:ListItem>
                                    <asp:ListItem>16</asp:ListItem>
                                    <asp:ListItem>17</asp:ListItem>
                                    <asp:ListItem>18</asp:ListItem>
                                    <asp:ListItem>19</asp:ListItem>
                                    <asp:ListItem>20</asp:ListItem>
                                    <asp:ListItem>21</asp:ListItem>
                                    <asp:ListItem>22</asp:ListItem>
                                    <asp:ListItem>23</asp:ListItem>
                                    <asp:ListItem>24</asp:ListItem>
                                    <asp:ListItem>25</asp:ListItem>
                                    <asp:ListItem>26</asp:ListItem>
                                    <asp:ListItem>27</asp:ListItem>
                                    <asp:ListItem>28</asp:ListItem>
                                    <asp:ListItem>29</asp:ListItem>
                                    <asp:ListItem>30</asp:ListItem>
                                    <asp:ListItem>31</asp:ListItem>
                                    <asp:ListItem>32</asp:ListItem>
                                    <asp:ListItem>33</asp:ListItem>
                                    <asp:ListItem>34</asp:ListItem>
                                    <asp:ListItem>35</asp:ListItem>
                                    <asp:ListItem>36</asp:ListItem>
                                    <asp:ListItem>37</asp:ListItem>
                                    <asp:ListItem>38</asp:ListItem>
                                    <asp:ListItem>39</asp:ListItem>
                                    <asp:ListItem>40</asp:ListItem>
                                    <asp:ListItem>41</asp:ListItem>
                                    <asp:ListItem>42</asp:ListItem>
                                    <asp:ListItem>43</asp:ListItem>
                                    <asp:ListItem>44</asp:ListItem>
                                    <asp:ListItem>45</asp:ListItem>
                                    <asp:ListItem>46</asp:ListItem>
                                    <asp:ListItem>47</asp:ListItem>
                                    <asp:ListItem>48</asp:ListItem>
                                    <asp:ListItem>49</asp:ListItem>
                                    <asp:ListItem>50</asp:ListItem>
                                    <asp:ListItem>51</asp:ListItem>
                                    <asp:ListItem>52</asp:ListItem>
                                    <asp:ListItem>53</asp:ListItem>
                                    <asp:ListItem>54</asp:ListItem>
                                    <asp:ListItem>55</asp:ListItem>
                                    <asp:ListItem>56</asp:ListItem>
                                    <asp:ListItem>57</asp:ListItem>
                                    <asp:ListItem>58</asp:ListItem>
                                    <asp:ListItem>59</asp:ListItem>                                    
                                </asp:DropDownList>
                                <asp:Label ID="Label40" runat="server" ForeColor="Black" Text="分"></asp:Label>
                            </td>
                        </tr>
                        <tr <% =Show(read_only,flowAdmin,"ArriveTime") %>>
                            <td style="height: 26px; width: 120px;">
                                *
                                <asp:Label ID="Label36" runat="server" ForeColor="Black" Text="到修時間：" 
                                    CssClass="form"></asp:Label>
                            </td>
                            <td>
                <asp:TextBox ID="txtArriveDate" runat="server" OnKeyDown="return false" Width="100px"></asp:TextBox>
                <div style="display: none;">
                    <img id="Img1" src="../../Image/calendar.gif" alt="選擇日期" />
                </div>
                                <asp:DropDownList ID="ddlArriveTimeHour" runat="server">
                                    <asp:ListItem>1</asp:ListItem>
                                    <asp:ListItem>2</asp:ListItem>
                                    <asp:ListItem>3</asp:ListItem>
                                    <asp:ListItem>4</asp:ListItem>
                                    <asp:ListItem>5</asp:ListItem>
                                    <asp:ListItem>6</asp:ListItem>
                                    <asp:ListItem>7</asp:ListItem>
                                    <asp:ListItem>8</asp:ListItem>
                                    <asp:ListItem>9</asp:ListItem>
                                    <asp:ListItem>10</asp:ListItem>
                                    <asp:ListItem>11</asp:ListItem>
                                    <asp:ListItem>12</asp:ListItem>
                                    <asp:ListItem>13</asp:ListItem>
                                    <asp:ListItem>14</asp:ListItem>
                                    <asp:ListItem>15</asp:ListItem>
                                    <asp:ListItem>16</asp:ListItem>
                                    <asp:ListItem>17</asp:ListItem>
                                    <asp:ListItem>18</asp:ListItem>
                                    <asp:ListItem>19</asp:ListItem>
                                    <asp:ListItem>20</asp:ListItem>
                                    <asp:ListItem>21</asp:ListItem>
                                    <asp:ListItem>22</asp:ListItem>
                                    <asp:ListItem>23</asp:ListItem>
                                    <asp:ListItem>24</asp:ListItem>
                                </asp:DropDownList>
                                <asp:Label ID="Label41" runat="server" ForeColor="Black" Text="時"></asp:Label>
                                <asp:DropDownList ID="ddlArriveTimeMin" runat="server">
                                    <asp:ListItem>00</asp:ListItem>
                                    <asp:ListItem>01</asp:ListItem>
                                    <asp:ListItem>02</asp:ListItem>
                                    <asp:ListItem>03</asp:ListItem>
                                    <asp:ListItem>04</asp:ListItem>
                                    <asp:ListItem>05</asp:ListItem>
                                    <asp:ListItem>06</asp:ListItem>
                                    <asp:ListItem>07</asp:ListItem>
                                    <asp:ListItem>08</asp:ListItem>
                                    <asp:ListItem>09</asp:ListItem>
                                    <asp:ListItem>10</asp:ListItem>
                                    <asp:ListItem>11</asp:ListItem>
                                    <asp:ListItem>12</asp:ListItem>
                                    <asp:ListItem>13</asp:ListItem>
                                    <asp:ListItem>14</asp:ListItem>
                                    <asp:ListItem>15</asp:ListItem>
                                    <asp:ListItem>16</asp:ListItem>
                                    <asp:ListItem>17</asp:ListItem>
                                    <asp:ListItem>18</asp:ListItem>
                                    <asp:ListItem>19</asp:ListItem>
                                    <asp:ListItem>20</asp:ListItem>
                                    <asp:ListItem>21</asp:ListItem>
                                    <asp:ListItem>22</asp:ListItem>
                                    <asp:ListItem>23</asp:ListItem>
                                    <asp:ListItem>24</asp:ListItem>
                                    <asp:ListItem>25</asp:ListItem>
                                    <asp:ListItem>26</asp:ListItem>
                                    <asp:ListItem>27</asp:ListItem>
                                    <asp:ListItem>28</asp:ListItem>
                                    <asp:ListItem>29</asp:ListItem>
                                    <asp:ListItem>30</asp:ListItem>
                                    <asp:ListItem>31</asp:ListItem>
                                    <asp:ListItem>32</asp:ListItem>
                                    <asp:ListItem>33</asp:ListItem>
                                    <asp:ListItem>34</asp:ListItem>
                                    <asp:ListItem>35</asp:ListItem>
                                    <asp:ListItem>36</asp:ListItem>
                                    <asp:ListItem>37</asp:ListItem>
                                    <asp:ListItem>38</asp:ListItem>
                                    <asp:ListItem>39</asp:ListItem>
                                    <asp:ListItem>40</asp:ListItem>
                                    <asp:ListItem>41</asp:ListItem>
                                    <asp:ListItem>42</asp:ListItem>
                                    <asp:ListItem>43</asp:ListItem>
                                    <asp:ListItem>44</asp:ListItem>
                                    <asp:ListItem>45</asp:ListItem>
                                    <asp:ListItem>46</asp:ListItem>
                                    <asp:ListItem>47</asp:ListItem>
                                    <asp:ListItem>48</asp:ListItem>
                                    <asp:ListItem>49</asp:ListItem>
                                    <asp:ListItem>50</asp:ListItem>
                                    <asp:ListItem>51</asp:ListItem>
                                    <asp:ListItem>52</asp:ListItem>
                                    <asp:ListItem>53</asp:ListItem>
                                    <asp:ListItem>54</asp:ListItem>
                                    <asp:ListItem>55</asp:ListItem>
                                    <asp:ListItem>56</asp:ListItem>
                                    <asp:ListItem>57</asp:ListItem>
                                    <asp:ListItem>58</asp:ListItem>
                                    <asp:ListItem>59</asp:ListItem>
                                </asp:DropDownList>
                                <asp:Label ID="Label42" runat="server" ForeColor="Black" Text="分"></asp:Label>
                            </td>
                        </tr>
                        <tr <% =Show(read_only,flowAdmin,"FixRecord") %>>
                            <td style="height: 26px; width: 120px;">
                                &nbsp;&nbsp;&nbsp;
                                <asp:Label ID="Label38" runat="server" ForeColor="Black" Text="維修紀錄：" 
                                    CssClass="form"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="txtFIXRECORD" runat="server" MaxLength="100" Width="500px" Height="80px"
                                    TextMode="MultiLine"></asp:TextBox>
                            </td>
                        </tr>
                        <tr <% =Show(read_only,flowAdmin,"FixDate") %>>
                            <td style="height: 26px; width: 120px;">
                                *
                                <asp:Label ID="Label37" runat="server" ForeColor="Black" Text="完修時間：" 
                                    CssClass="form"></asp:Label>
                            </td>
                            <td>
                <asp:TextBox ID="txtFinalDate" runat="server" OnKeyDown="return false" Width="100px"></asp:TextBox>
                                <asp:DropDownList ID="ddlFinalTimeHour" runat="server">
                                    <asp:ListItem>1</asp:ListItem>
                                    <asp:ListItem>2</asp:ListItem>
                                    <asp:ListItem>3</asp:ListItem>
                                    <asp:ListItem>4</asp:ListItem>
                                    <asp:ListItem>5</asp:ListItem>
                                    <asp:ListItem>6</asp:ListItem>
                                    <asp:ListItem>7</asp:ListItem>
                                    <asp:ListItem>8</asp:ListItem>
                                    <asp:ListItem>9</asp:ListItem>
                                    <asp:ListItem>10</asp:ListItem>
                                    <asp:ListItem>11</asp:ListItem>
                                    <asp:ListItem>12</asp:ListItem>
                                    <asp:ListItem>13</asp:ListItem>
                                    <asp:ListItem>14</asp:ListItem>
                                    <asp:ListItem>15</asp:ListItem>
                                    <asp:ListItem>16</asp:ListItem>
                                    <asp:ListItem>17</asp:ListItem>
                                    <asp:ListItem>18</asp:ListItem>
                                    <asp:ListItem>19</asp:ListItem>
                                    <asp:ListItem>20</asp:ListItem>
                                    <asp:ListItem>21</asp:ListItem>
                                    <asp:ListItem>22</asp:ListItem>
                                    <asp:ListItem>23</asp:ListItem>
                                    <asp:ListItem>24</asp:ListItem>
                                </asp:DropDownList>
                                <asp:Label ID="Label43" runat="server" ForeColor="Black" Text="時"></asp:Label>
                                <asp:DropDownList ID="ddlFinalTimeMin" runat="server">
                                    <asp:ListItem>00</asp:ListItem>
                                    <asp:ListItem>01</asp:ListItem>
                                    <asp:ListItem>02</asp:ListItem>
                                    <asp:ListItem>03</asp:ListItem>
                                    <asp:ListItem>04</asp:ListItem>
                                    <asp:ListItem>05</asp:ListItem>
                                    <asp:ListItem>06</asp:ListItem>
                                    <asp:ListItem>07</asp:ListItem>
                                    <asp:ListItem>08</asp:ListItem>
                                    <asp:ListItem>09</asp:ListItem>
                                    <asp:ListItem>10</asp:ListItem>
                                    <asp:ListItem>11</asp:ListItem>
                                    <asp:ListItem>12</asp:ListItem>
                                    <asp:ListItem>13</asp:ListItem>
                                    <asp:ListItem>14</asp:ListItem>
                                    <asp:ListItem>15</asp:ListItem>
                                    <asp:ListItem>16</asp:ListItem>
                                    <asp:ListItem>17</asp:ListItem>
                                    <asp:ListItem>18</asp:ListItem>
                                    <asp:ListItem>19</asp:ListItem>
                                    <asp:ListItem>20</asp:ListItem>
                                    <asp:ListItem>21</asp:ListItem>
                                    <asp:ListItem>22</asp:ListItem>
                                    <asp:ListItem>23</asp:ListItem>
                                    <asp:ListItem>24</asp:ListItem>
                                    <asp:ListItem>25</asp:ListItem>
                                    <asp:ListItem>26</asp:ListItem>
                                    <asp:ListItem>27</asp:ListItem>
                                    <asp:ListItem>28</asp:ListItem>
                                    <asp:ListItem>29</asp:ListItem>
                                    <asp:ListItem>30</asp:ListItem>
                                    <asp:ListItem>31</asp:ListItem>
                                    <asp:ListItem>32</asp:ListItem>
                                    <asp:ListItem>33</asp:ListItem>
                                    <asp:ListItem>34</asp:ListItem>
                                    <asp:ListItem>35</asp:ListItem>
                                    <asp:ListItem>36</asp:ListItem>
                                    <asp:ListItem>37</asp:ListItem>
                                    <asp:ListItem>38</asp:ListItem>
                                    <asp:ListItem>39</asp:ListItem>
                                    <asp:ListItem>40</asp:ListItem>
                                    <asp:ListItem>41</asp:ListItem>
                                    <asp:ListItem>42</asp:ListItem>
                                    <asp:ListItem>43</asp:ListItem>
                                    <asp:ListItem>44</asp:ListItem>
                                    <asp:ListItem>45</asp:ListItem>
                                    <asp:ListItem>46</asp:ListItem>
                                    <asp:ListItem>47</asp:ListItem>
                                    <asp:ListItem>48</asp:ListItem>
                                    <asp:ListItem>49</asp:ListItem>
                                    <asp:ListItem>50</asp:ListItem>
                                    <asp:ListItem>51</asp:ListItem>
                                    <asp:ListItem>52</asp:ListItem>
                                    <asp:ListItem>53</asp:ListItem>
                                    <asp:ListItem>54</asp:ListItem>
                                    <asp:ListItem>55</asp:ListItem>
                                    <asp:ListItem>56</asp:ListItem>
                                    <asp:ListItem>57</asp:ListItem>
                                    <asp:ListItem>58</asp:ListItem>
                                    <asp:ListItem>59</asp:ListItem>
                                </asp:DropDownList>
                                <asp:Label ID="Label44" runat="server" ForeColor="Black" Text="分"></asp:Label>
                            </td>
                        </tr>
                        <tr <% =Show(read_only,flowAdmin,"FixUploadFinish") %>>
                            <td style="height: 26px; width: 120px;">
                                &nbsp;&nbsp;&nbsp;
                                <asp:Label ID="Label45" runat="server" ForeColor="Black" Text="完修附件上傳：" 
                                    CssClass="form"></asp:Label>
                            </td>
                            <td style="height: 26px; width: 550px;">
                                <asp:Panel ID="pnlFinishUploadFile" runat="server">
                                    <asp:FileUpload ID="FileUpload2" runat="server"/>
                                    <asp:Label ID="Label46" runat="server" CssClass="form" ForeColor="Red" 
                                        Text="附件大小不可超過2MB"></asp:Label>                                                                     
                                </asp:Panel>
                                   <asp:Panel ID="pnlViewFinishUploadFile" runat="server" Visible="False">
                                       <asp:HyperLink ID="lnkFinishUploadFile" runat="server">HyperLink</asp:HyperLink>
                                    </asp:Panel>
                            </td>
                        </tr>
                        <tr <% =Show(read_only,flowAdmin,"Comment") %>>
                            <td style="height: 26px; width: 120px;">
                                &nbsp;<asp:Label ID="Label28" runat="server" ForeColor="Black" Text="批核意見：" CssClass="form"></asp:Label><br />
                            </td>
                            <td>
                                <asp:TextBox ID="txtcomment" runat="server" Height="59px" MaxLength="255" Rows="3"
                                    TextMode="MultiLine" Width="529px"></asp:TextBox>
                                <asp:Button ID="But_PHRASE" runat="server" Text="批核片語" />
                            </td>
                        </tr>
                    </table>
                    <table border="0" style="width: 750px; height: 57px">
                        <tr>
                            <td align="center">
                                <asp:Button ID="But_exe" runat="server" Text="送件" />
                                <asp:Button ID="But_prt" runat="server" Text="列印" Visible="False" />&nbsp;<asp:Button ID="backBtn"
                                    runat="server" Text="駁回" />&nbsp;<asp:Button ID="tranBtn" runat="server" 
                                    Text="呈轉" Visible="False" />
                                <asp:Button ID="btnReturn"
                                    runat="server" Text="退回" Visible="False" />
                                <asp:Button ID="btnReAppointment"
                                    runat="server" Text="重分" Visible="False" />
                            </td>
                        </tr>
                    </table>
                    <uc1:FlowRoute ID="FlowRoute1" runat="server" />
                </td>
            </tr>
        </table>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
            SelectCommand=""></asp:SqlDataSource>
    </div>
    <div id="Div_grid10" runat="server" style="position: absolute; z-index: 3; background-color: white;
        width: 300pt; height: 80pt; left: 700px; top: 1047px; display: block;" 
        visible="false">
        <asp:GridView ID="GridView10" runat="server" CssClass="form" Width="100%" Height="50px"
            DataSourceID="SqlDataSource10" PageSize="5" AutoGenerateColumns="False" AllowPaging="True"
            BorderColor="Lime" BorderWidth="2px">
            <Columns>
                <asp:BoundField DataField="comment" HeaderText="批核片語">
                    <HeaderStyle HorizontalAlign="Center" Width="90%" BackColor="#80FF80" CssClass="form" />
                </asp:BoundField>
                <asp:BoundField DataField="Phrase_num" HeaderText="Phrase_num" Visible="False">
                    <ItemStyle Wrap="False" />
                </asp:BoundField>
                <asp:CommandField ShowSelectButton="True">
                    <HeaderStyle HorizontalAlign="Center" Width="10%" />
                </asp:CommandField>
            </Columns>
            <RowStyle Height="10px" />
        </asp:GridView>
        &nbsp;
        <asp:Button ID="Btn_PHclose" runat="server" Text="關閉" Width="389px" />
        <asp:SqlDataSource ID="SqlDataSource10" runat="server" SelectCommand="SELECT Phrase_num, employee_id, comment FROM PHRASE WHERE [employee_id] = @employee_id ORDER BY Phrase_num"
            ConnectionString="<%$ ConnectionStrings:ConnectionString %>">
            <SelectParameters>
                <asp:SessionParameter Name="employee_id" SessionField="user_id" Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
    </div>
    </form>
    <script language="javascript">
    
function window_onload() {
  <%if do_sql.G_errmsg <>"" then %>  
    alert('<%= do_sql.G_errmsg %>');
  <%end if %>
  
  //顯示報表檔
  <%if pfilename <>"" then %>  
    window.location.replace('<%= pfilename %>');
  <%end if %>
  
  //產生列印檔過程有錯誤
  <%if ErrMsg <>"" then %>  
    alert('<%= ErrMsg %>');
  <%end if %>
}        
    </script>
</body>
</html>
