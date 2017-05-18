'==================================================================================================================================================================================
'   Author      : Andy Lin
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Created at 2010/07/23
'   Description : First Version for Development
'==================================================================================================================================================================================
Imports Microsoft.Reporting.WebForms

Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary

Imports WebUtilities.Pages

Partial Class Source_01_MOA01010
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String
    Private report_parameters As SortedList = Nothing

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        user_id = Session("user_id")
        org_uid = Session("ORG_UID")

        'Session被清空回首頁
        If user_id = "" Or org_uid = "" Then

            Response.Write("<script language='javascript'>")
            Response.Write("alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write("window.parent.location='../../index.aspx';")
            Response.Write("</script>")

        Else

            '判斷登入者權限
            Dim LoginCheck As New C_Public

            If LoginCheck.LoginCheck(user_id, "MOA01009") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA01010.aspx")
                Response.End()
            End If

        End If

        If Not Me.IsPostBack Then

            Dim htInformation As Hashtable = Me.ReportInformation_Deserialization()

            If Not htInformation Is Nothing Then

                Me.SqlDataSourceController_Configuration(htInformation)

                Dim ds_01010 As DS01010 = Nothing

                Using data_source_table As DataTable = DirectCast(sdsSignRecords.Select(New DataSourceSelectArguments), DataView).Table

                    ds_01010 = New DS01010()

                    For Each data_row As DataRow In data_source_table.Rows
                        ds_01010.SignRecords.ImportRow(data_row)
                    Next

                End Using

                If Not ds_01010 Is Nothing Then

                    Dim over_time_baseline As TimeSpan = Nothing

                    If TimeSpan.TryParse(htInformation("Over_Time_Baseline").ToString(), over_time_baseline) Then

                        htInformation.Remove("Over_Time_Baseline")

                        Dim result As List(Of Integer) = Nothing

                        For Each data_row As DataRow In ds_01010.SignRecords.Rows

                            MOA01009.HoursAndSubTotal_Calculations _
                                     ( _
                                             data_row.Item("Out_Time_nvc"), _
                                             Me.report_parameters.Item("HourLimit"), _
                                             Me.report_parameters.Item("OverTime_Price"), _
                                             data_row.Item("NoonHour"), _
                                             over_time_baseline, _
                                             result _
                                     )

                            If Not result.Item(0).Equals(Integer.MinValue) Then
                                data_row("Hours") = result.Item(0).ToString()
                            End If

                            If Not result.Item(1).Equals(Integer.MinValue) Then
                                data_row("SubTotal") = result.Item(1).ToString()
                            End If

                        Next

                        If TypeOf result Is List(Of Integer) Then

                            result.Clear()
                            result = Nothing

                        End If

                    End If

                    over_time_baseline = Nothing

                End If

                Me.MicrosoftReport_Generation(htInformation, ds_01010.SignRecords)

                ClientScript.RegisterStartupScript _
                             ( _
                                     Me.GetType(), _
                                    "RSClientPrint_ActiveX_Setup_Error_Fix", _
                                     String.Format("document.getElementById('{0}').ClientController.ActionHandler('Refresh', null);", rvSignRecords.ID), _
                                     True _
                             )

            End If

        End If

    End Sub

    Private Function ReportInformation_Deserialization() As Hashtable

        Dim htResult As Hashtable = Nothing

        Dim file_path As String = Server.MapPath("~/Drs/") & _
                                  String.Format("SERIAL01009{0:yyyyMMddHHmmssffff}.dat", Request.QueryString("sn"))

        Using file_reader As New StreamReader(file_path, Encoding.GetEncoding("BIG5"))

            Dim formatter As New BinaryFormatter

            Try
                htResult = DirectCast(formatter.Deserialize(file_reader.BaseStream), Hashtable)

            Catch ex As Exception
                Throw

            Finally

                file_reader.Close()

                If File.Exists(file_path) Then
                    File.Delete(file_path)
                End If

            End Try

        End Using

        Return htResult

    End Function

    Private Sub SqlDataSourceController_Configuration(ByRef params_collection As Hashtable)

        Dim setting_value(1) As String

        For Each element As DictionaryEntry In params_collection.Clone()

            If element.Key.Equals("SQL_Statements") Then

                sdsSignRecords.SelectCommand = element.Value.ToString()
                params_collection.Remove(element.Key)

            ElseIf element.Key.ToString().StartsWith("@") Then

                setting_value = element.Value.ToString().Split("|")

                sdsSignRecords.SelectParameters.Add _
                                                ( _
                                                    New Parameter _
                                                        ( _
                                                            element.Key.ToString().TrimStart("@"), _
                                                            DirectCast([Enum].Parse(GetType(DbType), setting_value(0)), DbType), _
                                                            setting_value(1) _
                                                        ) _
                                                )

                Array.Clear(setting_value, 0, setting_value.Length)

                params_collection.Remove(element.Key)

            End If

        Next

        setting_value = Nothing

    End Sub

    Protected Sub sdsSignRecords_Selecting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceSelectingEventArgs) Handles sdsSignRecords.Selecting

        If e.Command.Parameters.Contains("@employee_id") Then

            Dim command_text As New StringBuilder

            command_text.Append("Select @OverTime_Price = P.OverTime, @HourLimit = P.HourLimit, @Salary_Level = P.MoneyKind, ") _
                        .Append("@Title_Level = E.AD_TITLE, @Orgnization_Name = GP.ORG_NAME From EMPLOYEE As E Left Join ADMINGROUP ") _
                        .Append("As GP On GP.ORG_UID = E.ORG_UID Left Join P_0103 As P On P.employee_id = E.employee_id ") _
                        .Append("Where E.employee_id = @employee_id;")

            e.Command.CommandText += command_text.ToString()

            command_text.Remove(0, command_text.Length)
            command_text = Nothing

            If Not e.Command.Parameters.Contains("@OverTime_Price") Then

                With e.Command.Parameters(e.Command.Parameters.Add(New SqlParameter))
                    .ParameterName = "@OverTime_Price"
                    .DbType = DbType.Int32
                    .Direction = ParameterDirection.Output
                End With

            End If

            If Not e.Command.Parameters.Contains("@HourLimit") Then

                With e.Command.Parameters(e.Command.Parameters.Add(New SqlParameter))
                    .ParameterName = "@HourLimit"
                    .DbType = DbType.Int32
                    .Direction = ParameterDirection.Output
                End With

            End If

            If Not e.Command.Parameters.Contains("@Title_Level") Then

                With e.Command.Parameters(e.Command.Parameters.Add(New SqlParameter))
                    .ParameterName = "@Title_Level"
                    .DbType = DbType.String
                    .Direction = ParameterDirection.Output
                    .Size = 30
                End With

            End If

            If Not e.Command.Parameters.Contains("@Salary_Level") Then

                With e.Command.Parameters(e.Command.Parameters.Add(New SqlParameter))
                    .ParameterName = "@Salary_Level"
                    .DbType = DbType.String
                    .Direction = ParameterDirection.Output
                    .Size = 5
                End With

            End If

            If Not e.Command.Parameters.Contains("@Orgnization_Name") Then

                With e.Command.Parameters(e.Command.Parameters.Add(New SqlParameter))
                    .ParameterName = "@Orgnization_Name"
                    .DbType = DbType.String
                    .Direction = ParameterDirection.Output
                    .Size = 50
                End With

            End If

        Else

            If e.Command.Parameters.Contains("@OverTime_Price") Then
                e.Command.Parameters.Remove(e.Command.Parameters("@OverTime_Price"))
            End If

            If e.Command.Parameters.Contains("@HourLimit") Then
                e.Command.Parameters.Remove(e.Command.Parameters("@HourLimit"))
            End If

            If e.Command.Parameters.Contains("@Title_Level") Then
                e.Command.Parameters.Remove(e.Command.Parameters("@Title_Level"))
            End If

            If e.Command.Parameters.Contains("@Salary_Level") Then
                e.Command.Parameters.Remove(e.Command.Parameters("@Salary_Level"))
            End If

            If e.Command.Parameters.Contains("@Orgnization_Name") Then
                e.Command.Parameters.Remove(e.Command.Parameters("@Orgnization_Name"))
            End If

        End If

    End Sub

    Protected Sub sdsSignRecords_Selected(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles sdsSignRecords.Selected

        Me.report_parameters = New SortedList(4) : With report_parameters
            .Add("OverTime_Price", Nothing)
            .Add("HourLimit", Nothing)
            .Add("Title_Level", String.Empty)
            .Add("Salary_Level", String.Empty)
            .Add("Orgnization_Name", String.Empty)
        End With

        If e.Command.Parameters.Contains("@OverTime_Price") Then
            Me.report_parameters("OverTime_Price") = e.Command.Parameters("@OverTime_Price").Value
        End If

        If e.Command.Parameters.Contains("@HourLimit") Then
            Me.report_parameters("HourLimit") = e.Command.Parameters("@HourLimit").Value
        End If

        If e.Command.Parameters.Contains("@Title_Level") Then
            Me.report_parameters("Title_Level") = e.Command.Parameters("@Title_Level").Value
        End If

        If e.Command.Parameters.Contains("@Salary_Level") Then
            Me.report_parameters("Salary_Level") = e.Command.Parameters("@Salary_Level").Value
        End If

        If e.Command.Parameters.Contains("@Orgnization_Name") Then
            Me.report_parameters("Orgnization_Name") = e.Command.Parameters("@Orgnization_Name").Value
        End If

    End Sub

    Private Sub MicrosoftReport_Generation(ByRef parameters As Hashtable, ByRef data_source As DataTable)

        Dim report_source As New ReportDataSource

        With report_source
            .Name = String.Format("{0}_{1}", data_source.DataSet.DataSetName, data_source.TableName)
            .Value = data_source
        End With

        Me.rvSignRecords.ProcessingMode = ProcessingMode.Local
        Me.rvSignRecords.LocalReport.ReportPath = Server.MapPath("~/Rpt/RPT01010.rdlc")

        Me.rvSignRecords.LocalReport.SetParameters _
                                     ( _
                                                   New ReportParameter() _
                                                   { _
                                                             New ReportParameter("rpCompensatoryLeave", parameters("Compensatory_Leave").ToString()), _
                                                             New ReportParameter("rpReportYear", (Date.Parse(parameters("Report_BeginDate")).Year - 1911).ToString()), _
                                                             New ReportParameter("rpReportMonth", Date.Parse(parameters("Report_BeginDate")).Month.ToString()), _
                                                             New ReportParameter("rpTitle_Level", Me.report_parameters("Title_Level").ToString()), _
                                                             New ReportParameter("rpSalary_Level", Me.report_parameters("Salary_Level").ToString()), _
                                                             New ReportParameter("rpOrgnization_Name", Me.report_parameters("Orgnization_Name").ToString()) _
                                                   } _
                                     )

        If Not String.IsNullOrEmpty(Me.report_parameters("OverTime_Price").ToString()) Then
            Me.rvSignRecords.LocalReport.SetParameters(New ReportParameter() {New ReportParameter("rpOverTime_Price", Me.report_parameters("OverTime_Price").ToString())})
        End If

        If Not String.IsNullOrEmpty(Me.report_parameters("HourLimit").ToString()) Then
            Me.rvSignRecords.LocalReport.SetParameters(New ReportParameter() {New ReportParameter("rpHourLimit", Me.report_parameters("HourLimit").ToString())})
        End If

        Me.report_parameters.Clear()
        Me.report_parameters = Nothing

        parameters.Clear()
        parameters = Nothing

        Me.rvSignRecords.LocalReport.DataSources.Add(report_source)
        Me.rvSignRecords.LocalReport.Refresh()

    End Sub

End Class