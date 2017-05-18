'==================================================================================================================================================================================
'   Author      : Andy Lin
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Create at 2010/06/18
'   Description : First version of development.
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Modified at 2010/07/27
'   Description : 新增 MOA01009 與 MOA01010 共用新增靜態方法。
'==================================================================================================================================================================================
Imports Microsoft.VisualBasic
Imports System.Collections.Generic

Namespace WebUtilities.Functions

    Public Class Utilities

        Public Shared Function SubmitFormGeneration(ByRef parameters As SortedList) As String

            Dim scripts As New StringBuilder
            Dim enumer As IDictionaryEnumerator = parameters.GetEnumerator()

            While (enumer.MoveNext())

                Select Case (enumer.Key.ToString().Trim().ToUpper())

                    Case "ACTION"
                        scripts.Insert(0, String.Format("<html><body><form id=""SubmitForm"" method=""post"" action=""{0}"">", enumer.Value.ToString().Trim()))

                    Case Else
                        scripts.Append("<input type=""hidden"" ") _
                               .AppendFormat("name=""{0}"" ", enumer.Key.ToString().Trim()) _
                               .AppendFormat("value=""{0}"">", enumer.Value.ToString().Trim())

                End Select

            End While

            If (Not scripts.Length.Equals(0)) Then

                scripts.Append("</form><script language='javascript'>SubmitForm.submit();") _
                       .Append("</script></body></html>")

            End If

            parameters.Clear()
            parameters = Nothing

            Return scripts.ToString()

        End Function

    End Class

End Namespace

Namespace WebUtilities.Pages

    Public Class MOA01008

        ''' <summary>
        ''' 處單位人事員的群組代碼。
        ''' </summary>
        ''' <remarks></remarks>
        Public Const DivisionPersonnelAdministrator As String = "1654314ON8"

        ''' <summary>
        ''' 單位薪俸管理員的群組代碼。
        ''' </summary>
        ''' <remarks></remarks>
        Public Const UnitSalariesAdministrator As String = "B99A77QB5H"

        ''' <summary>
        ''' 公務人員薪俸管理員的群組代碼。
        ''' </summary>
        ''' <remarks></remarks>
        Public Const PublicServantSalariesAdministrator As String = "R42NZOWYF5"

    End Class

    Public Class MOA01009

        Public Shared ReadOnly Property Special_AD_Title() As String

            Get
                Return "[簡薦委]"
            End Get

        End Property

        Public Shared ReadOnly Property Out_Time_NVC_End() As String

            Get
                Return "06:00:00"
            End Get

        End Property

        Public Shared Sub HoursAndSubTotal_Calculations(ByRef out_time_nvc As Object, ByRef hour_limit As Object, ByRef over_time As Object, ByRef noon_hour As Object, _
                                                        ByRef over_time_baseline As TimeSpan, ByRef result As List(Of Integer))

            If result Is Nothing Then

                result = New List(Of Integer)(2)

                With result
                    .Add(Nothing)
                    .Add(Nothing)
                End With

            End If

            For index As Integer = 0 To (result.Count - 1)
                result.Item(index) = Integer.MinValue
            Next

            If Not ( _
                        String.IsNullOrEmpty(out_time_nvc.ToString()) _
                        Or _
                        String.IsNullOrEmpty(hour_limit.ToString()) _
                        Or _
                        String.IsNullOrEmpty(over_time.ToString()) _
                   ) Then

                Dim hour_limitation As Integer = Nothing
                Dim over_time_hourly As Integer = Nothing

                If Int32.TryParse(hour_limit.ToString(), hour_limitation) And Int32.TryParse(over_time.ToString(), over_time_hourly) Then

                    Dim over_time_span As TimeSpan = Nothing

                    If TimeSpan.TryParse(out_time_nvc.ToString(), over_time_span) Then

                        Select Case over_time_span.CompareTo(TimeSpan.Parse("12:00:00"))

                            Case 1

                                result.Item(0) = Convert.ToInt32 _
                                                         ( _
                                                                 Math.Floor _
                                                                      ( _
                                                                            DateTime.MinValue.Add(over_time_span) _
                                                                                             .Subtract _
                                                                                              ( _
                                                                                                       DateTime.MinValue.Add(over_time_baseline) _
                                                                                              ) _
                                                                                             .TotalHours _
                                                                      ) _
                                                         )

                            Case 0, -1

                                If over_time_span < TimeSpan.Parse(MOA01009.Out_Time_NVC_End) Then

                                    result.Item(0) = Convert.ToInt32 _
                                                             ( _
                                                                     Math.Floor _
                                                                          ( _
                                                                                (24 - over_time_baseline.TotalHours) _
                                                                                 + _
                                                                                 over_time_span.TotalHours _
                                                                          ) _
                                                             )

                                End If

                        End Select

                        If result.Item(0) > 0 Then

                            Dim noon_hour_value As Integer = Nothing

                            If Integer.TryParse(noon_hour.ToString(), noon_hour_value) Then
                                result.Item(0) += noon_hour_value
                            End If

                            noon_hour_value = Nothing

                            If hour_limitation.Equals(20) And result.Item(0) > 4 Then
                                result.Item(0) = 4
                            End If

                            result.Item(1) = result.Item(0) * over_time_hourly

                        Else

                            result.Item(0) = Integer.MinValue
                            result.Item(1) = Integer.MinValue

                        End If

                    End If

                    over_time_span = Nothing

                End If

                hour_limitation = Nothing
                over_time_hourly = Nothing

            End If

        End Sub

    End Class

End Namespace