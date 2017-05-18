Public Function F_NextStep(ByVal FormArr As String, ByVal conn As String) As Object
    Dim message As Object = ""
    Dim instance As Object = Strings.Split(FormArr, ",", -1, CompareMethod.Binary)
    Dim num As Integer = 0
    Dim str5 As String = ""
    Dim str6 As String = ""
    Dim num3 As Integer = -1
    Dim str7 As String = ""
    Dim str8 As String = ""
    Dim str9 As String = ""
    Dim str10 As String = ""
    Try 
        Dim str As String = Conversions.ToString(NewLateBinding.LateIndexGet(instance, New Object() {0}, Nothing)) 'eformid
        Dim str4 As String = Conversions.ToString(NewLateBinding.LateIndexGet(instance, New Object() {1}, Nothing)) 'employee_id
        Dim str3 As String = Conversions.ToString(NewLateBinding.LateIndexGet(instance, New Object() {2}, Nothing)) 'eformsn
        Dim str2 As String = Conversions.ToString(NewLateBinding.LateIndexGet(instance, New Object() {3}, Nothing)) 'eformrole
        Dim connection As New SqlConnection(conn)
        connection.Open
        Dim command As New SqlCommand(("select flowsn from flowctl where eformsn = '" & str3 & "'"), connection)
        If Conversions.ToBoolean(NewLateBinding.LateGet(command.ExecuteReader, Nothing, "read", New Object(0  - 1) {}, Nothing, Nothing, Nothing)) Then
            num = 1
        End If
        connection.Close
        Select Case num
            Case 0
                Dim num2 As Integer
                connection.Open
                Dim obj5 As Object = New SqlCommand(String.Concat(New String() { "select object_uid,object_name,object_type from flow,SYSTEMOBJ where group_id=object_uid and eformid = '", str, "' and eformrole = '", str2, "' ORDER BY flow.y DESC " }), connection).ExecuteReader
                If Conversions.ToBoolean(NewLateBinding.LateGet(obj5, Nothing, "read", New Object(0  - 1) {}, Nothing, Nothing, Nothing)) Then
                    str5 = Conversions.ToString(NewLateBinding.LateIndexGet(obj5, New Object() { "object_uid" }, Nothing))
                    str6 = Conversions.ToString(NewLateBinding.LateIndexGet(obj5, New Object() { "object_name" }, Nothing))
                    num2 = Conversions.ToInteger(NewLateBinding.LateIndexGet(obj5, New Object() { "object_type" }, Nothing))
                End If
                connection.Close
                If (num2 = Conversions.ToDouble("2")) Then
                    connection.Open
                    Dim obj6 As Object = New SqlCommand(("select count(empuid) as PerCount from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = (select ORG_UID from EMPLOYEE where employee_id ='" & str4 & "'))"), connection).ExecuteReader
                    If Conversions.ToBoolean(NewLateBinding.LateGet(obj6, Nothing, "read", New Object(0  - 1) {}, Nothing, Nothing, Nothing)) Then
                        num3 = Conversions.ToInteger(NewLateBinding.LateIndexGet(obj6, New Object() { "PerCount" }, Nothing))
                    End If
                    connection.Close
                ElseIf (num2 = 1) Then
                    num3 = 1
                ElseIf (num2 = 3) Then
                    num3 = 1
                ElseIf (num2 = 4) Then
                    num3 = 1
                End If
                Exit Select
            Case 1
                connection.Open
                Dim obj7 As Object = New SqlCommand(String.Concat(New String() { "select deptcode,group_name,nextstep from flowctl where eformsn = '", str3, "' and empuid = '", str4, "' and hddate IS NULL " }), connection).ExecuteReader
                If Conversions.ToBoolean(NewLateBinding.LateGet(obj7, Nothing, "read", New Object(0  - 1) {}, Nothing, Nothing, Nothing)) Then
                    str7 = Conversions.ToString(NewLateBinding.LateIndexGet(obj7, New Object() { "deptcode" }, Nothing))
                    str8 = Conversions.ToString(NewLateBinding.LateIndexGet(obj7, New Object() { "group_name" }, Nothing))
                    str9 = Conversions.ToString(NewLateBinding.LateIndexGet(obj7, New Object() { "nextstep" }, Nothing))
                End If
                connection.Close
                connection.Open
                Dim obj8 As Object = New SqlCommand(("select s.object_name from flow f,SYSTEMOBJ s where f.group_id = s.object_uid and steps = '" & str9 & "'"), connection).ExecuteReader
                If Conversions.ToBoolean(NewLateBinding.LateGet(obj8, Nothing, "read", New Object(0  - 1) {}, Nothing, Nothing, Nothing)) Then
                    str10 = Conversions.ToString(NewLateBinding.LateIndexGet(obj8, New Object() { "object_name" }, Nothing))
                End If
                connection.Close
                If ((str8 = ChrW(19978) & ChrW(19968) & ChrW(32026) & ChrW(20027) & ChrW(31649)) Or (str10 = ChrW(19978) & ChrW(19968) & ChrW(32026) & ChrW(20027) & ChrW(31649))) Then
                    connection.Open
                    Dim obj9 As Object = New SqlCommand(("select count(empuid) as PerCount from EMPLOYEE where ORG_UID = (select PARENT_ORG_UID from ADMINGROUP where ORG_UID = '" & str7 & "')"), connection).ExecuteReader
                    If Conversions.ToBoolean(NewLateBinding.LateGet(obj9, Nothing, "read", New Object(0  - 1) {}, Nothing, Nothing, Nothing)) Then
                        num3 = Conversions.ToInteger(NewLateBinding.LateIndexGet(obj9, New Object() { "PerCount" }, Nothing))
                    End If
                    connection.Close
                Else
                    num3 = 1
                End If
                Exit Select
        End Select
        message = num3
    Catch exception1 As Exception
        ProjectData.SetProjectError(exception1)
        Dim exception As Exception = exception1
        message = exception.Message
        ProjectData.ClearProjectError
    End Try
    Return message
End Function

 

 
