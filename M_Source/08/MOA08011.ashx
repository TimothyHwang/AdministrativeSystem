<%@ WebHandler Language="VB" Class="MOA08011" %>

Imports System
Imports System.Web
Imports System.Data
Imports System.Data.SqlClient

Public Class MOA08011 : Implements IHttpHandler
    Public do_sql As New C_SQLFUN
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sn As String = context.Request("sn")
        Dim action As String =context.Request("action")
        Dim connstr, sql, msg As String
        Dim dbreturn As Boolean
        Dim dt As DataTable = New DataTable("securitydt")
        
        connstr = do_sql.G_conn_string
        sql = ""
        msg = ""
        If action = "smfWriteTicket" Then
            sql = "update P_08 set UPdate_Date = getdate(),Status = 0,WriteCard = 0 where EFORMSN = '" + sn + "'"
        End If
        If action = "smfverifyWriteTicket" Then
            sql = "Update P_08 set WriteCard = 1 where EFORMSN='" + sn + "'"
        End If
        
        dbreturn = do_sql.db_exec(sql, do_sql.G_conn_string)
        If dbreturn = True Then
            If action = "smfWriteTicket" Then
                msg = "smfWriteTicket"
            End If
            If action = "smfverifyWriteTicket" Then
                msg = "smfverifyWriteTicket"
            End If
        Else
            msg = "Error!!"
        End If
   
        context.Response.ContentType = "text/plain"
        context.Response.Write(msg)
    End Sub
 
    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class