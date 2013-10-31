﻿Imports System.Globalization
Imports System.Data.Common
Imports System.Data.SqlServerCe

Module DASL

    ''' <summary>
    ''' Funzione di DASL per la convalida delle credenziali di accesso 
    ''' </summary>
    ''' <param name="cred"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Credenziali(ByVal cred As List(Of String)) As Boolean

        Dim res As Boolean = False

        Try

            Dim Qstring As String = "select count(*) from Account where " _
                                    & "Nome = @nome And Password = @password"

            Using connection As New SqlCeConnection(My.Settings.Database1ConnectionString)
                Dim command As New SqlCeCommand(Qstring, connection)
                Dim param As SqlCeParameter = Nothing

                param = New SqlCeParameter("@nome", SqlDbType.NVarChar, 50)
                command.Parameters.Add(param)
                command.Parameters("@nome").Value = cred(0).ToString
                param = New SqlCeParameter("@password", SqlDbType.NVarChar, 50)
                command.Parameters.Add(param)
                command.Parameters("@password").Value = cred(1).ToString
                connection.Open()
                command.Prepare()
                Dim i = command.ExecuteScalar()

                'Dim DPT As SqlCeDataAdapter = New SqlCeDataAdapter(command.CommandText.ToString, connection)
                'DPT.FillSchema(DST, SchemaType.Source)
                'DPT.Fill(DST, "Account")

                'account = DST.Tables("Account")

                'Dim accountQuery = From Account In Account.AsEnumerable() _
                '                   Select Account

                'Dim result = accountQuery.Where(Function(p) p.Field(Of String)("username") _
                '                                    = cred(0) And p.Field(Of String)("password") = cred(1))
                'Dim resultCount = result.Count

                If i > 0 Then res = True

            End Using

            Return res

        Catch ex As Exception

            WorkflowBL.Err(ex)
            Return res

        End Try

    End Function

End Module