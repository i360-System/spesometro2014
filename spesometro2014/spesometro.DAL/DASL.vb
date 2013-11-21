Imports System.Globalization
Imports System.Data.Common
Imports System.Data.SqlServerCe
Imports System.Data.OleDb

Module DASL

    ''' <summary>
    ''' Funzione di DASL per la convalida delle credenziali di accesso 
    ''' </summary>
    ''' <param name="cred"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Credenziali(ByVal cred As List(Of String)) As Boolean

        Dim res As Boolean = False
        Dim con As String = Nothing

        Try

            Dim Qstring As String = "select count(*) from Account where " _
                                    & "username = @nome And password = @password"
            'If Environment.OSVersion.ToString = "6.0" Or Environment.OSVersion.ToString = "6.1" Then 'vista, seven
            '    con = "Data Source=" & Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\spesometro.DAL.Database\Database2.sdf"
            ''elseif 'windows 8
            'Else
            '    con = "" ' todo xp version
            'End If
            Dim dataSource As String = Nothing
            Dim dataSourcedbg As String = My.Settings.Database1ConnectionString
            Dim dataSourceexe As String = "Data Source=" & Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\spesometro.DAL.Database\Database2.sdf"

            If Debugger.IsAttached Then
                '
                dataSource = dataSourcedbg
                '
            Else
                '
                dataSource = dataSourceexe
                '
            End If

            Using connection As New SqlCeConnection(dataSource)
                Dim command As New SqlCeCommand(Qstring, connection)
                Dim param As SqlCeParameter = Nothing

                param = New SqlCeParameter("@nome", SqlDbType.NVarChar, 50)
                command.Parameters.Add(param)
                command.Parameters("@nome").Value = cred(0).ToString
                'MsgBox(cred(0).ToString) ' msgbox1
                param = New SqlCeParameter("@password", SqlDbType.NVarChar, 50)
                command.Parameters.Add(param)
                command.Parameters("@password").Value = cred(1).ToString
                'MsgBox(cred(1).ToString)

                connection.Open()
                command.Prepare()
                Dim i = command.ExecuteScalar()
                'MsgBox(i.ToString)
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
            MsgBox(ex.ToString)
            WorkflowBL.Err(ex)
            Return res

        End Try

    End Function

    ''' <summary>
    ''' Funzione che restituisce un oggetto SQLCEconnection
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function OleDBcommandConn() As OleDbConnection

        Dim con As New OleDbConnection(MakeConnectionstring)
        'ConnectionOledb = con
        'Dim command As New OleDbCommand(insertSQL)

        ' Set the Connection to the new OleDbConnection.
        'command.Connection = con
        'CommandOleDB = command

        Return con

    End Function

    Public Function MakeConnectionstring() As String

        Dim mkcstr As String = Nothing
        If My.Settings.TipoOleDb = 0 Then 'if access 97-2003
            If My.Settings.conCredenziali Then
                
                mkcstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & My.Settings.PercorsoDB.ToString & ";" & "Jet OLEDB:Database Password=" & My.Settings.PassCred & ";"
            Else
                mkcstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & My.Settings.PercorsoDB.ToString & ";User Id=admin; Password=;"
            End If
        ElseIf My.Settings.TipoOleDb = 1 Then 'elseif access 2007 - 2013
            mkcstr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & My.Settings.PercorsoDB.ToString & ";Persist Security Info=False;"
        End If
        Return mkcstr

    End Function

    Public Property ConnectionOledb As OleDbConnection

    Public Property CommandOleDB As OleDbCommand
        Set(ByVal value As OleDbCommand)
            _CommandOleDB = value
        End Set
        Get
            Return _CommandOleDB
        End Get
    End Property
    Private _CommandOleDB As OleDbCommand

End Module
