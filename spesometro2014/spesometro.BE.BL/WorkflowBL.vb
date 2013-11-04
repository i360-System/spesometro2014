Imports System.Text.RegularExpressions
Imports System.Data.Common
Imports System.Data.OleDb
Imports Microsoft.Office.Interop.Excel

Module WorkflowBL

    Dim exc As List(Of Exception)

    ''' <summary>
    ''' MEtodo main di elaborazione del tracciato excell.
    ''' val = 0 --> select
    ''' val = 1 --> insert
    ''' val = 2 --> update
    ''' val = 3 --> delete
    ''' </summary>
    ''' <param name="val"></param>
    ''' <remarks></remarks>
    Public Sub mainXls(ByVal val As String)

        Select Case val.ToString

            Case 0 ' select
                Try
                    'Dim t As New List(Of String) : Dim f As New List(Of String) : Dim p As New List(Of String)
                    'With t
                    '    .Add("")
                    'End With
                    GeneraXls()
                Catch ex As Exception
                    MsgBox(ex.ToString())
                End Try

            Case "insert"
            Case "update"
            Case "delete"

        End Select


    End Sub

    Public Sub Err(ByVal ex As Exception)

        exc.Add(ex)

    End Sub

    Private Sub GeneraXls()

    End Sub

    Private Function prepare2() As String

    End Function

    ''' <summary>
    ''' Funzione che accetta in ingresso parametri list of string ed
    ''' elabora la query sostituendoli
    ''' </summary>
    ''' <param name="query"></param>
    ''' <param name="table"></param>
    ''' <param name="field"></param>
    ''' <param name="param"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function prepare(ByVal query As String, ByVal table As List(Of String), ByVal field As List(Of String), ByVal param As List(Of String)) As String

        Dim Finalquery As String = Nothing

        'table
        If table.Count > 1 Then

        Else
            Finalquery = Finalquery & Replace(query, "@table", table(0))
        End If

        'field
        If field.Count > 1 Then

        Else

            Finalquery = Finalquery & Replace(Finalquery, "@field", field(0))

        End If

        'param
        If param.Count > 1 Then

        Else

            Finalquery = Finalquery & Replace(Finalquery, "@param", param(0))

        End If

        Return Finalquery

    End Function



    Private Structure QueryBuilder
        Shared insert = 0
        Shared selec = "Select * from @table where @field1 = @param1"
        Shared update = 0
        Shared delete = 0
        Shared FieldX = "@fieldX"
        Shared ParamX = "@paramX"
    End Structure


    Public Sub FoglioExcel()

        On Error Resume Next

        If MsgBox("Procedere con l'inserimento in Excel: Azienda " & _
           ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text & " Esercizio " & ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text & " ?", vbYesNo + vbQuestion + vbDefaultButton2, nometabella) = vbNo Then
            MsgBox("Procedura abbandonata", vbCritical)
            Exit Sub
        End If

        Dim Criterio As String, i, j, k, t, tt, Riga, RigaExcel As Long
        Dim CodiceFiscaleAzienda, CodiceFiscale, PartitaIva, RagioneSociale, _
            NumeroDocumento, NumeroRegistrazione As String, DataDocumento, DataRegistrazione As Date, TipoConto As String

        Dim mainDb As New DataSet
        Dim mainAd As OleDbDataAdapter

        Dim NomeFoglio As String
        Select Case ElaborazioneExcell.UserControlMenuXLS1.ComboBox2.SelectedIndex
            Case 0 ' FE
                NomeFoglio = My.Settings.FlussoQuadro1.ToString
            Case 1 ' NE
                NomeFoglio = My.Settings.FlussoQuadro2.ToString
            Case 2 ' FR
                NomeFoglio = My.Settings.FlussoQuadro3.ToString
            Case 3 ' NR+
                NomeFoglio = My.Settings.FlussoQuadro4.ToString
            Case 4 ' all
                'todo
        End Select
        Dim table As System.Data.DataTable
        If ConnectionState.Open = 0 Then
            Dim pp = DASL.OleDBcommandConn("SELECT * FROM Aziende WHERE Azienda='" & ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text & "'")
            pp.Open()
            If CommandOleDB.ExecuteNonQuery() = 0 Then
                MsgBox("Azienda non codificata: " & ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text)
                Exit Sub
            Else
                mainAd = New OleDbDataAdapter(CommandOleDB.CommandText.ToString, ConnectionOledb)
                mainAd.FillSchema(mainDb, SchemaType.Source)
                mainAd.Fill(mainDb)
                table = mainDb.Tables("Aziende")
                CodiceFiscaleAzienda = table.Rows(0)("CodiceFiscale").ToString()
                pp.Close()
                table = Nothing
            End If

        End If
        Dim wbk As Workbook : Dim ap As ApplicationClass : Dim sht As Worksheet
        ap.Workbooks.Add(NomeFoglio)
        'sht = wbk.Sheets(0)
        wbk = ap.Workbooks.Open(NomeFoglio)
        ap.Visible = False

        'seleziona il foglio di lavoro 1 del file excel
        sht = wbk.Worksheets(1)

        Criterio = "SELECT * FROM MovimentiIvaTestata WHERE Azienda='" & ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text _
            & "' AND Esercizio='" & ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text & "' AND TipoRegistro = 'V' " _
            & "AND NumeroRegistro = 1"

        Dim p = DASL.OleDBcommandConn(Criterio) : mainAd = Nothing : mainDb = Nothing
        If ConnectionState.Closed Then
            p.Open()
        End If
        mainAd = New OleDbDataAdapter(CommandOleDB.CommandText.ToString, p)
        mainAd.FillSchema(mainDb, SchemaType.Source)
        mainAd.Fill(mainDb)
        table = mainDb.Tables("MovimentiIvaTestata")
        mainAd = Nothing
        p.Close()
        p.Dispose()
        mainDb = Nothing
        Dim table2 As DataTable
        For Each r As DataRow In table.Rows
            Dim quer = "SELECT * FROM MovimentiContabiliTestata WHERE Azienda='" & r("Azienda").value() _
            & "' AND Esercizio='" & r("Esercizio").value() & "' AND NumeroPrimaNota = " & r("NumeroPrimaNota").value()
            p = DASL.OleDBcommandConn(quer)
            p.Open()
            mainAd = New OleDbDataAdapter(quer, p)
            mainAd.FillSchema(mainDb, SchemaType.Source)
            mainAd.Fill(mainDb)
            table2 = mainDb.Tables("MovimentiContabiliTestata")
            If table2.Rows(0)("Causale").value = "001" Then
                For Each ro In table.Rows
                    ro()
                    'sviluppo riga excell dati anagrafici(tab anagrafiche) e numerici iva - azienda conto sottoconto e vado in sottoconti e ricavo il campo anagrafica
                Next
            End If
        Next
        table.Rows(0)("").ToString()

        locMov = mainDb.OpenRecordset(Criterio, dbOpenDynaset)
        If Err() <> 0 Then
            MsgBox("MovimentiIva:" & Err.Description & Chr$(13) & "ELABORAZIONE ANNULLATA.", vbCritical)
            Exit Sub
        End If
        locMov.MoveLast()
        locMov.MoveFirst()
        ElaborazioneExcell.ProgressBar1.MinimumSize = 0
        ElaborazioneExcell.ProgressBar1.MaximumSize = locMov.RecordCount
        For t = 1 To locMov.RecordCount
            ProgressBar1.Value = t
            Dim Imponibile, Iva As Double
            Imponibile = 0 : Iva = 0
            Dim locRighe As Recordset
            Criterio = "SELECT * FROM MovimentiIvaRighe WHERE Azienda='" & locMov("Azienda") & _
                       "' AND Esercizio='" & locMov("Esercizio") & _
                       "' AND TipoRegistro='" & locMov("TipoRegistro") & _
                       "' AND NumeroRegistro=" & locMov("NumeroRegistro") & _
                       " AND NumeroProtocollo=" & locMov("NumeroProtocollo")
            locRighe = mainDb.OpenRecordset(Criterio, dbOpenDynaset)
            locRighe.MoveLast()
            locRighe.MoveFirst()
            For tt = 1 To locRighe.RecordCount
                'controllo sul cod. iva
                'Select Case locMov("SiglaIva")
                '       Case "20", "21"   ' qui altri cod. iva
                Imponibile = Imponibile + locRighe("Imponibile")
                Iva = Iva + locRighe("Iva")
                'End Select
                locRighe.MoveNext()
            Next
            locRighe.Close()
            'If Imponibile >= 3000 Then
            RagioneSociale = "" : CodiceFiscale = "" : PartitaIva = "" : TipoConto = ""
            Dim Anagrafica As Long
            Anagrafica = 0
            'prelievo n.documento e data, dati rag. sociale
            rsTab = mainDb.OpenRecordset("SELECT * FROM Sottoconti WHERE Azienda='" & Text1 & "' AND Conto=" & locMov("Conto") & " AND Sottoconto=" & locMov("Sottoconto"), dbOpenDynaset)
            rsTab.MoveFirst()
            If rsTab.RecordCount = 1 Then
                Anagrafica = rsTab("Anagrafica")
            End If
            rsTab.Close()

            rsTab = mainDb.OpenRecordset("SELECT * FROM Anagrafiche WHERE Anagrafica=" & Anagrafica & "And tipoConto not in 'N','n'", dbOpenDynaset)
            rsTab.MoveFirst()
            If rsTab.RecordCount = 1 Then
                RagioneSociale = rsTab("Denominazione1")
                CodiceFiscale = rsTab("CodiceFiscale")
                PartitaIva = rsTab("PartitaIva")
                TipoConto = rsTab("TipoConto")
            End If
            rsTab.Close()
            NumeroDocumento = "" : DataDocumento = ""
            NumeroRegistrazione = "" : DataRegistrazione = ""
            rsTab = mainDb.OpenRecordset("SELECT * FROM MovimentiContabiliTestata WHERE Azienda='" & locMov("Azienda") & "' AND Esercizio='" & locMov("Esercizio") & "' AND NumeroPrimaNota=" & locMov("NumeroPrimaNota"), dbOpenDynaset)
            rsTab.MoveFirst()
            If rsTab.RecordCount = 1 Then
                NumeroDocumento = rsTab("EstremiDocumento")
                DataDocumento = rsTab("DataDocumento")
                NumeroRegistrazione = rsTab("NumeroPrimaNota")
                DataRegistrazione = rsTab("DataOperazione")
            End If
            rsTab.Close()
            'pagina attiva excel = 1
            Excel_Sheet = Excel_Book.Worksheets(1)
            'ricerca prima riga vuota disponibile
            RigaExcel = 1
            Do
                RigaExcel = RigaExcel + 1
            Loop Until Excel_Sheet.cells(RigaExcel, 1) = ""
            Excel_Sheet.cells(RigaExcel, 1) = Text2
            Excel_Sheet.cells(RigaExcel, 2) = CodiceFiscaleAzienda
            If locMov("TipoRegistro") = "V" Then k = 1 Else k = 2
            Excel_Sheet.cells(RigaExcel, 3) = k
            Excel_Sheet.cells(RigaExcel, 4) = 2
            Excel_Sheet.cells(RigaExcel, 5) = TipoConto
            Excel_Sheet.cells(RigaExcel, 6) = CodiceFiscale
            Excel_Sheet.cells(RigaExcel, 7) = PartitaIva
            Excel_Sheet.cells(RigaExcel, 8) = Mid(RagioneSociale, 1, 30)
            Excel_Sheet.cells(RigaExcel, 23) = Right("00000000" & Mid(DataRegistrazione, 1, 2) & Mid(DataRegistrazione, 4, 2) & Mid(DataRegistrazione, 7), 8)
            Excel_Sheet.cells(RigaExcel, 24) = NumeroRegistrazione
            Excel_Sheet.cells(RigaExcel, 26) = Right("00000000" & Mid(DataDocumento, 1, 2) & Mid(DataDocumento, 4, 2) & Mid(DataDocumento, 7), 8)
            Excel_Sheet.cells(RigaExcel, 27) = NumeroDocumento
            Excel_Sheet.cells(RigaExcel, 33) = 0
            Excel_Sheet.cells(RigaExcel, 34) = 1
            Excel_Sheet.cells(RigaExcel, 35) = Imponibile
            Excel_Sheet.cells(RigaExcel, 36) = Iva
            'End If
            locMov.MoveNext()
        Next
        locMov.Close()
        mainDb.Close()
        Err = 0

        'salvataggio e chiusura
        Excel_Sheet.SaveAs(NomeFoglio)
        If Err() <> 0 Then
            MsgBox("Errore Excel: " & Err.Description)
        End If
        ' chiude l'elaborazione
        Excel_App.ActiveWorkbook.Close(True)
        ' chiude excel
        Excel_App.Quit()
        Excel_Sheet = Nothing
        Excel_App = Nothing

        MsgBox("E' terminata la fase di importazione documenti in Excel", vbInformation)

    End Sub
End Module