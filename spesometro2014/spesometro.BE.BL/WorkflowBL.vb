﻿Imports System.Text.RegularExpressions
Imports System.Data.Common
Imports System.Data.OleDb
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop

Module WorkflowBL

    Dim NomeFoglio As String
    Dim riga As Integer = 1
    Dim lista = New List(Of String)

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
                    riga = 1
                    ElaboraDati()
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


    Public Sub ElaboraDati()

        On Error Resume Next

        If MsgBox("Procedere con l'inserimento in Excel: Azienda " & _
           ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text & " Esercizio " & ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text & " ?", vbYesNo + vbQuestion + vbDefaultButton2, "Azienda") = vbNo Then
            MsgBox("Procedura abbandonata", vbCritical)
            Exit Sub
        End If

        Dim Criterio, anagrafica, quer As String, i, j, k, t, tt, Riga, RigaExcel As Long
        Dim CodiceFiscaleAzienda, CodiceFiscale, PartitaIva, RagioneSociale, TipoConto, Azienda, _
            NumeroDocumento, NumeroRegistrazione As String, DataDocumento, DataRegistrazione As Date, conto As Byte, _
            sottoconto As Integer
        Dim imponibile, iva As Double

        Dim counter As Long = 0
        Dim mainDb As New DataSet
        Dim mainAd As OleDbDataAdapter
        ElaborazioneExcell.Labelattendere.Visible = True
        ElaborazioneExcell.ProgressBar1.Value = 0
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
        Dim table As System.Data.DataTable : Dim tableIvatestata As System.Data.DataTable
        ' If ConnectionState.Closed = 0 Then
        'Dim pp = DASL.OleDBcommandConn()
        ElaborazioneExcell.ProgressBar1.Minimum = 0
        Dim pp As New OleDb.OleDbConnection(DASL.MakeConnectionstring)
        pp.Open()
        Dim command As New OleDbCommand("SELECT * FROM Aziende WHERE Azienda='" & ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text & "'")
        command.Connection = pp
        Dim rslset = command.ExecuteReader()
        If Not rslset.HasRows Then
            MsgBox("Azienda non codificata: " & ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text)
            pp.Close()
            Exit Sub
        Else
            mainAd = New OleDbDataAdapter(command.CommandText.ToString, pp)
            mainAd.FillSchema(mainDb, SchemaType.Source)
            mainAd.Fill(mainDb, "Aziende")
            table = mainDb.Tables("Aziende")
            CodiceFiscaleAzienda = table.Rows(0)("CodiceFiscale").ToString()
            pp.Close()
            mainDb.Dispose()
            'table = Nothing
        End If


        Dim text1 = ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text
        Dim text2 = ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text
        Criterio = "SELECT * FROM MovimentiIvaTestata WHERE Azienda='" & text1 _
            & "' AND Esercizio='" & text2 & "' AND TipoRegistro = 'V' " _
            & "AND NumeroRegistro = 1"

        Dim p As New OleDb.OleDbConnection(DASL.MakeConnectionstring) : mainAd = Nothing : mainDb = New DataSet

        p.Open()
        'Dim command2 As New OleDbCommand(Criterio)
        'command2.Connection = p

        mainAd = New OleDbDataAdapter(Criterio, p)
        mainAd.FillSchema(mainDb, SchemaType.Source)
        mainAd.Fill(mainDb, "MovimentiIvaTestata")

        tableIvatestata = mainDb.Tables("MovimentiIvaTestata")
        
        mainAd = Nothing
        p.Close() : p.Dispose()
        mainDb.Dispose()
        Dim table2 As System.Data.DataTable : Dim tablesottoconti As System.Data.DataTable : Dim anagraficaTable As  _
            System.Data.DataTable : Dim MovimentiTable As System.Data.DataTable
        Dim FiveValueanagrafica As New List(Of String) : Dim threeValue As New List(Of String) : Dim ArrFiveValue() As String
        ElaborazioneExcell.ProgressBar1.Maximum = tableIvatestata.Rows.Count
        ElaborazioneExcell.ProgressBar1.Step = 1
        For Each r As DataRow In tableIvatestata.Rows
            quer = "SELECT * FROM MovimentiContabiliTestata WHERE Azienda='" & r("Azienda").ToString _
            & "' AND Esercizio='" & text2 & "' AND NumeroPrimaNota = " & r("NumeroPrimaNota").ToString & _
            " And Causale = '001'" ' in '016','0XX'
            p = DASL.OleDBcommandConn()
            p.Open()
            'Dim command3 As New OleDbCommand(quer)
            'command3.Connection = p
            mainAd = New OleDbDataAdapter(quer, p)
            mainDb = New DataSet
            mainAd.FillSchema(mainDb, SchemaType.Source)
           
            mainAd.Fill(mainDb, "MovimentiContabiliTestata")
            table2 = mainDb.Tables("MovimentiContabiliTestata")
            mainAd = Nothing
            mainDb.Dispose()
            p.Close()
            ElaborazioneExcell.ProgressBar1.PerformStep()
            ElaborazioneExcell.ProgressBar1.Refresh()
            If table2.Rows.Count > 0 Then
                
                ' For Each ro As DataRow In tableIvatestata.Rows
                Azienda = r("Azienda").ToString()
                sottoconto = r("Sottoconto").ToString() : conto = (r("Conto").ToString())
                Dim querySottoconto = "select * from Sottoconti where Azienda='" & Azienda & "'" & " And Conto = " & conto & " And Sottoconto =" & sottoconto
                p = Nothing
                p = DASL.OleDBcommandConn()
                p.Open()
                mainDb = New DataSet
                mainAd = New OleDbDataAdapter(querySottoconto, p)
                mainAd.FillSchema(mainDb, SchemaType.Source)
                mainAd.Fill(mainDb, "Sottoconti")
                tablesottoconti = mainDb.Tables("Sottoconti")
                p.Close() : p = Nothing : mainAd = Nothing : mainDb.Dispose()
                anagrafica = tablesottoconti.Rows(0)("Anagrafica").ToString()
                Dim queryAnagrafiche = "select * from anagrafiche where Anagrafica=" & anagrafica
                p = DASL.OleDBcommandConn()
                p.Open()
                mainAd = New OleDbDataAdapter(queryAnagrafiche, p)
                mainDb = New DataSet
                mainAd.FillSchema(mainDb, SchemaType.Source)
                mainAd.Fill(mainDb, "Anagrafiche")
                anagraficaTable = mainDb.Tables("Anagrafiche")
                FiveValueanagrafica = Nothing : threeValue = Nothing
                ArrFiveValue = {anagraficaTable.Rows(0)("Denominazione1").ToString, _
                                              IIf(IsNothing(anagraficaTable.Rows(0)("Denominazione2").ToString), "", anagraficaTable.Rows(0)("Denominazione2").ToString), _
                                              IIf(IsNothing(anagraficaTable.Rows(0)("PartitaIva").ToString), "", anagraficaTable.Rows(0)("PartitaIva").ToString), _
                                              IIf(IsNothing(anagraficaTable.Rows(0)("CodiceFiscale").ToString), "", anagraficaTable.Rows(0)("CodiceFiscale").ToString), _
                                              IIf(IsNothing(UCase(anagraficaTable.Rows(0)("TipoConto").ToString)), "", UCase(anagraficaTable.Rows(0)("TipoConto").ToString))}
                FiveValueanagrafica = New List(Of String)(ArrFiveValue)
                If (Not FiveValueanagrafica(4).ToString = "N") And (IsNumeric(FiveValueanagrafica(2).ToString)) Then
                    Dim ArrthreeValue() As String
                    Dim format As String = "ddMMyyyy"
                    ArrthreeValue = {CDate(table2.Rows(0)("DataOperazione").ToString()).Date.ToString(format), _
                                     table2.Rows(0)("EstremiDocumento").ToString(), CDate(table2.Rows(0)("DataDocumento").ToString()).Date.ToString(format)}
                    threeValue = New List(Of String)(ArrthreeValue)
                Else
                    GoTo prossimo
                End If
                p = Nothing
                mainDb = New DataSet
                mainAd = Nothing
                Dim QueryMultiRecord = "Select * from MovimentiIvaRighe where Azienda = '" & Azienda & "'" & " And Esercizio = '" _
                                       & r("Esercizio").ToString & "'" & " And TipoRegistro = 'V'" _
                                       & " And NumeroRegistro = 1 And NumeroProtocollo = " & r("NumeroProtocollo").ToString
                p = DASL.OleDBcommandConn()
                p.Open()
                mainAd = New OleDbDataAdapter(QueryMultiRecord, p)
                mainAd.FillSchema(mainDb, SchemaType.Source) : mainAd.Fill(mainDb, "MovimentiIvaRighe")
                MovimentiTable = mainDb.Tables("MovimentiIvaRighe")
                imponibile = 0 : iva = 0
                For Each rigaTabella As DataRow In MovimentiTable.Rows
                    imponibile = imponibile + rigaTabella("Imponibile").ToString
                    iva = iva + rigaTabella("Iva").ToString
                    ' call
                Next

                Dim arrlista() As String
                arrlista = {r("Esercizio").ToString(), "00", CodiceFiscaleAzienda, "2", FiveValueanagrafica(4).ToString, _
                                anagrafica, FiveValueanagrafica(2).ToString, FiveValueanagrafica(3).ToString, _
                                FiveValueanagrafica(0).ToString, FiveValueanagrafica(1).ToString, "", "", "", "", "", _
                                "", "", "", "", "", "", "", "", "", "", "", "", threeValue(2).ToString, threeValue(0).ToString, _
                                threeValue(1).ToString, imponibile + iva, iva, ""}


                lista.AddRange(arrlista)

                Counter = Counter + 1

                '  Next ro
            End If
Prossimo:
        Next
        ElaborazioneExcell.Labelattendere.Visible = False
        ElaborazioneExcell.Labelcompletato.Visible = True
        lista.add(counter)
        ProduciXls(lista)
        ElaborazioneExcell.Labelxls.Visible = False
        ElaborazioneExcell.Labelcompletato.Visible = True
        MsgBox("E' terminata la fase di importazione documenti in Excel", vbInformation)

    End Sub
    Private Sub ProduciXls(ByVal obj As List(Of String))

        riga = riga + 1
        ElaborazioneExcell.Labelcompletato.Visible = False
        ElaborazioneExcell.Labelxls.Visible = True
        Cursor.Current = Cursors.WaitCursor
        Dim oXL As Excel.Application
        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet
        Dim oRng As Excel.Range

        ' Start Excel and get Application object.
        oXL = CreateObject("Excel.Application")
        oXL.Visible = True

        ' Get a  workbook NomeFoglio.
        oWB = oXL.Workbooks.Add(NomeFoglio)
        oSheet = oWB.ActiveSheet

        Dim ar() = obj.ToArray
        'For i = 0 to 
        'Next
        'oSheet.Range("A" & riga, "AG" & obj.Item(obj.Count - 1) + 1).Value = ar
        ' oSheet.Range("A2", "").Value = ar 'range
        oSheet.SaveAs(NomeFoglio)

        ' release object references.
        oRng = Nothing
        oSheet = Nothing
        oWB = Nothing
        oXL.Quit()
        oXL = Nothing

        Cursor.Current = Cursors.Default
        '-----------------------------------------
        ''Dim wbk As Workbook : Dim ap As ApplicationClass = Nothing : Dim sht As Worksheet
        ''ap.Workbooks.Add(NomeFoglio)
        ' ''sht = wbk.Sheets(0)
        ''wbk = ap.Workbooks.Open(NomeFoglio)
        ''ap.Visible = False

        ' ''seleziona il foglio di lavoro 1 del file excel
        ''sht = wbk.Worksheets(1)
        '-----------------------------------
        '   End If
        'rsTab.Close()
        ''pagina attiva excel = 1
        'Excel_Sheet = Excel_Book.Worksheets(1)
        ''ricerca prima riga vuota disponibile
        'RigaExcel = 1
        'Do
        '    RigaExcel = RigaExcel + 1
        'Loop Until Excel_Sheet.cells(RigaExcel, 1) = ""
        'Excel_Sheet.cells(RigaExcel, 1) = Text2
        'Excel_Sheet.cells(RigaExcel, 2) = CodiceFiscaleAzienda
        'If locMov("TipoRegistro") = "V" Then k = 1 Else k = 2
        'Excel_Sheet.cells(RigaExcel, 3) = k
        'Excel_Sheet.cells(RigaExcel, 4) = 2
        'Excel_Sheet.cells(RigaExcel, 5) = TipoConto
        'Excel_Sheet.cells(RigaExcel, 6) = CodiceFiscale
        'Excel_Sheet.cells(RigaExcel, 7) = PartitaIva
        'Excel_Sheet.cells(RigaExcel, 8) = Mid(RagioneSociale, 1, 30)
        'Excel_Sheet.cells(RigaExcel, 23) = Right("00000000" & Mid(DataRegistrazione, 1, 2) & Mid(DataRegistrazione, 4, 2) & Mid(DataRegistrazione, 7), 8)
        'Excel_Sheet.cells(RigaExcel, 24) = NumeroRegistrazione
        'Excel_Sheet.cells(RigaExcel, 26) = Right("00000000" & Mid(DataDocumento, 1, 2) & Mid(DataDocumento, 4, 2) & Mid(DataDocumento, 7), 8)
        'Excel_Sheet.cells(RigaExcel, 27) = NumeroDocumento
        'Excel_Sheet.cells(RigaExcel, 33) = 0
        'Excel_Sheet.cells(RigaExcel, 34) = 1
        'Excel_Sheet.cells(RigaExcel, 35) = Imponibile
        'Excel_Sheet.cells(RigaExcel, 36) = Iva
        ''End If
        'locMov.MoveNext()
        'Next
        'locMov.Close()
        'mainDb.Close()
        'Err = 0

        ''salvataggio e chiusura
        'Excel_Sheet.SaveAs(NomeFoglio)
        'If Err() <> 0 Then
        '    MsgBox("Errore Excel: " & Err.Description)
        'End If
        '' chiude l'elaborazione
        'Excel_App.ActiveWorkbook.Close(True)
        '' chiude excel
        'Excel_App.Quit()
        'Excel_Sheet = Nothing
        'Excel_App = Nothing

    End Sub

End Module