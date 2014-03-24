Imports System.Text.RegularExpressions
Imports System.Data.Common
Imports System.Data.OleDb
'Imports Microsoft.Office.Interop.Excel
'Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.IO

Module WorkflowBL

    Const sette As Byte = 7
    Const trentasette As Byte = 37 'fornitori
    Const dieci As Byte = 10 'clienti
    Const trentatre As Byte = 33
    Const trenta As Byte = 31
    Const noteRicevute As Byte = 30
    Dim NomeFoglio As String
    Dim riga As Integer = 1
    Dim CodiceFiscaleContribuente, denominazioneAzienda, CodiceAttivita, PeriodicitaIva, partitaIva, numeroTel, Fx, emai As String
    Dim descrizioneAttivitaIva As String
    Dim indirizzo, Cap, localita, provincia As String

    Dim exc As New List(Of Exception)

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

        If My.Settings.txtMod = False Then
            Select Case val.ToString

                Case 0 ' analiticA
                    Try
                        'Dim t As New List(Of String) : Dim f As New List(Of String) : Dim p As New List(Of String)
                        'With t
                        '    .Add("")
                        'End With
                        If InitController.OpzioniGeneraliXls Then
                            ElaboraDati()
                        Else
                            MsgBox("Prima di iniziare l'elaborazione analitica," & vbCrLf & "devi inserire i file excel vuoti nel pannello ""Opzioni""," & _
                                   vbCrLf & "sotto il tab ""OutputXls"".")
                            Exit Sub
                        End If

                    Catch ex As Exception
                        MsgBox(ex.ToString())
                    End Try

                Case 1 'aggregata
                    Try
                        If InitController.OutputXLS Then
                            ElaboraDatiAggregati()
                        Else
                            MsgBox("Prima di iniziare l'elaborazione aggregata," & vbCrLf & "devi inserire un percorso dove scrivere il file, nel pannello ""Opzioni""" & _
                                  vbCrLf & "sotto il tab ""Generale"".")
                            Exit Sub
                        End If

                    Catch ex As Exception
                        MsgBox(ex.ToString())
                    End Try
            End Select
        Else

            Select Case val.ToString

                Case 0 'analitico telematico

                    MsgBox("Funzione non implementata")
                    Exit Sub

                Case 1 'aggregato telematico
                    Try
                        If InitController.OutputXLS Then
                            ElaboraDatiAggregatiTelematici()
                        Else
                            MsgBox("Prima di iniziare l'elaborazione aggregata," & vbCrLf & "devi inserire un percorso dove scrivere il file, nel pannello ""Opzioni""" & _
                                  vbCrLf & "sotto il tab ""Generale"".")
                            Exit Sub
                        End If
                    Catch ex As Exception
                        MsgBox(ex.ToString())
                    End Try

            End Select


        End If



    End Sub

    ''' <summary>
    ''' Aggiunge errori alla lista di eccezzioni.
    ''' </summary>
    ''' <param name="ex"></param>
    ''' <remarks></remarks>
    Public Sub Err(ByVal ex As Exception)

        exc.Add(ex)

    End Sub

    Private Sub ElaboraDatiAggregatiTelematici()

        Dim comboInvolucro As Byte
        Dim Criterio, quer, anagrafica, azienda, esercizio As String


        'Dim sitointernet, descrizioneAttivitaIva As String
        Dim TipoComunicazione As Char
        Dim mainDb As New DataSet
        Dim mainAd As OleDbDataAdapter
        Dim importoFattureEmesse As Double = 0 : Dim importoNoteCreditoEmesse As Double = 0
        Dim importoFattureRicevute As Double = 0 : Dim importoNoteCreditoRicevute As Double = 0 : Dim IvaFattureEmesse As Double = 0

        Dim IvaNoteCreditoEmesse As Double = 0 : Dim IvaNoteCreditoRicevute As Double = 0 : Dim IvaFattureRicevute As Double = 0
        Dim numeroFattureEmesse As Long = 0 : Dim numeroNoteCreditoEmesse As Long = 0 : Dim numeroFattureRicevute As Long = 0
        Dim numeroNoteCreditoRicevute As Long = 0 : Dim counter As Long = 0

        If Not InitController.OutputXLS Then
            MsgBox("Non è stato inserito un percorso di Output, dove verrà creato il file CSV.")
            Exit Sub
        End If
        If (ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text = "") Then
            MsgBox("Campo esercizio vuoto.", vbCritical)
            Exit Sub
        ElseIf Not IsNumeric(ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text) Then
            MsgBox("Campo esercizio non valorizzato correttamente.", vbCritical)
            Exit Sub
        End If
        If MsgBox("Procedere con l'inserimento in Excel: Azienda " & _
           ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text & " Esercizio " & ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text & " ?", vbYesNo + vbQuestion + vbDefaultButton2, "Azienda") = vbNo Then
            MsgBox("Procedura abbandonata", vbCritical)
            Exit Sub
        End If

        ElaborazioneExcell.Labelcompletato.Visible = False
        ElaborazioneExcell.Labelattendere.Visible = True

        Dim lista = New List(Of String)

        comboInvolucro = ElaborazioneExcell.UserControlMenuXLS1.ComboBox2.SelectedIndex

        Dim table, tableMovimentiIvaTestata10, tableMovimentiIvaTestata37 As System.Data.DataTable : Dim tableEserciziContabili As System.Data.DataTable : Dim tableAnagrafiche As System.Data.DataTable
        Dim tableMovimentiContabiliTestata, tableMovimentiIvaRighe As System.Data.DataTable
        Dim listaAnagraficheSommatoria As New List(Of String)

  
        Try 'main elab

            azienda = ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text
            esercizio = ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text

            ''lancio query select con input utente
            Dim pp As New OleDb.OleDbConnection(DASL.MakeConnectionstring)
            pp.Open()
            Dim command As New OleDbCommand("SELECT * FROM Aziende WHERE Azienda='" & azienda & "'")
            command.Connection = pp
            Dim rslset = command.ExecuteReader()

            If Not rslset.HasRows Then
                MsgBox("Azienda non codificata: " & ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text & "." & _
                       vbCrLf & "Oppure non sono stati impostati correttamente i parametri nel pannello opzioni:" & vbCrLf & _
                       " credenziali, database, tipo di database. Non è stato possibile eseguire la query" & vbCrLf & _
                       " di ricerca o la query di ricerca non ha prodotto risultati.")
                ElaborazioneExcell.Labelattendere.Visible = False
                pp.Close()
                Exit Sub
            Else

                mainAd = New OleDbDataAdapter(command.CommandText.ToString, pp)
                mainAd.FillSchema(mainDb, SchemaType.Source)
                mainAd.Fill(mainDb, "Aziende")
                table = mainDb.Tables("Aziende")
                CodiceFiscaleContribuente = table.Rows(0)("CodiceFiscale").ToString()
                TipoComunicazione = IIf(ElaborazioneExcell.UserControlMenuXLS1.ComboBox3.SelectedIndex = 0, "O", "S")
                CodiceAttivita = table.Rows(0)("CodiceAttivitaIva").ToString()
                partitaIva = table.Rows(0)("PartitaIva").ToString()
                numeroTel = table.Rows(0)("Telefono").ToString()

                Fx = table.Rows(0)("Fax").ToString()
                emai = table.Rows(0)("EMail").ToString()
                denominazioneAzienda = table.Rows(0)("Denominazione1").ToString()
                indirizzo = table.Rows(0)("Indirizzo").ToString()
                Cap = table.Rows(0)("Cap").ToString()
                localita = table.Rows(0)("Localita").ToString()
                provincia = table.Rows(0)("Provincia").ToString()
                'sitointernet = table.Rows(0)("SitoInternet").ToString()
                descrizioneAttivitaIva = table.Rows(0)("DescrizioneAttivitaIva").ToString()
                pp.Close()
                mainDb.Dispose()

            End If
            '''''''fine record A e B

            Dim text1 = ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text
            Dim text2 = ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text
            ''''''eventualmente   qui recuperiamo periodicitaIva
            Dim p As New OleDb.OleDbConnection(DASL.MakeConnectionstring)


            '''''''''''fine periodicitaiva
            ObjectTableMovimentiIvaTestata.Nullable()
            preProcessing(azienda, esercizio)

            ''Inizia l'elaborazione vera e propria con i dati ricavati fin qui.
            quer = "SELECT * FROM Anagrafiche WHERE NOT (TipoConto = 'N') or TipoConto Is Null order by anagrafica"
            p = DASL.OleDBcommandConn()
            p.Open()
            mainAd = New OleDbDataAdapter(quer, p)
            mainDb = New DataSet
            mainAd.FillSchema(mainDb, SchemaType.Source)

            mainAd.Fill(mainDb, "Anagrafiche")
            tableAnagrafiche = mainDb.Tables("Anagrafiche")

            mainAd = Nothing
            mainDb.Dispose()
            p.Close()
            ''''''''''''''''*'*****''**'***'*'*'**'*'*'**'*'*'*'*'*'*'*'*'**'*'*'*


            '''''elaborazione dei risultati ottenuti
            Dim ArrFiveValue() As String : Dim flg As Boolean = True
            ElaborazioneExcell.ProgressBar1.Value = Nothing : ElaborazioneExcell.ProgressBar1.Minimum = 0
            ElaborazioneExcell.ProgressBar1.Maximum = tableAnagrafiche.Rows.Count : ElaborazioneExcell.ProgressBar1.Step = 1

            Dim indiceArray As Integer : Dim countarr = arr.GetUpperBound(1)

            For Each r As DataRow In tableAnagrafiche.Rows

                flg = True

                indiceArray = 0 : Dim conter As Integer = 0

                For n = 0 To countarr 'Match tra numero anagrafica da processare e tutti i record prelevati in anagrafiche


                    If arr(0, n) = r("anagrafica").ToString() Then

                        indiceArray = n
                        flg = False
                        conter += 1
                        Exit For

                    End If

                    conter += 1

                Next n

                If flg Then GoTo prossimo

                importoFattureEmesse = 0 : importoNoteCreditoEmesse = 0
                importoFattureRicevute = 0 : importoNoteCreditoRicevute = 0 : IvaFattureEmesse = 0
                IvaNoteCreditoEmesse = 0 : IvaNoteCreditoRicevute = 0 : IvaFattureRicevute = 0
                numeroFattureEmesse = 0 : numeroNoteCreditoEmesse = 0 : numeroFattureRicevute = 0
                numeroNoteCreditoRicevute = 0
                anagrafica = r("anagrafica").ToString()
                ArrFiveValue = {r("Denominazione1").ToString, _
                                            IIf(IsNothing(r("Denominazione2").ToString), "", r("Denominazione2").ToString), _
                                            IIf(IsNothing(r("PartitaIva").ToString), "", r("PartitaIva").ToString), _
                                            IIf(IsNothing(r("CodiceFiscale").ToString), "", r("CodiceFiscale").ToString), _
                                            IIf(IsNothing(UCase(r("TipoConto").ToString)), "", UCase(r("TipoConto").ToString))}

                Dim arrlista() As String = Nothing

                If ObjectTableMovimentiIvaTestata.arr(1, indiceArray) = "10" Then

                    'restituiamo tableMovimentiIvaTestata
                    tableMovimentiIvaTestata10 = ObjectTableMovimentiIvaTestata.tableMovimentiIvaTestata(conter - 1)

                    For Each riga As DataRow In tableMovimentiIvaTestata10.Rows

                        Dim queryMovimentiContabiliTestata = "Select * from MovimentiContabiliTestata where azienda = '" & azienda & "' " _
                                          & "And esercizio = '" & esercizio & "' And NumeroPrimaNota = " & riga("NumeroPrimaNota").ToString

                        p = DASL.OleDBcommandConn()
                        p.Open()
                        mainAd = New OleDbDataAdapter(queryMovimentiContabiliTestata, p)
                        mainDb = New DataSet
                        mainAd.FillSchema(mainDb, SchemaType.Source)

                        mainAd.Fill(mainDb, "MovimentiContabiliTestata")
                        tableMovimentiContabiliTestata = mainDb.Tables("MovimentiContabiliTestata")

                        mainAd = Nothing
                        mainDb.Dispose()
                        p.Close()

                        Dim queryMovimentiIvaRighe = "Select * from MovimentiIvaRighe where azienda ='" & azienda & _
                                  "' and esercizio = '" & esercizio & "' and tiporegistro = 'V' and numeroregistro = 1 and " _
                                  & "numeroprotocollo = " & riga("NumeroProtocollo")
                        p = DASL.OleDBcommandConn()
                        p.Open()
                        mainAd = New OleDbDataAdapter(queryMovimentiIvaRighe, p)
                        mainDb = New DataSet
                        mainAd.FillSchema(mainDb, SchemaType.Source)

                        mainAd.Fill(mainDb, "MovimentiIvaRighe")
                        tableMovimentiIvaRighe = mainDb.Tables("MovimentiIvaRighe")

                        mainAd = Nothing
                        mainDb.Dispose()
                        p.Close()

                        Select Case tableMovimentiContabiliTestata.Rows(0)("Causale").ToString()

                            Case "001" 'Fattureemesse
                                'totalizzazione imponibili iva e numero documenti


                                For Each rig As DataRow In tableMovimentiIvaRighe.Rows

                                    importoFattureEmesse += CDbl(rig("Imponibile").ToString)
                                    IvaFattureEmesse += CDbl(rig("Iva").ToString)
                                    numeroFattureEmesse += 1

                                Next

                            Case "003" 'NoteCreditoemesse
                                'totalizzazione imponibili iva e numero documenti

                                For Each rig As DataRow In tableMovimentiIvaRighe.Rows

                                    importoNoteCreditoEmesse += rig("Imponibile").ToString
                                    IvaNoteCreditoEmesse += rig("Iva").ToString
                                    numeroNoteCreditoEmesse += 1

                                Next

                        End Select

                    Next

                End If

                If ObjectTableMovimentiIvaTestata.arr(1, indiceArray) = "37" Then

                    tableMovimentiIvaTestata37 = ObjectTableMovimentiIvaTestata.tableMovimentiIvaTestata(conter - 1)

                    For Each riga As DataRow In tableMovimentiIvaTestata37.Rows

                        Dim queryMovimentiContabiliTestata = "Select * from MovimentiContabiliTestata where azienda = '" & azienda & "' " _
                                          & "And esercizio = '" & esercizio & "' And NumeroPrimaNota = " & riga("NumeroPrimaNota").ToString

                        p = DASL.OleDBcommandConn()
                        p.Open()
                        mainAd = New OleDbDataAdapter(queryMovimentiContabiliTestata, p)
                        mainDb = New DataSet
                        mainAd.FillSchema(mainDb, SchemaType.Source)

                        mainAd.Fill(mainDb, "MovimentiContabiliTestata")
                        tableMovimentiContabiliTestata = mainDb.Tables("MovimentiContabiliTestata")

                        mainAd = Nothing
                        mainDb.Dispose()
                        p.Close()

                        Dim queryMovimentiIvaRighe = "Select * from MovimentiIvaRighe where azienda ='" & azienda & _
                                  "' and esercizio = '" & esercizio & "' and tiporegistro = 'A' and numeroregistro = 11 and " _
                                  & "numeroprotocollo = " & riga("NumeroProtocollo")
                        p = DASL.OleDBcommandConn()
                        p.Open()
                        mainAd = New OleDbDataAdapter(queryMovimentiIvaRighe, p)
                        mainDb = New DataSet
                        mainAd.FillSchema(mainDb, SchemaType.Source)

                        mainAd.Fill(mainDb, "MovimentiIvaRighe")
                        tableMovimentiIvaRighe = mainDb.Tables("MovimentiIvaRighe")

                        mainAd = Nothing
                        mainDb.Dispose()
                        p.Close()

                        Select Case tableMovimentiContabiliTestata.Rows(0)("Causale").ToString()

                            Case "011" 'Fatturericevute
                                'totalizzazione imponibili iva e numero documenti


                                For Each rig As DataRow In tableMovimentiIvaRighe.Rows

                                    importoFattureRicevute += CDbl(rig("Imponibile").ToString)
                                    IvaFattureRicevute += CDbl(rig("Iva").ToString)
                                    numeroFattureRicevute += 1

                                Next

                            Case "015" 'NoteCreditoRicevute
                                'totalizzazione imponibili iva e numero documenti

                                For Each rig As DataRow In tableMovimentiIvaRighe.Rows

                                    importoNoteCreditoRicevute += CDbl(rig("Imponibile").ToString)
                                    IvaNoteCreditoRicevute += CDbl(rig("Iva").ToString)
                                    numeroNoteCreditoRicevute += 1

                                Next

                        End Select
                        'riga("").ToString()

                    Next
                End If                          'piva                   'CF
                arrlista = {esercizio, ArrFiveValue(2).ToString, ArrFiveValue(3).ToString, _
                            numeroFattureEmesse, numeroFattureRicevute, importoFattureEmesse, _
                            IvaFattureEmesse, importoNoteCreditoEmesse, IvaNoteCreditoEmesse, _
                            importoFattureRicevute, IvaFattureRicevute, importoNoteCreditoRicevute, _
                            IvaNoteCreditoRicevute}
                'anno esercizio             0
                'piva societa in questione  1
                'Cf societa in questione    2
                'NFE                        3
                'NFR                        4
                'IFE                        5
                'IvFE                       6
                'INE                        7
                'IvNE                       8
                'IFR                        9
                'IvFR                      10
                'INR                       11
                'IvNR                      13

                If (importoFattureEmesse + IvaFattureEmesse + importoNoteCreditoEmesse + IvaNoteCreditoEmesse + importoFattureRicevute + IvaFattureRicevute _
                    + importoNoteCreditoRicevute + IvaNoteCreditoRicevute) > 0 Then
                    lista.AddRange(arrlista)
                    counter += 1
                End If
prossimo:
                ElaborazioneExcell.ProgressBar1.PerformStep() : ElaborazioneExcell.ProgressBar1.Refresh()
            Next

            ElaborazioneExcell.Labelattendere.Visible = False
            ElaborazioneExcell.Labelcompletato.Visible = True
            'Dim appo = ElaboraStringa(lista)
            Dim appo As New List(Of String)
            appo.AddRange(lista)
            appo.Add(counter.ToString)
            GeneraCSVTelematico(appo)
            ElaborazioneExcell.Labelxls.Visible = False
            ElaborazioneExcell.Labelelaborazione.Visible = False
            ElaborazioneExcell.Labelcompletato.Visible = True
            MsgBox("E' terminata la fase di importazione documenti Telematici", vbInformation)

        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' Genera un file CSV
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <remarks></remarks>
    Private Sub GeneraCSV(ByVal obj As List(Of String))

        Dim count As Integer
        Dim righe As Integer = 0 : Dim p As Integer = 0
        Dim nomeFile As String = Nothing

        ''if nomefile selezionato then metti quello altrimenti autocstruisco il nome
        nomeFile = "spesometro_" & ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text.ToString & "_" & ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text.ToString & ".csv"
        ElaborazioneExcell.Labelcompletato.Visible = False
        ElaborazioneExcell.Labelelaborazione.Visible = True
        Cursor.Current = Cursors.WaitCursor
        Dim tempFile2 = My.Settings.OutPutXls & "\" & nomeFile

        'File.Create(tempFile2)
        Using sw = New StreamWriter(tempFile2)
            Try
                Dim eser, Tipoel As String
                eser = ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text.ToString
                Tipoel = IIf(ElaborazioneExcell.UserControlMenuXLS1.ComboBox3.SelectedIndex = 0, "O", "S")
                Try '' Record di testa
                    sw.WriteLine("B;" & eser & ";" & CodiceFiscaleContribuente & ";" _
                                 & Tipoel & ";" & CodiceAttivita & ";" & _
                                 PeriodicitaIva & ";" & "1;;;;;;;;;;;;;;;;;")
                    'sw.WriteLine(vbCrLf)
                    'record vuoto'sw.WriteLine(";;;;;;;;;;;;;;;;;;;;;;;")
                Catch ex As Exception
                    MsgBox("Line " & ex.Message & " is invalid.  Skipping. Elaborazione terminata.")
                    Exit Sub
                End Try

                'File.Delete(inputFile)'File.Move(tempfile, inputFile)
                count = obj.Last()
                obj.RemoveAt(obj.Count - 1)
                righe = (obj.Count) / 19
                Dim colonnaD As Boolean = False
                Dim colonnaE As Boolean = False
                Dim colonnaF As Boolean = False


                For ciclo = 1 To righe 'righe

                    For i = p To (ciclo * 19) - 1  'campi

                        If ((i + 1) Mod 4 = 0) Then ' centro la quarta posizione o indice 3

                            If Not obj(i).ToString() = "" Then
                                sw.Write(obj(i).ToString() & ";") ' compilo colonna D
                                colonnaD = True
                                colonnaE = False
                                colonnaF = False
                            Else
                                sw.Write(";") 'metto vuoto e controllo colnne E ed F
                                colonnaD = False
                                colonnaE = True
                            End If

                        ElseIf ((i + 1) Mod 5 = 0) Then ' centro la quinta posizione o indice 4

                            If colonnaE Then
                                If Not obj(i).ToString() = "" Then
                                    sw.Write(obj(i).ToString() & ";") ' compilo colonna E
                                    colonnaF = False
                                Else
                                    sw.Write(";") 'metto vuoto e compilo F
                                    colonnaE = False
                                    colonnaF = True
                                End If
                            End If

                        ElseIf ((i + 1) Mod 6 = 0) Then ' centro la sesta posizione o indice 5

                            If colonnaF Then

                                sw.Write(obj(i).ToString() & ";") ' compilo colonna F
                                colonnaF = False

                            End If

                        Else ' lascio correre per tutte le altre posizioni

                            sw.Write(obj(i).ToString() & ";")

                        End If

                    Next

                    sw.Write(";;;;") ' aggiungo i campi vuoti
                    sw.Write(vbCrLf) ' vado a capo
                    p = ciclo * 19

                Next

                ' release object references.
                Cursor.Current = Cursors.Default

                sw.Close()


            Catch ex As Exception

                ex.ToString()

            Finally

                sw.Dispose()

            End Try

        End Using

    End Sub

    ''' <summary>
    ''' elaborazione di tipo aggregato
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ElaboraDatiAggregati()

        Dim Criterio, quer, anagrafica, azienda, esercizio As String
        Dim TipoComunicazione As Char
        Dim comboInvolucro As Byte
        Dim mainDb As New DataSet
        Dim mainAd As OleDbDataAdapter : Dim importoFattureEmesse As Double = 0 : Dim importoNoteCreditoEmesse As Double = 0
        Dim importoFattureRicevute As Double = 0 : Dim importoNoteCreditoRicevute As Double = 0 : Dim IvaFattureEmesse As Double = 0
        Dim IvaNoteCreditoEmesse As Double = 0 : Dim IvaNoteCreditoRicevute As Double = 0 : Dim IvaFattureRicevute As Double = 0
        Dim numeroFattureEmesse As Long = 0 : Dim numeroNoteCreditoEmesse As Long = 0 : Dim numeroFattureRicevute As Long = 0
        Dim numeroNoteCreditoRicevute As Long = 0 : Dim counter As Long = 0

        If Not InitController.OutputXLS Then
            MsgBox("Non è stato inserito un percorso di Output, dove verrà creato il file CSV.")
            Exit Sub
        End If
        If (ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text = "") Then
            MsgBox("Campo esercizio vuoto.", vbCritical)
            Exit Sub
        ElseIf Not IsNumeric(ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text) Then
            MsgBox("Campo esercizio non valorizzato correttamente.", vbCritical)
            Exit Sub
        End If
        If MsgBox("Procedere con l'inserimento in Excel: Azienda " & _
           ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text & " Esercizio " & ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text & " ?", vbYesNo + vbQuestion + vbDefaultButton2, "Azienda") = vbNo Then
            MsgBox("Procedura abbandonata", vbCritical)
            Exit Sub
        End If

        ElaborazioneExcell.Labelcompletato.Visible = False
        ElaborazioneExcell.Labelattendere.Visible = True
        Dim lista = New List(Of String)
        comboInvolucro = ElaborazioneExcell.UserControlMenuXLS1.ComboBox2.SelectedIndex
        Dim table, tableMovimentiIvaTestata10, tableMovimentiIvaTestata37 As System.Data.DataTable : Dim tableEserciziContabili As System.Data.DataTable : Dim tableAnagrafiche As System.Data.DataTable
        Dim tableMovimentiContabiliTestata, tableMovimentiIvaRighe As System.Data.DataTable
        Dim listaAnagraficheSommatoria As New List(Of String)

        Try

            azienda = ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text
            esercizio = ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text
            Dim pp As New OleDb.OleDbConnection(DASL.MakeConnectionstring)
            pp.Open()
            Dim command As New OleDbCommand("SELECT * FROM Aziende WHERE Azienda='" & azienda & "'")
            command.Connection = pp
            Dim rslset = command.ExecuteReader()

            If Not rslset.HasRows Then
                MsgBox("Azienda non codificata: " & ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text & "." & _
                       vbCrLf & "Oppure non sono stati impostati correttamente i parametri nel pannello opzioni:" & vbCrLf & _
                       " credenziali, database, tipo di database. Non è stato possibile eseguire la query" & vbCrLf & _
                       " di ricerca o la query di ricerca non ha prodotto risultati.")
                ElaborazioneExcell.Labelattendere.Visible = False
                pp.Close()
                Exit Sub
            Else

                mainAd = New OleDbDataAdapter(command.CommandText.ToString, pp)
                mainAd.FillSchema(mainDb, SchemaType.Source)
                mainAd.Fill(mainDb, "Aziende")
                table = mainDb.Tables("Aziende")
                CodiceFiscaleContribuente = table.Rows(0)("CodiceFiscale").ToString()
                TipoComunicazione = IIf(ElaborazioneExcell.UserControlMenuXLS1.ComboBox3.SelectedIndex = 0, "O", "S")
                CodiceAttivita = table.Rows(0)("CodiceAttivitaIva").ToString()
                pp.Close()
                mainDb.Dispose()

            End If

            Dim text1 = ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text
            Dim text2 = ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text
            Criterio = "SELECT * FROM EserciziContabili WHERE Azienda='" & text1 & "'" & " And Esercizio = '" & _
                text2 & "'"

            'fillo dataset
            Dim p As New OleDb.OleDbConnection(DASL.MakeConnectionstring) : mainAd = Nothing : mainDb = New DataSet
            p.Open()
            mainAd = New OleDbDataAdapter(Criterio, p)
            mainAd.FillSchema(mainDb, SchemaType.Source)
            mainAd.Fill(mainDb, "EserciziContabili")
            tableEserciziContabili = mainDb.Tables("EserciziContabili")
            PeriodicitaIva = tableEserciziContabili.Rows(0)("PeriodicitaIva").ToString()

            'rilascio
            mainAd = Nothing
            p.Close() : p.Dispose()
            mainDb.Dispose()
            ObjectTableMovimentiIvaTestata.Nullable()
            preProcessing(azienda, esercizio)

            ''Inizia l'elaborazione vera e propria con i dati ricavati fin qui.
            quer = "SELECT * FROM Anagrafiche WHERE NOT (TipoConto = 'N') or TipoConto Is Null order by anagrafica"
            p = DASL.OleDBcommandConn()
            p.Open()
            mainAd = New OleDbDataAdapter(quer, p)
            mainDb = New DataSet
            mainAd.FillSchema(mainDb, SchemaType.Source)

            mainAd.Fill(mainDb, "Anagrafiche")
            tableAnagrafiche = mainDb.Tables("Anagrafiche")

            mainAd = Nothing
            mainDb.Dispose()
            p.Close()

            Dim ArrFiveValue() As String : Dim flg As Boolean = True
            ElaborazioneExcell.ProgressBar1.Value = Nothing : ElaborazioneExcell.ProgressBar1.Minimum = 0
            ElaborazioneExcell.ProgressBar1.Maximum = tableAnagrafiche.Rows.Count : ElaborazioneExcell.ProgressBar1.Step = 1

            Dim indiceArray As Integer : Dim countarr = arr.GetUpperBound(1)

            For Each r As DataRow In tableAnagrafiche.Rows

                flg = True

                indiceArray = 0 : Dim conter As Integer = 0

                For n = 0 To countarr 'Match tra numero anagrafica da processare e tutti i record prelevati in anagrafiche


                    If arr(0, n) = r("anagrafica").ToString() Then

                        indiceArray = n
                        flg = False
                        conter += 1
                        Exit For

                    End If

                    conter += 1

                Next n

                If flg Then GoTo prossimo

                importoFattureEmesse = 0 : importoNoteCreditoEmesse = 0
                importoFattureRicevute = 0 : importoNoteCreditoRicevute = 0 : IvaFattureEmesse = 0
                IvaNoteCreditoEmesse = 0 : IvaNoteCreditoRicevute = 0 : IvaFattureRicevute = 0
                numeroFattureEmesse = 0 : numeroNoteCreditoEmesse = 0 : numeroFattureRicevute = 0
                numeroNoteCreditoRicevute = 0
                anagrafica = r("anagrafica").ToString()
                ArrFiveValue = {r("Denominazione1").ToString, _
                                            IIf(IsNothing(r("Denominazione2").ToString), "", r("Denominazione2").ToString), _
                                            IIf(IsNothing(r("PartitaIva").ToString), "", r("PartitaIva").ToString), _
                                            IIf(IsNothing(r("CodiceFiscale").ToString), "", r("CodiceFiscale").ToString), _
                                            IIf(IsNothing(UCase(r("TipoConto").ToString)), "", UCase(r("TipoConto").ToString))}

                Dim arrlista() As String = Nothing

                If ObjectTableMovimentiIvaTestata.arr(1, indiceArray) = "10" Then

                    'restituiamo tableMovimentiIvaTestata
                    tableMovimentiIvaTestata10 = ObjectTableMovimentiIvaTestata.tableMovimentiIvaTestata(conter - 1)

                    For Each riga As DataRow In tableMovimentiIvaTestata10.Rows

                        Dim queryMovimentiContabiliTestata = "Select * from MovimentiContabiliTestata where azienda = '" & azienda & "' " _
                                          & "And esercizio = '" & esercizio & "' And NumeroPrimaNota = " & riga("NumeroPrimaNota").ToString

                        p = DASL.OleDBcommandConn()
                        p.Open()
                        mainAd = New OleDbDataAdapter(queryMovimentiContabiliTestata, p)
                        mainDb = New DataSet
                        mainAd.FillSchema(mainDb, SchemaType.Source)

                        mainAd.Fill(mainDb, "MovimentiContabiliTestata")
                        tableMovimentiContabiliTestata = mainDb.Tables("MovimentiContabiliTestata")

                        mainAd = Nothing
                        mainDb.Dispose()
                        p.Close()

                        Dim queryMovimentiIvaRighe = "Select * from MovimentiIvaRighe where azienda ='" & azienda & _
                                  "' and esercizio = '" & esercizio & "' and tiporegistro = 'V' and numeroregistro = 1 and " _
                                  & "numeroprotocollo = " & riga("NumeroProtocollo")
                        p = DASL.OleDBcommandConn()
                        p.Open()
                        mainAd = New OleDbDataAdapter(queryMovimentiIvaRighe, p)
                        mainDb = New DataSet
                        mainAd.FillSchema(mainDb, SchemaType.Source)

                        mainAd.Fill(mainDb, "MovimentiIvaRighe")
                        tableMovimentiIvaRighe = mainDb.Tables("MovimentiIvaRighe")

                        mainAd = Nothing
                        mainDb.Dispose()
                        p.Close()

                        Select Case tableMovimentiContabiliTestata.Rows(0)("Causale").ToString()

                            Case "001" 'Fattureemesse
                                'totalizzazione imponibili iva e numero documenti


                                For Each rig As DataRow In tableMovimentiIvaRighe.Rows

                                    importoFattureEmesse += CDbl(rig("Imponibile").ToString)
                                    IvaFattureEmesse += CDbl(rig("Iva").ToString)
                                    numeroFattureEmesse += 1

                                Next

                            Case "003" 'NoteCreditoemesse
                                'totalizzazione imponibili iva e numero documenti

                                For Each rig As DataRow In tableMovimentiIvaRighe.Rows

                                    importoNoteCreditoEmesse += rig("Imponibile").ToString
                                    IvaNoteCreditoEmesse += rig("Iva").ToString
                                    numeroNoteCreditoEmesse += 1

                                Next

                        End Select

                    Next

                End If

                If ObjectTableMovimentiIvaTestata.arr(1, indiceArray) = "37" Then

                    tableMovimentiIvaTestata37 = ObjectTableMovimentiIvaTestata.tableMovimentiIvaTestata(conter - 1)

                    For Each riga As DataRow In tableMovimentiIvaTestata37.Rows

                        Dim queryMovimentiContabiliTestata = "Select * from MovimentiContabiliTestata where azienda = '" & azienda & "' " _
                                          & "And esercizio = '" & esercizio & "' And NumeroPrimaNota = " & riga("NumeroPrimaNota").ToString

                        p = DASL.OleDBcommandConn()
                        p.Open()
                        mainAd = New OleDbDataAdapter(queryMovimentiContabiliTestata, p)
                        mainDb = New DataSet
                        mainAd.FillSchema(mainDb, SchemaType.Source)

                        mainAd.Fill(mainDb, "MovimentiContabiliTestata")
                        tableMovimentiContabiliTestata = mainDb.Tables("MovimentiContabiliTestata")

                        mainAd = Nothing
                        mainDb.Dispose()
                        p.Close()

                        Dim queryMovimentiIvaRighe = "Select * from MovimentiIvaRighe where azienda ='" & azienda & _
                                  "' and esercizio = '" & esercizio & "' and tiporegistro = 'A' and numeroregistro = 11 and " _
                                  & "numeroprotocollo = " & riga("NumeroProtocollo")
                        p = DASL.OleDBcommandConn()
                        p.Open()
                        mainAd = New OleDbDataAdapter(queryMovimentiIvaRighe, p)
                        mainDb = New DataSet
                        mainAd.FillSchema(mainDb, SchemaType.Source)

                        mainAd.Fill(mainDb, "MovimentiIvaRighe")
                        tableMovimentiIvaRighe = mainDb.Tables("MovimentiIvaRighe")

                        mainAd = Nothing
                        mainDb.Dispose()
                        p.Close()

                        Select Case tableMovimentiContabiliTestata.Rows(0)("Causale").ToString()

                            Case "011" 'Fatturericevute
                                'totalizzazione imponibili iva e numero documenti


                                For Each rig As DataRow In tableMovimentiIvaRighe.Rows

                                    importoFattureRicevute += CDbl(rig("Imponibile").ToString)
                                    IvaFattureRicevute += CDbl(rig("Iva").ToString)
                                    numeroFattureRicevute += 1

                                Next

                            Case "015" 'NoteCreditoRicevute
                                'totalizzazione imponibili iva e numero documenti

                                For Each rig As DataRow In tableMovimentiIvaRighe.Rows

                                    importoNoteCreditoRicevute += CDbl(rig("Imponibile").ToString)
                                    IvaNoteCreditoRicevute += CDbl(rig("Iva").ToString)
                                    numeroNoteCreditoRicevute += 1

                                Next

                        End Select
                        'riga("").ToString()

                    Next
                End If
                arrlista = {"M", esercizio, CodiceFiscaleContribuente, ArrFiveValue(2).ToString, ArrFiveValue(3).ToString, "S", _
                            IIf(numeroFattureEmesse = 0, "", numeroFattureEmesse), IIf(numeroFattureRicevute = 0, "", numeroFattureRicevute), _
                            " ", IIf(importoFattureEmesse = 0, "", importoFattureEmesse), IIf(IvaFattureEmesse = 0, "", IvaFattureEmesse), "", _
                            IIf(importoNoteCreditoEmesse = 0, "", importoNoteCreditoEmesse), IIf(IvaNoteCreditoEmesse = 0, "", IvaNoteCreditoEmesse), _
                            IIf(importoFattureRicevute = 0, "", importoFattureRicevute), IIf(IvaFattureRicevute = 0, "", IvaFattureRicevute), "", _
                            IIf(importoNoteCreditoRicevute = 0, "", importoNoteCreditoRicevute), IIf(IvaNoteCreditoRicevute = 0, "", IvaNoteCreditoRicevute)}

                If (importoFattureEmesse + IvaFattureEmesse + importoNoteCreditoEmesse + IvaNoteCreditoEmesse + importoFattureRicevute + IvaFattureRicevute _
                    + importoNoteCreditoRicevute + IvaNoteCreditoRicevute) > 0 Then
                    lista.AddRange(arrlista)
                    counter += 1
                End If
prossimo:
                ElaborazioneExcell.ProgressBar1.PerformStep() : ElaborazioneExcell.ProgressBar1.Refresh()
            Next

            ElaborazioneExcell.Labelattendere.Visible = False
            ElaborazioneExcell.Labelcompletato.Visible = True
            Dim appo = ElaboraStringa(lista)
            appo.Add(counter)
            GeneraCSV(appo)
            ElaborazioneExcell.Labelxls.Visible = False
            ElaborazioneExcell.Labelelaborazione.Visible = False
            ElaborazioneExcell.Labelcompletato.Visible = True
            MsgBox("E' terminata la fase di importazione documenti in Excel", vbInformation)

        Catch ex As Exception
            MsgBox(ex.ToString & vbCrLf & "Elaborazione terminata.")
            If ElaborazioneExcell.Labelattendere.Visible = True Then ElaborazioneExcell.Labelattendere.Visible = False
        End Try


    End Sub

    ''' <summary>
    ''' Genera un file CSV
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <remarks></remarks>
    Private Sub GeneraCSVTelematico(ByVal obj As List(Of String))

        Dim numeroQuadri As Integer = 0
        Dim strin As String = Nothing
        Dim totalizzazioneRighe As Integer = 0
        For i = 1 To 1900
            strin &= " "
        Next
        Dim blank As String = "                " '16 spazi
        Dim StrinDariempire As String = strin


        Dim count, contatore As Integer
        Dim righe As Integer = 0 : Dim p As Integer = 0
        Dim nomeFile As String = Nothing

        ''if nomefile selezionato then metti quello altrimenti autocstruisco il nome
        nomeFile = "spesometro_" & ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text.ToString & "_" & ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text.ToString & ".txt"
        ElaborazioneExcell.Labelcompletato.Visible = False
        ElaborazioneExcell.Labelelaborazione.Visible = True
        Cursor.Current = Cursors.WaitCursor
        Dim tempFile2 = My.Settings.OutPutXls & "\" & nomeFile

        'File.Create(tempFile2)
        Using sw = New StreamWriter(tempFile2)
            Try
                Dim eser, Tipoel As String
                eser = ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text.ToString
                Tipoel = IIf(ElaborazioneExcell.UserControlMenuXLS1.ComboBox3.SelectedIndex = 0, "O", "S")
                Try '' Record di testa A
                    Mid(StrinDariempire, 1) = "A"
                    Mid(StrinDariempire, 16) = "NSP00"
                    With My.Settings
                        Mid(StrinDariempire, 21) = .TipoFornitore
                        Mid(StrinDariempire, 23) = .CodiceFisacaleFornitore
                        Mid(StrinDariempire, 1898) = "A"
                        Mid(StrinDariempire, 1899) = Chr(13) & Chr(10)
                        sw.WriteLine(StrinDariempire)
                        StrinDariempire = strin
                        'sw.WriteLine(vbCrLf)
                        'record vuoto'sw.WriteLine(";;;;;;;;;;;;;;;;;;;;;;;")
                    End With
                Catch ex As Exception
                    MsgBox("Line " & ex.Message & " is invalid.  Skipping. Elaborazione terminata.")
                    Exit Sub
                End Try

                StrinDariempire = strin

                Try '' Record di testa B
                    Mid(StrinDariempire, 1) = "B"
                    Mid(StrinDariempire, 2) = CodiceFiscaleContribuente
                    Mid(StrinDariempire, 18) = "       1"
                    With My.Settings
                        Mid(StrinDariempire, 74) = .CodiceFiscaleProduttoreSW
                        If ElaborazioneExcell.Tipocomunicazione = 0 Then
                            Mid(StrinDariempire, 90) = "1" 'ordinaria
                        ElseIf ElaborazioneExcell.Tipocomunicazione = 1 Then
                            Mid(StrinDariempire, 91) = "0" 'sostitutiva 
                        ElseIf ElaborazioneExcell.Tipocomunicazione = 2 Then
                            Mid(StrinDariempire, 92) = "0" 'sostitutiva (not implemented)
                        End If
                        Mid(StrinDariempire, 116) = "1" 'dati aggregati, quindi valore fisso perche siamo nell'aelaborazione aggregata
                        Mid(StrinDariempire, 117) = "0" 'dati analatici, mettiamo 0 perche non siamo nell'eleaborazione analitica
                        Mid(StrinDariempire, 118) = "1" 'Quadro FA (avendo scelto aggregato mettiamo 1)
                        Mid(StrinDariempire, 119) = "0" 'Quadro SA enumera le somme ricevute o dat senza fattura (non esiste)
                        Mid(StrinDariempire, 120) = "0" 'Quadro BL clienti/fornitori black list (stati canaglia)
                        Mid(StrinDariempire, 121) = "0" 'Quadro FE fatture emesse
                        Mid(StrinDariempire, 122) = "0" 'Quadro FR ricevute
                        Mid(StrinDariempire, 123) = "0" 'Quadro NE note credito emesse
                        Mid(StrinDariempire, 124) = "0" 'Quadro NR note credito ricevute
                        Mid(StrinDariempire, 125) = "0" 'Quadro DF 
                        Mid(StrinDariempire, 126) = "0" 'Quadro fn
                        Mid(StrinDariempire, 127) = "0" 'quadro SE
                        Mid(StrinDariempire, 128) = "0" 'quadro TU
                        Mid(StrinDariempire, 129) = "0" 'quadro TA, riepilogo!
                        Mid(StrinDariempire, 130) = partitaIva
                        Mid(StrinDariempire, 141) = CodiceAttivita
                        For Each charac As Char In numeroTel
                            If Not IsNumeric(charac) Then Replace(charac, charac, "")
                        Next
                        If numeroTel.Length > 12 Then numeroTel = Mid(numeroTel, 1, 12)
                        Mid(StrinDariempire, 147) = numeroTel
                        For Each chrt As Char In Fx
                            If Not IsNumeric(chrt) Then Replace(chrt, chrt, "")
                        Next
                        If Fx.Length > 12 Then Fx = Mid(Fx, 1, 12)
                        Mid(StrinDariempire, 159) = Fx
                        Mid(StrinDariempire, 171) = emai
                        Mid(StrinDariempire, 316) = denominazioneAzienda
                        Mid(StrinDariempire, 376) = eser
                        Mid(StrinDariempire, 382) = .CodiceFisacaleFornitore
                        Mid(StrinDariempire, 398) = .CodiceCarica
                        Mid(StrinDariempire, 400) = .DataInizioProcedura
                        Mid(StrinDariempire, 408) = .DataFineProcedura
                        Mid(StrinDariempire, 511) = denominazioneAzienda
                        Mid(StrinDariempire, 571) = .CodiceFisacaleFornitore
                        Mid(StrinDariempire, 587) = .NumeroCAF
                        Mid(StrinDariempire, 592) = .ImpegnoATrasmettere
                        Mid(StrinDariempire, 594) = .DataImpegno
                        Mid(StrinDariempire, 1898) = "A"
                        Mid(StrinDariempire, 1899) = Chr(13) & Chr(10)
                        sw.WriteLine(StrinDariempire)

                        'sw.WriteLine(vbCrLf)
                        'record vuoto'sw.WriteLine(";;;;;;;;;;;;;;;;;;;;;;;")
                    End With
                Catch ex As Exception
                    MsgBox("Line " & ex.Message & " is invalid.  Skipping. Elaborazione terminata.")
                    Exit Sub
                End Try

                Try '' Record C - Dati
                    count = obj.Last()
                    obj.RemoveAt(obj.Count - 1)
                    righe = (obj.Count) / 13
                    Dim numeroModuli As Integer = 0
                    StrinDariempire = strin
                    Dim i = 1
                    Dim colonnaD As Boolean = False
                    Dim colonnaE As Boolean = False
                    Dim colonnaF As Boolean = False

                    Dim entraPosizionale As Boolean = True
                    p = 0
                    Dim counterGiro As Integer = 0

                    For ciclo = 1 To righe 'numero di clienti/fornitori da processare

                        counterGiro += 1
                        For i = p To (ciclo * 13) - 1 'ciclo sui 13 variabili '39 --->52

                            If (i > (obj.Count - 1)) Then Exit For


                            With My.Settings
                                If entraPosizionale Then

                                    'parte posizionale
                                    entraPosizionale = False
                                    Mid(StrinDariempire, 1) = "C"
                                    Mid(StrinDariempire, 2) = CodiceFiscaleContribuente
                                    numeroModuli += 1
                                    Mid(StrinDariempire, 18) = Right("        " & numeroModuli, 8)
                                    Mid(StrinDariempire, 74) = .CodiceFiscaleProduttoreSW

                                End If 'fine posizionale

                                'righe/ciclo                1   2    3   <----ciclo

                                'anno esercizio             0   13   26   <------i          '''
                                'piva societa in questione  1   14   27                     '''FA001
                                'Cf societa in questione    2   15   28                     '''FA002
                                'NFE                        3   16   29                     '''FA004
                                'NFR                        4   17   30                     '''FA005
                                'IFE                        5   18   31                     '''FA007
                                'IvFE                       6   19   32                     '''FA008
                                'INE                        7   20   33                     '''FA010
                                'IvNE                       8   21   34                     '''FA011
                                'IFR                        9   22   35                     '''FA012
                                'IvFR                      10   23   36                     '''FA013
                                'INR                       11   24   37                     '''FA015
                                'IvNR                      12   25   38                     '''FA016
                                '                                                           '''FA003 -> vale 1 costante
                                Select Case i - (13 * (ciclo - 1))
                                    Case 1 'piva
                                        If Not obj(i).ToString() = "" Then
                                            Mid(StrinDariempire, 90 + (24 * contatore)) = "FA00" & Trim(Str(counterGiro)) & "001" & Left(obj(i).ToString() & blank, 16) & "FA00" & Trim(Str(counterGiro)) & "003" & Left(blank, 15) & "1"
                                            contatore += 2
                                            numeroQuadri += 2
                                            'sw.Write(obj(i).ToString() & ";") ' compilo colonna D
                                            colonnaD = True
                                            colonnaE = False
                                            'colonnaF = False
                                        Else
                                            ' sw.Write(";") 'metto vuoto e controllo colnne E ed F
                                            colonnaD = False
                                            colonnaE = True
                                        End If

                                    Case 2 'CF

                                        If colonnaE Then
                                            If Not obj(i).ToString() = "" Then
                                                Mid(StrinDariempire, 90 + (24 * contatore)) = "FA00" & Trim(Str(counterGiro)) & "002" & Left(obj(i).ToString() & blank, 16) & "FA00" & Trim(Str(counterGiro)) & "003" & Left(blank, 15) & "1"
                                                contatore += 2
                                                numeroQuadri += 2
                                                'sw.Write(obj(i).ToString() & ";") ' compilo colonna E
                                                'colonnaF = False
                                            Else
                                                'sw.Write(";") 'metto vuoto e compilo F
                                                colonnaE = False
                                                'colonnaF = True
                                            End If

                                        End If

                                    Case 3 'NFE

                                        If Not obj(i).ToString() = "0" Then
                                            Mid(StrinDariempire, 90 + (24 * contatore)) = "FA00" & Trim(Str(counterGiro)) & "004" & Right(blank & obj(i).ToString(), 16)
                                            contatore += 1
                                            numeroQuadri += 1
                                        End If

                                    Case 4 'NFR

                                        If Not obj(i).ToString() = "0" Then
                                            Mid(StrinDariempire, 90 + (24 * contatore)) = "FA00" & Trim(Str(counterGiro)) & "005" & Right(blank & obj(i).ToString(), 16)
                                            contatore += 1
                                            numeroQuadri += 1
                                        End If

                                    Case 5 'IFE

                                        If Not obj(i).ToString() = "0" Then
                                            Mid(StrinDariempire, 90 + (24 * contatore)) = "FA00" & Trim(Str(counterGiro)) & "007" & Right(blank & obj(i).ToString(), 16)
                                            contatore += 1
                                            numeroQuadri += 1
                                        End If

                                    Case 6 'IvFE

                                        If Not obj(i).ToString() = "0" Then
                                            Mid(StrinDariempire, 90 + (24 * contatore)) = "FA00" & Trim(Str(counterGiro)) & "008" & Right(blank & obj(i).ToString(), 16)
                                            contatore += 1
                                            numeroQuadri += 1
                                        End If

                                    Case 7 'INE

                                        If Not obj(i).ToString() = "0" Then 'tot importo note credito emesse
                                            Mid(StrinDariempire, 90 + (24 * contatore)) = "FA00" & Trim(Str(counterGiro)) & "010" & Left(blank & obj(i).ToString(), 16)
                                            contatore += 1
                                            numeroQuadri += 1
                                        End If

                                    Case 8 'IvNE

                                        If Not obj(i).ToString() = "0" Then
                                            Mid(StrinDariempire, 90 + (24 * contatore)) = "FA00" & Trim(Str(counterGiro)) & "011" & Right(blank & obj(i).ToString(), 16)
                                            contatore += 1
                                            numeroQuadri += 1
                                        End If

                                    Case 9 'IFR

                                        If Not obj(i).ToString() = "0" Then
                                            Mid(StrinDariempire, 90 + (24 * contatore)) = "FA00" & Trim(Str(counterGiro)) & "012" & Right(blank & obj(i).ToString(), 16)
                                            contatore += 1
                                            numeroQuadri += 1
                                        End If

                                    Case 10 'IvFR

                                        If Not obj(i).ToString() = "0" Then
                                            Mid(StrinDariempire, 90 + (24 * contatore)) = "FA00" & Trim(Str(counterGiro)) & "013" & Right(blank & obj(i).ToString(), 16)
                                            contatore += 1
                                            numeroQuadri += 1
                                        End If

                                    Case 11 'INR    

                                        If Not obj(i).ToString() = "0" Then
                                            Mid(StrinDariempire, 90 + (24 * contatore)) = "FA00" & Trim(Str(counterGiro)) & "015" & Right(blank & obj(i).ToString(), 16)
                                            contatore += 1
                                            numeroQuadri += 1
                                        End If

                                    Case 12  'IvNR

                                        If Not obj(i).ToString() = "0" Then
                                            Mid(StrinDariempire, 90 + (24 * contatore)) = "FA00" & Trim(Str(counterGiro)) & "016" & Right(blank & obj(i).ToString(), 16)
                                            contatore += 1
                                            numeroQuadri += 1
                                        End If

                                End Select

                            End With

                        Next

                        p = ciclo * 13 '39


                        If counterGiro Mod 3 = 0 Then

                            ' sw.WriteLine(StrinDariempire)
                            entraPosizionale = True
                            contatore = 0 ' per il prossimo giro
                            counterGiro = 0
                            Mid(StrinDariempire, 1898) = "A"
                            Mid(StrinDariempire, 1899) = Chr(13) & Chr(10)
                            'scrive record
                            sw.WriteLine(StrinDariempire)
                            totalizzazioneRighe += 1
                            StrinDariempire = strin
                        ElseIf ciclo = righe Then ' caso in cui: record finale, e diverso da 3
                            Mid(StrinDariempire, 1898) = "A"
                            Mid(StrinDariempire, 1899) = Chr(13) & Chr(10)
                            sw.WriteLine(StrinDariempire)
                            totalizzazioneRighe += 1
                            StrinDariempire = strin

                        End If

                    Next



                Catch ex As Exception
                    MsgBox("Line " & ex.Message & " is invalid.  Skipping. Elaborazione terminata.")
                    sw.Close()
                    sw.Dispose()
                    Exit Sub
                End Try

                Try 'Record E
                    With My.Settings
                        StrinDariempire = strin
                        Mid(StrinDariempire, 1) = "E"
                        Mid(StrinDariempire, 2) = CodiceFiscaleContribuente
                        Mid(StrinDariempire, 18) = "       1" ' 7 space and number one
                        Mid(StrinDariempire, 90) = "TA001001" & Right(blank & numeroQuadri.ToString, 16)
                        Mid(StrinDariempire, 1898) = "A"
                        Mid(StrinDariempire, 1899) = Chr(13) & Chr(10)
                        sw.WriteLine(StrinDariempire)
                        StrinDariempire = strin
                    End With
                Catch ex As Exception
                    MsgBox("Line " & ex.Message & " is invalid.  Skipping. Elaborazione terminata.")
                    sw.Close()
                    sw.Dispose()
                    Exit Sub
                End Try


                Try 'record Z
                    Mid(StrinDariempire, 1) = "Z"
                    'Mid(StrinDariempire, 2) = CodiceFiscaleContribuente
                    Mid(StrinDariempire, 16) = "        1" ' 8 space and number one
                    Mid(StrinDariempire, 25) = Right("         " & totalizzazioneRighe.ToString, 9)
                    Mid(StrinDariempire, 43) = "        1" ' 8 space and number one
                    Mid(StrinDariempire, 1898) = "A"
                    Mid(StrinDariempire, 1899) = Chr(13) & Chr(10)
                    sw.WriteLine(StrinDariempire)
                    StrinDariempire = strin

                Catch ex As Exception
                    MsgBox("Line " & ex.Message & " is invalid.  Skipping. Elaborazione terminata.")
                    sw.Close()
                    sw.Dispose()
                    Exit Sub
                End Try
                Cursor.Current = Cursors.Default
                sw.Close()

            Catch ex As Exception

                ex.ToString()

            Finally

                sw.Dispose()

            End Try

        End Using

    End Sub


    ''' <summary>
    ''' Elabora le stringhe di output e restituisce l'output stesso con i dati formattati secondo
    ''' controllo.
    ''' </summary>
    ''' <param name="valor"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ElaboraStringa(ByVal valor As List(Of String)) As List(Of String)

        Dim p As Integer = 0
        Dim strin, strin2 As New List(Of String)
        Dim arrIndiceStringa, insidefloater As New List(Of String)
        Dim codfisc, partiv, sestacolonna As String
        Dim righe = valor.Count / 19

        For i = 1 To righe

            For n = p To (i * 19) - 1

                If (((n - ((i - 1) * 19)) / 3) = 1) Then

                    partiv = valor(n).ToString '4
                    codfisc = valor(n + 1).ToString '5
                    insidefloater.Add(codfisc)
                    sestacolonna = valor(n + 2).ToString '6

                    If (codfisc = partiv) And (Not codfisc = "" Or Not partiv = "") Then

                        valor(n) = ""
                        valor(n + 2) = " " ' e metto F vuoto

                    ElseIf (Not Trim(partiv) = "") Then 'se la piva non è null metto la piva e non metto cf

                        'la piva non la tocco
                        valor(n + 1) = "" 'invece il cf lo metto null
                        valor(n + 2) = " " ' e metto F vuoto

                    ElseIf (Not Trim(codfisc) = "") Then ' se cf non è null metto il cf

                        valor(n + 2) = " " ' e metto F vuoto

                    End If

                End If

            Next

            p = i * 19

        Next

        p = 0
        strin = valor
        Dim floater As New List(Of List(Of String))
        Dim insideFloaterFinal As New List(Of String)

        insideFloaterFinal = insidefloater
        Dim contenitoreUno = insidefloater.Count - 1
        Dim contenitoreFinale As New List(Of String)

        For i = 0 To contenitoreUno

            If i < contenitoreUno Then
                For we = i + 1 To contenitoreUno


                    If Not insidefloater(i).ToString = "" Then

                        If (insidefloater(i).ToString = insideFloaterFinal(we).ToString) Then

                            If Not contenitoreFinale.Contains(insidefloater(i)) Then

                                contenitoreFinale.Add(insidefloater(i))

                            End If

                        End If

                    End If


                Next
            End If

            '
        Next
        Dim importo1, importo2, importo3, importo4, importo5, importo6 As Double
        Dim rigacorrente As Integer = 1
        Dim listina As New List(Of Double)
        Dim listinaIndici, listinaPrimiIndici As New List(Of Integer)
        Dim listona As New List(Of List(Of Double))
        Dim listonaIndici, listonaPrimiIndici As New List(Of List(Of Integer))
        Dim numeroVolteentra As Integer
        If contenitoreFinale.Count > 0 Then
            Dim index As Integer = 0
            'elaboroindici
            For n = 0 To contenitoreFinale.Count - 1



                For Each itl In strin

                    If (index Mod 19 = 0) And (index > 0) Then rigacorrente += 1

                    If itl = contenitoreFinale(n) And ((index - ((rigacorrente - 1) * 19)) / 4 = 1) Then 'codice fiscale quinta colonna


                        importo1 += IIf(strin(index + 5) = "", 0, strin(index + 5))
                        importo2 += IIf(strin(index + 6) = "", 0, strin(index + 6))
                        importo3 += IIf(strin(index + 10) = "", 0, strin(index + 10))
                        importo4 += IIf(strin(index + 11) = "", 0, strin(index + 11))
                        importo5 += IIf(strin(index + 2) = "", 0, strin(index + 2))
                        importo6 += IIf(strin(index + 3) = "", 0, strin(index + 3))

                        If numeroVolteentra > 0 Then
                            listinaIndici.Add(index)
                        ElseIf (Not listinaPrimiIndici.Contains(index)) And (Not listinaIndici.Contains(index)) Then
                            listinaPrimiIndici.Add(index)
                        End If

                        numeroVolteentra += 1

                    End If

                    index += 1

                Next

                rigacorrente = 1
                numeroVolteentra = 0
                index = 0

                With listina

                    .Add(importo1)
                    .Add(importo2)
                    .Add(importo3)
                    .Add(importo4)
                    .Add(importo5)
                    .Add(importo6)

                End With

                listona.Add(listina)
                listonaIndici.Add(listinaIndici)
                importo1 = 0
                importo2 = 0
                importo3 = 0
                importo4 = 0
                importo5 = 0
                importo6 = 0
                listina = New List(Of Double)
                listinaIndici = New List(Of Integer)
            Next

            'sostituisco importi e cancello righe
            Dim cnt = listinaPrimiIndici.Count

            For w = 0 To cnt - 1 'sostituisco gli importi in ogni primo record

                Dim itkl = listinaPrimiIndici(w)
                Dim itklListona = listona(w)
                strin(Convert.ToInt32(itkl + 5)) = IIf(itklListona(0) = 0, "", itklListona(0))
                strin(Convert.ToInt32(itkl + 6)) = IIf(itklListona(1) = 0, "", itklListona(1))
                strin(Convert.ToInt32(itkl + 10)) = IIf(itklListona(2) = 0, "", itklListona(2))
                strin(Convert.ToInt32(itkl + 11)) = IIf(itklListona(3) = 0, "", itklListona(3))
                strin(Convert.ToInt32(itkl + 2)) = IIf(itklListona(4) = 0, "", itklListona(4))
                strin(Convert.ToInt32(itkl + 3)) = IIf(itklListona(5) = 0, "", itklListona(5))

            Next

            Dim listonaUltimaLista As New List(Of Integer)

            'ricavo tutti gli indici doppi in una lista
            For Each mento In listonaIndici
                For Each itemMento In mento
                    listonaUltimaLista.Add(itemMento)
                Next
            Next

            listonaUltimaLista.Sort() 'ordino in maniera crescente
            listonaUltimaLista.Reverse() ' e poi inverto l'ordine

            'ricavo i range da cancellare
            Dim inizioRange, fineRange As Integer
            Dim EccoUnaLista As New List(Of List(Of Integer))
            Dim appoggiolista As New List(Of Integer)

            For Each valorelista In listonaUltimaLista
                inizioRange = valorelista - 4
                fineRange = valorelista + 14
                appoggiolista.Add(inizioRange)
                appoggiolista.Add(fineRange)
                EccoUnaLista.Add(appoggiolista)
                appoggiolista = New List(Of Integer)
            Next

            'per ogni range cancello gli elementi
            Dim rng = EccoUnaLista.Count
            For nm = 0 To rng - 1
                Dim qList = EccoUnaLista(nm)
                strin.RemoveRange(qList(0), (qList(1) - qList(0)) + 1)

            Next


        End If

        Return strin

    End Function


    ' Indirizzo.RemoveAll(AddressOf(IsInteger))

    Public Function IsInteger(o As Object) As Boolean
        Return TypeOf (o) Is Integer
    End Function

    ''' <summary>
    ''' Verifica che il record corrente della tabella anagrafiche abbia dei movimenti correlati nella tabella
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub preProcessing(ByVal az As String, ByVal es As String)

        Dim mainDb As New DataSet
        Dim mainAd As OleDbDataAdapter
        Dim table, table2, table37 As System.Data.DataTable
        Dim p As New OleDbConnection()
        Dim ret As Boolean = False
        Dim listaanagrafiche As New List(Of Long)

        ElaborazioneExcell.Labelraccoltainfo.Visible = True
        Dim querysottoconti = "Select * from sottoconti where azienda = '" & az & "' and (conto = 10 or conto = 37)"
        p = DASL.OleDBcommandConn()
        p.Open()
        mainAd = New OleDbDataAdapter(querysottoconti, p)
        mainDb = New DataSet
        mainAd.FillSchema(mainDb, SchemaType.Source, "sottoconti")

        mainAd.Fill(mainDb, "sottoconti")
        table = mainDb.Tables("sottoconti")

        mainAd = Nothing
        mainDb.Dispose()
        p.Close()

        '*******************************Progress bar e label attive***************************************
        ElaborazioneExcell.ProgressBar2.Visible = True
        ElaborazioneExcell.ProgressBar2.Value = Nothing : ElaborazioneExcell.ProgressBar2.Minimum = 0
        ElaborazioneExcell.ProgressBar2.Maximum = table.Rows.Count : ElaborazioneExcell.ProgressBar2.Step = 1
        '*************************************************************************************************
        Dim con As Integer = 0
        For Each ro As DataRow In table.Rows
            Select Case ro("Conto")

                Case 10

                    Dim quer = "select * from movimentiivatestata where azienda = '" & az & "' and esercizio = '" & es & "' and conto = " _
                       & ro("Conto").ToString & " and sottoconto = " & ro("Sottoconto").ToString
                    p = DASL.OleDBcommandConn()
                    p.Open()
                    mainAd = New OleDbDataAdapter(quer, p)
                    mainDb = New DataSet
                    mainAd.FillSchema(mainDb, SchemaType.Source)

                    mainAd.Fill(mainDb, "movimentiivatestata")
                    table2 = mainDb.Tables("movimentiivatestata")

                    mainAd = Nothing
                    mainDb.Dispose()
                    p.Close()
                    If IsDBNull(ro("anagrafica")) Then
                        MsgBox("Attenzione: non è stata assegnata un'anagrafica al conto = 10, sottoconto =" & ro("sottoconto").ToString _
                               & "." & vbCrLf & "Si prega di inserire l'anagrafica nel sottoconto, tramite l'applicativo GeCog.", MsgBoxStyle.Critical)
                        Exit Sub
                    End If
                    ObjectTableMovimentiIvaTestata.CreateMatrice(ro("anagrafica").ToString, "10")
                    ObjectTableMovimentiIvaTestata.tableMovimentiIvaTestata.Add(table2)

                Case 37

                    Dim queryIvaTestata = "Select * from MovimentiIvaTestata where azienda = '" & az & "' " _
                                          & "And esercizio = '" & es & "' And tipoRegistro = 'A' And " _
                                          & "NumeroRegistro = 11 and Conto = " & trentasette & " And Sottoconto = " _
                                          & ro("Sottoconto").ToString
                    p = DASL.OleDBcommandConn()
                    p.Open()
                    mainAd = New OleDbDataAdapter(queryIvaTestata, p)
                    mainDb = New DataSet
                    mainAd.FillSchema(mainDb, SchemaType.Source)

                    mainAd.Fill(mainDb, "MovimentiIvaTestata")
                    table37 = mainDb.Tables("MovimentiIvaTestata")

                    mainAd = Nothing
                    mainDb.Dispose()
                    p.Close()

                    If IsDBNull(ro("anagrafica")) Then
                        MsgBox("Attenzione: non è stata assegnata un'anagrafica al conto = 37, sottoconto =" & ro("sottoconto").ToString _
                               & "." & vbCrLf & "Si prega di inserire l'anagrafica nel sottoconto, tramite l'applicativo GeCog.", MsgBoxStyle.Critical)
                        Exit Sub
                    End If

                    ObjectTableMovimentiIvaTestata.CreateMatrice(ro("anagrafica").ToString, "37")
                    ObjectTableMovimentiIvaTestata.tableMovimentiIvaTestata.Add(table37)

            End Select

            'listaanagrafiche.Add(ro("anagrafica").ToString)

            ElaborazioneExcell.ProgressBar2.PerformStep() : ElaborazioneExcell.ProgressBar2.Refresh()

        Next

        ElaborazioneExcell.ProgressBar2.Visible = False : ElaborazioneExcell.Labelraccoltainfo.Visible = False

    End Sub


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

    Public Sub ElaboraDati()

        On Error Resume Next

        If MsgBox("Procedere con l'inserimento in Excel: Azienda " & _
           ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text & " Esercizio " & ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text & " ?", vbYesNo + vbQuestion + vbDefaultButton2, "Azienda") = vbNo Then
            MsgBox("Procedura abbandonata", vbCritical)
            Exit Sub
        End If
        ElaborazioneExcell.Labelcompletato.Visible = False
        ElaborazioneExcell.Labelattendere.Visible = True
        Dim lista = New List(Of String)
        Dim Criterio, anagrafica, quer, elab, tiporegistro, numeroregistro As String, i, j, k, t, tt, Riga, RigaExcel As Long
        Dim CodiceFiscaleAzienda, CodiceFiscale, PartitaIva, RagioneSociale, TipoConto, Azienda, _
            NumeroDocumento, NumeroRegistrazione As String, DataDocumento, DataRegistrazione As Date, conto, comboInvolucro As Byte, _
            sottoconto As Integer
        Dim imponibile, iva As Double

        Dim counter As Long = 0
        Dim mainDb As New DataSet
        Dim mainAd As OleDbDataAdapter

        ElaborazioneExcell.ProgressBar1.Value = 0
        elab = Nothing : numeroregistro = Nothing : tiporegistro = Nothing
        comboInvolucro = ElaborazioneExcell.UserControlMenuXLS1.ComboBox2.SelectedIndex
        Select Case comboInvolucro
            Case 0 ' FE
                If Not controlloInvolucro(0) Then
                    MsgBox("Il tipo di Quadro richiesto nell'elaborazione, non ha un file corrispondente" & vbCrLf _
                           & " selezionato nel pannello ""Opzioni"" sul quale poter scrivere." & vbCrLf _
                           & "Si prega di selezionare un file vuoto nel pannello ""Opzioni"", salvare e ripetere l'elaborazione.")
                    ElaborazioneExcell.Labelattendere.Visible = False
                    Exit Sub
                End If
                NomeFoglio = My.Settings.FlussoQuadro1.ToString
                elab = flusso.fatturEmesse
                tiporegistro = flusso.tipoRegistroFattureEmesse
                numeroregistro = flusso.numeroRegistroFattureEmesse
            Case 1 ' NE
                If Not controlloInvolucro(1) Then
                    MsgBox("Il tipo di Quadro richiesto nell'elaborazione, non ha un file corrispondente" & vbCrLf _
                           & " selezionato nel pannello ""Opzioni"" sul quale poter scrivere." & vbCrLf _
                           & "Si prega di selezionare un file vuoto nel pannello ""Opzioni"", salvare e ripetere l'elaborazione.")
                    ElaborazioneExcell.Labelattendere.Visible = False
                    Exit Sub
                End If
                NomeFoglio = My.Settings.FlussoQuadro2.ToString
                elab = flusso.noteCreditoEmesse
                tiporegistro = flusso.tipoRegistroNoteCreditoEmesse
                numeroregistro = flusso.numeroRegistroNoteCreditoEmesse
            Case 2 ' FR
                If Not controlloInvolucro(2) Then
                    MsgBox("Il tipo di Quadro richiesto nell'elaborazione, non ha un file corrispondente" & vbCrLf _
                           & " selezionato nel pannello ""Opzioni"" sul quale poter scrivere." & vbCrLf _
                           & "Si prega di selezionare il file vuoto nel pannello ""Opzioni"", salvare e ripetere l'elaborazione.")
                    ElaborazioneExcell.Labelattendere.Visible = False
                    Exit Sub
                End If
                NomeFoglio = My.Settings.FlussoQuadro3.ToString
                elab = flusso.fattureRicevute
                tiporegistro = flusso.tipoRegistroFattureRicevute
                numeroregistro = flusso.numeroRegistroFattureRicevute
            Case 3 ' NR
                If Not controlloInvolucro(3) Then
                    MsgBox("Il tipo di Quadro richiesto nell'elaborazione, non ha un file corrispondente" & vbCrLf _
                           & " selezionato nel pannello ""Opzioni"" sul quale poter scrivere." & vbCrLf _
                           & "Si prega di selezionare un file vuoto nel pannello ""Opzioni"", salvare e ripetere l'elaborazione.")
                    ElaborazioneExcell.Labelattendere.Visible = False
                    Exit Sub
                End If
                NomeFoglio = My.Settings.FlussoQuadro4.ToString
                elab = flusso.noteCreditoRicevute
                tiporegistro = flusso.tipoRegistroNoteCreditoRicevute
                numeroregistro = flusso.numeroRegistroNoteCreditoRicevute
            Case 4 ' all
                'todo
        End Select
        Dim table As System.Data.DataTable : Dim tableIvatestata As System.Data.DataTable

        ElaborazioneExcell.ProgressBar1.Minimum = 0
        Dim pp As New OleDb.OleDbConnection(DASL.MakeConnectionstring)
        pp.Open()
        Dim command As New OleDbCommand("SELECT * FROM Aziende WHERE Azienda='" & ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text & "'")
        command.Connection = pp
        Dim rslset = command.ExecuteReader()
        If Not rslset.HasRows Then
            MsgBox("Azienda non codificata: " & ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text & "." & _
                   vbCrLf & "Oppure non sono stati impostati correttamente i parametri nel pannello opzioni:" & vbCrLf & _
                   " credenziali, database, tipo di database. Non è stato possibile eseguire la query" & vbCrLf & _
                   " di ricerca o la query di ricerca non ha prodotto risultati.")
            ElaborazioneExcell.Labelattendere.Visible = False
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

        'costruisco la query
        Dim text1 = ElaborazioneExcell.UserControlMenuXLS1.TextBox1.Text
        Dim text2 = ElaborazioneExcell.UserControlMenuXLS1.TextBox2.Text
        Criterio = "SELECT * FROM MovimentiIvaTestata WHERE Azienda='" & text1 _
            & "' AND Esercizio='" & text2 & "' AND TipoRegistro = " & tiporegistro _
            & " AND NumeroRegistro = " & numeroregistro

        'fillo dataset
        Dim p As New OleDb.OleDbConnection(DASL.MakeConnectionstring) : mainAd = Nothing : mainDb = New DataSet
        p.Open()
        mainAd = New OleDbDataAdapter(Criterio, p)
        mainAd.FillSchema(mainDb, SchemaType.Source)
        mainAd.Fill(mainDb, "MovimentiIvaTestata")
        tableIvatestata = mainDb.Tables("MovimentiIvaTestata")

        'rilascio
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
            " And Causale = " & elab ' in '016','0XX'
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
                                       & r("Esercizio").ToString & "'" & " And TipoRegistro = " & tiporegistro _
                                       & " And NumeroRegistro = " & numeroregistro & " And NumeroProtocollo = " & r("NumeroProtocollo").ToString
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

                Dim arrlista() As String = Nothing
                Select Case comboInvolucro
                    Case 0 'FE
                        arrlista = {r("Esercizio").ToString(), "00", CodiceFiscaleAzienda, "2", FiveValueanagrafica(4).ToString, _
                                anagrafica, FiveValueanagrafica(2).ToString, FiveValueanagrafica(3).ToString, _
                                FiveValueanagrafica(0).ToString, FiveValueanagrafica(1).ToString, "", "", "", "", "", _
                                "", "", "", "", "", "", "", "", "", "", "", "", threeValue(2).ToString, threeValue(0).ToString, _
                                threeValue(1).ToString, imponibile + iva, iva, ""}
                        lista.AddRange(arrlista)

                        counter = counter + 1

                    Case 1 'NE
                        arrlista = {r("Esercizio").ToString(), "00", CodiceFiscaleAzienda, "2", FiveValueanagrafica(4).ToString, _
                                anagrafica, FiveValueanagrafica(2).ToString, FiveValueanagrafica(3).ToString, _
                                FiveValueanagrafica(0).ToString, FiveValueanagrafica(1).ToString, "", "", "", "", "", _
                                "", "", "", "", "", "", "", "", "", threeValue(2).ToString, threeValue(0).ToString, _
                                threeValue(1).ToString, 0, imponibile + iva, iva, ""}
                        lista.AddRange(arrlista)

                        counter = counter + 1
                    Case 3 'NR
                        arrlista = {r("Esercizio").ToString(), "00", CodiceFiscaleAzienda, "2", FiveValueanagrafica(4).ToString, _
                               anagrafica, FiveValueanagrafica(2).ToString, FiveValueanagrafica(3).ToString, _
                               FiveValueanagrafica(0).ToString, FiveValueanagrafica(1).ToString, "", "", "", "", "", _
                               "", "", "", "", "", "", "", "", threeValue(2).ToString, threeValue(0).ToString, _
                               threeValue(1).ToString, 0, imponibile + iva, iva, ""}
                        lista.AddRange(arrlista)

                        counter = counter + 1
                    Case 2 'FR
                        arrlista = {r("Esercizio").ToString(), "00", CodiceFiscaleAzienda, "2", FiveValueanagrafica(4).ToString, _
                                anagrafica, FiveValueanagrafica(2).ToString, FiveValueanagrafica(3).ToString, _
                                FiveValueanagrafica(0).ToString, FiveValueanagrafica(1).ToString, "", "", "", "", "", _
                                "", "", "", "", "", "", "", "", "", threeValue(2).ToString, threeValue(0).ToString, _
                                "", "", "", threeValue(1).ToString, imponibile + iva, iva, ""}
                        lista.AddRange(arrlista)

                        counter = counter + 1

                End Select
            End If
Prossimo:
        Next
        ElaborazioneExcell.Labelattendere.Visible = False
        ElaborazioneExcell.Labelcompletato.Visible = True
        lista.Add(counter)
        ProduciXls(lista, comboInvolucro)
        ElaborazioneExcell.Labelxls.Visible = False
        ElaborazioneExcell.Labelcompletato.Visible = True
        MsgBox("E' terminata la fase di importazione documenti in Excel", vbInformation)

    End Sub

    ''' <summary>
    ''' Apre e popola un file excel
    ''' </summary>
    ''' <param name="obj"></param>
    ''' obj = lista con i dati.
    ''' val = dividendo
    ''' <param name="val"></param>
    ''' <remarks></remarks>
    Private Sub ProduciXls(ByVal obj As List(Of String), ByVal val As Byte)
        '#If EarlyBinding = 1 Then
        '    Rem VB IDE

        '    Rem OUTLOOK
        '    Dim myOlApp         As Outlook.Application
        '    Dim myNameSpace     As Outlook.NameSpace

        '    Rem CONTACT
        '    Dim myContacts      As Outlook.Items
        '    Dim myItem          As Outlook.ContactItem

        '    Rem APPOINTMENT
        '    Dim myAppointments  As Outlook.Items
        '    Dim myRestrictItems As Outlook.Items
        '    Dim myAppItem       As Outlook.AppointmentItem

        '    Rem Used both for CONTACTS and APPOINTMENTS
        '    Dim objItems        As Outlook.ItemProperties
        '    Dim objItem         As Outlook.ItemProperty
        '#Else
        '        REM EXE stand alone

        '        REM OUTLOOK
        '        Dim myOlApp As Object 'Outlook.Application
        '        Dim myNameSpace As Object 'Outlook.NameSpace

        '        REM CONTACT
        '        Dim myContacts As Object 'Outlook.Items
        '        Dim myItem As Object 'Outlook.ContactItem

        '        REM APPOINTMENT
        '        Dim myAppointments As Object 'Outlook.Items
        '        Dim myRestrictItems As Object 'Outlook.Items
        '        Dim myAppItem As Object 'Outlook.AppointmentItem

        '        REM Used both for CONTACTS and APPOINTMENTS
        '        Dim objItems As Object 'Outlook.ItemProperties
        '        Dim objItem As Object 'Outlook.ItemProperty
        'If App.LogMode = 1 Then ' sta eseguendo l'EXE (compilato)
        '    ' uso il Late-Binding
        '    myOlApp = CreateObject("Outlook.Application")
        'Else ' sta eseguendo il progetto nell'IDE
        '    ' uso l'Early-Binding
        '    myOlApp = Outlook.Application
        'End If


        Dim oXL As Object 'Excel.Application '  
        Dim oWB As Object 'Excel.Workbook '
        Dim oSheet As Object 'Excel.Worksheet
        Dim oRng As Object 'Excel.Range


        Dim count As Integer
        Dim righe As Integer : Dim p As Integer = 0
        Dim numcol As Byte
        Dim operando As Byte
        Try
            Select Case val
                Case 0, 2 'Fe,FR
                    operando = trentatre
                Case 1 'NE
                    operando = trenta
                Case 3 'NR
                    operando = noteRicevute
            End Select


            ElaborazioneExcell.Labelcompletato.Visible = False
            ElaborazioneExcell.Labelxls.Visible = True
            Cursor.Current = Cursors.WaitCursor

            ' Start Excel and get Application object.
            oXL = CreateObject("Excel.Application")
            oXL.Visible = My.Settings.MostraExcel

            ' Get a  workbook NomeFoglio.
            oWB = oXL.Workbooks.Add(NomeFoglio)
            oSheet = oWB.ActiveSheet
            Dim arr() As String = Nothing
            Dim ar() = obj.ToArray
            count = ar.Last
            ReDim Preserve ar(UBound(ar) - 1)
            righe = (UBound(ar) + 1) / operando
            For ciclo = 1 To righe 'righe
                For i = p To (ciclo * operando) - 1 'campi
                    ReDim Preserve arr(numcol)
                    arr(numcol) = ar(i)
                    numcol = numcol + 1
                Next
                p = ciclo * operando
                numcol = 0
                If val = 0 Or val = 2 Then
                    oSheet.Range("A" & (ciclo + 1), "AG" & (ciclo + 1)).Value = arr
                ElseIf val = 1 Then
                    oSheet.Range("A" & (ciclo + 1), "AE" & (ciclo + 1)).Value = arr
                ElseIf val = 3 Then
                    oSheet.Range("A" & (ciclo + 1), "AD" & (ciclo + 1)).Value = arr
                End If

            Next

            oSheet.SaveAs(NomeFoglio)
            ' release object references.
            Cursor.Current = Cursors.Default

        Catch ex As Exception
            ex.ToString()
            oRng = Nothing
            oSheet = Nothing
            oWB = Nothing
            oXL = Nothing
        Finally

            oRng = Nothing
            oSheet = Nothing
            'oXL.Quit()
            oWB.close()
            oWB = Nothing
            oXL.Quit()
            Marshal.ReleaseComObject(oXL)
            oXL = Nothing

        End Try

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

    Private Function controlloInvolucro(ByVal valore As Integer) As Boolean

        Dim ret As Boolean = False

        Select Case valore
            Case 0 'fe
                If Not My.Settings.FlussoQuadro1 = "" Then ret = True
            Case 1 'ce
                If Not My.Settings.FlussoQuadro2 = "" Then ret = True
            Case 2 'fr
                If Not My.Settings.FlussoQuadro3 = "" Then ret = True
            Case 3 'cr
                If Not My.Settings.FlussoQuadro4 = "" Then ret = True
            Case 4 'todo
            Case 5 'todo
            Case 6 'todo
            Case 7 'todo
            Case Else
                ret = False
        End Select

        Return ret

    End Function

#Region "Strutture"
    Private Structure flusso
        Shared fatturEmesse = "'001'"
        Shared fattureRicevute = "'011'"
        Shared noteCreditoEmesse = "'003'"
        Shared noteCreditoRicevute = "'015'"
        Shared tipoRegistroFattureEmesse = "'V'"
        Shared tipoRegistroFattureRicevute = "'A'"
        Shared tipoRegistroNoteCreditoEmesse = "'V'"
        Shared tipoRegistroNoteCreditoRicevute = "'A'"
        Shared numeroRegistroFattureEmesse = 1
        Shared numeroRegistroFattureRicevute = 11
        Shared numeroRegistroNoteCreditoEmesse = 1
        Shared numeroRegistroNoteCreditoRicevute = 11
        Shared formatoComunicazione = "1"
    End Structure
    Private Structure QueryBuilder
        Shared insert = 0
        Shared selec = "Select * from @table where @field1 = @param1"
        Shared update = 0
        Shared delete = 0
        Shared FieldX = "@fieldX"
        Shared ParamX = "@paramX"
    End Structure


#End Region

    '<ObsoleteAttribute("Microsoft.VisualBasic.Compatibility.* classes are obsolete and supported within 32 bit processes only. http://go.microsoft.com/fwlink/?linkid=160862")> _
    'Public Class FixedLengthString

    'End Class


End Module
'File.Delete(inputFile)'File.Move(tempfile, inputFile)
'count = obj.Last()
'obj.RemoveAt(obj.Count - 1)
'righe = (obj.Count) / 19
'Dim colonnaD As Boolean = False
'Dim colonnaE As Boolean = False
'Dim colonnaF As Boolean = False
'If ElaborazioneExcell.Tipocomunicazione = 0 Then
'    Mid(StrinDariempire, 90) = "1" 'ordinaria
'ElseIf ElaborazioneExcell.Tipocomunicazione = 1 Then
'    Mid(StrinDariempire, 91) = "0" 'sostitutiva 
'ElseIf ElaborazioneExcell.Tipocomunicazione = 2 Then
'    Mid(StrinDariempire, 92) = "0" 'sostitutiva (not implemented)
'End If

'Mid(StrinDariempire, 116) = "1" 'dati aggregati, quindi valore fisso perche siamo nell'aelaborazione aggregata
'Mid(StrinDariempire, 117) = "0" 'dati analatici, mettiamo 0 perche non siamo nell'eleaborazione analitica
'Mid(StrinDariempire, 118) = "1" 'Quadro FA (avendo scelto aggregato mettiamo 1)
'Mid(StrinDariempire, 119) = "0" 'Quadro SA enumera le somme ricevute o dat senza fattura (non esiste)
'Mid(StrinDariempire, 120) = "0" 'Quadro BL clienti/fornitori black list (stati canaglia)
'Mid(StrinDariempire, 121) = "0" 'Quadro FE fatture emesse
'Mid(StrinDariempire, 122) = "0" 'Quadro FR ricevute
'Mid(StrinDariempire, 123) = "0" 'Quadro NE note credito emesse
'Mid(StrinDariempire, 124) = "0" 'Quadro NR note credito ricevute
'Mid(StrinDariempire, 125) = "0" 'Quadro DF 
'Mid(StrinDariempire, 126) = "0" 'Quadro fn
'Mid(StrinDariempire, 127) = "0" 'quadro SE
'Mid(StrinDariempire, 128) = "0" 'quadro TU
'Mid(StrinDariempire, 129) = "0" 'quadro TA, riepilogo!
'Mid(StrinDariempire, 130) = partitaIva
'Mid(StrinDariempire, 141) = CodiceAttivita
'For Each charac In numeroTel
'    If Not IsNumeric(charac) Then Replace(charac, charac, "")
'Next
'If numeroTel.Length > 12 Then numeroTel = Mid(numeroTel, 1, 12)
'Mid(StrinDariempire, 147) = numeroTel
'For Each chrt In Fx
'    If Not IsNumeric(chrt) Then Replace(chrt, chrt, "")
'Next
'If Fx.Length > 12 Then Fx = Mid(Fx, 1, 12)
'Mid(StrinDariempire, 159) = Fx
'Mid(StrinDariempire, 171) = emai
'Mid(StrinDariempire, 316) = denominazioneAzienda
'Mid(StrinDariempire, 376) = eser
'Mid(StrinDariempire, 382) = .CodiceFisacaleFornitore
'Mid(StrinDariempire, 398) = .CodiceCarica
'Mid(StrinDariempire, 400) = .DataInizioProcedura
'Mid(StrinDariempire, 408) = .DataFineProcedura
'Mid(StrinDariempire, 511) = denominazioneAzienda
'Mid(StrinDariempire, 571) = .CodiceFisacaleFornitore
'Mid(StrinDariempire, 587) = .NumeroCAF
'Mid(StrinDariempire, 592) = .ImpegnoATrasmettere
'Mid(StrinDariempire, 594) = .DataImpegno
'Mid(StrinDariempire, 1898) = "A"
'Mid(StrinDariempire, 1899) = Chr(13) & Chr(10)
'sw.WriteLine(StrinDariempire)
'sw.WriteLine(vbCrLf)
'record vuoto'sw.WriteLine(";;;;;;;;;;;;;;;;;;;;;;;")