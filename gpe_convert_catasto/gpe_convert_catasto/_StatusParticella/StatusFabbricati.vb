Imports System.Data.OleDb
Imports System.IO
Imports System.Windows.Forms


Module StatusFabbricati

    Friend Function Fabbricati(ByVal dbConnection As OleDbConnection, ByVal ConversioneFile As String, ByRef BarProgress As GpeProgressBar) As String

        Dim ReadLineRows() As String
        Dim objReader As System.IO.StreamReader, NumeroRighe As Double = 0, IndexNumeroRighe As Double = 0
        Dim Folder As New IO.DirectoryInfo(Path.GetDirectoryName(ConversioneFile))
        Dim cmd As New OleDbCommand

        For Each FileTitoli As IO.FileInfo In Folder.GetFiles("*.fab", IO.SearchOption.TopDirectoryOnly)

            Application.DoEvents()
            objReader = New System.IO.StreamReader(FileTitoli.FullName)
            Do While objReader.Peek() <> -1
                objReader.ReadLine()
                NumeroRighe += 1
            Loop
            objReader.Close()
            BarProgress.PositionMax = NumeroRighe
            BarProgress.PositionMin = 1
            BarProgress.Position = 1

            objReader = New System.IO.StreamReader(FileTitoli.FullName)

            Do While objReader.Peek() <> -1

                ReadLineRows = objReader.ReadLine.Split("|")

                Select Case Val(ReadLineRows(5))
                    Case Is = 1 '1: Unità Immobiliare "cuarcuiu"

                        Call StatusFabbricati.UnitaImmobiliare(dbConnection, ReadLineRows, cmd)
                        Exit Select

                    Case Is = 2 '2: Identificativi "CUIDENTI"

                        Call StatusFabbricati.Identificativi(dbConnection, ReadLineRows, cmd)
                        Exit Select

                    Case Is = 3 '3: Indirizzo "CUINDIRI"
                        Call StatusFabbricati.Indirizzo(dbConnection, ReadLineRows, cmd)
                        Exit Select

                    Case Is = 4 '4: Utilità Comuni "CUUTILIT"
                        Call StatusFabbricati.UtilitaComuni(dbConnection, ReadLineRows, cmd, IndexNumeroRighe)
                        Exit Select

                    Case Is = 5 '5: Riserve "CURISERV"
                        Call StatusFabbricati.Riserve(dbConnection, ReadLineRows, cmd)
                        Exit Select
                End Select

                Application.DoEvents()
                IndexNumeroRighe += 1
                BarProgress.Position = IndexNumeroRighe
                '  ioio.Text = Testo & " :" & (NumeroRighe - IndexNumeroRighe) & " - " & FormatNumber(100 - (((NumeroRighe - IndexNumeroRighe) / NumeroRighe) * 100), 2)
            Loop

            objReader.Close()
        Next
        cmd.Dispose()
        Return NumeroRighe - IndexNumeroRighe


    End Function

    ''' <summary>Create Identificativi. "cuidenti"</summary>
    ''' <param name="Connect">Connessione del DB</param>
    ''' <param name="Values">Array contenente i dati da memorizzare</param>
    ''' <param name="cmd">paramertro OleDbCommand</param>
    Private Sub UnitaImmobiliare(Connect As OleDbConnection, ByVal Values() As String, cmd As OleDbCommand)


        Dim MyTransaction As OleDb.OleDbTransaction = Connect.BeginTransaction

        With cmd
            .CommandType = CommandType.Text
            .CommandText = "INSERT INTO cuarcuiu (codice,sezione,immobile,tipo_imm,progressiv,zona,categoria,classe,consistenz,superficie,rendita_l,rendita_e,lotto,edificio,scala,interno_1,interno_2,
                                                 piano_1,piano_2,piano_3,piano_4,gen_eff,gen_regist,gen_tipo,gen_numero,gen_progre,gen_anno,con_eff,con_regist,con_tipo,con_numero,con_progre,con_anno,
                                                 partita,annotazion,mutaz_iniz,mutaz_fine,prot_notif)" &
                                       "VALUES (@codice,@sezione,@immobile,@tipo_imm,@progressiv,@zona,@categoria,@classe,@consistenz,@superficie,@rendita_l,@rendita_e,@lotto,@edificio,@scala,@interno_1,@interno_2,
                                                @piano_1,@piano_2,@piano_3,@piano_4,@gen_eff,@gen_regist,@gen_tipo,@gen_numero,@gen_progre,@gen_anno,@con_eff,@con_regist,@con_tipo,@con_numero,@con_progre,@con_anno,
                                                @partita,@annotazion,@mutaz_iniz,@mutaz_fine,@prot_notif)"
            .Transaction = MyTransaction

            .Parameters.Add("@codice", OleDbType.VarChar).Value = NoNullText(Values(0))
            .Parameters.Add("@sezione", OleDbType.VarChar).Value = NoNullText(Values(1))
            .Parameters.Add("@immobile", OleDbType.Double).Value = NoNullDouble(Values(2))
            .Parameters.Add("@tipo_imm", OleDbType.VarChar).Value = NoNullText(Values(3))
            .Parameters.Add("@progressiv", OleDbType.Double).Value = NoNullDouble(Values(4))
            .Parameters.Add("@zona", OleDbType.VarChar).Value = NoNullText(Values(6))
            .Parameters.Add("@categoria", OleDbType.VarChar).Value = NoNullText(Values(7))
            .Parameters.Add("@classe", OleDbType.VarChar).Value = NoNullText(Values(8))
            .Parameters.Add("@consistenz", OleDbType.VarChar).Value = NoNullText(Values(9))
            .Parameters.Add("@superficie", OleDbType.VarChar).Value = NoNullText(Values(10))
            .Parameters.Add("@rendita_l", OleDbType.VarChar).Value = NoNullText(Values(11))
            .Parameters.Add("@rendita_e", OleDbType.VarChar).Value = NoNullText(Values(12))
            .Parameters.Add("@lotto", OleDbType.VarChar).Value = NoNullText(Values(13))
            .Parameters.Add("@edificio", OleDbType.VarChar).Value = NoNullText(Values(14))
            .Parameters.Add("@scala", OleDbType.VarChar).Value = NoNullText(Values(15))
            .Parameters.Add("@interno_1", OleDbType.VarChar).Value = NoNullText(Values(16))
            .Parameters.Add("@interno_2", OleDbType.VarChar).Value = NoNullText(Values(17))
            .Parameters.Add("@piano_1", OleDbType.VarChar).Value = NoNullText(Values(18))
            .Parameters.Add("@piano_2", OleDbType.VarChar).Value = NoNullText(Values(19))
            .Parameters.Add("@piano_3", OleDbType.VarChar).Value = NoNullText(Values(20))
            .Parameters.Add("@piano_4", OleDbType.VarChar).Value = NoNullText(Values(21))
            .Parameters.Add("@gen_eff", OleDbType.VarChar).Value = NoNullText(Values(2))
            .Parameters.Add("@gen_regist", OleDbType.VarChar).Value = NoNullText(Values(23))
            .Parameters.Add("@gen_tipo", OleDbType.VarChar).Value = NoNullText(Values(24))
            .Parameters.Add("@gen_numero", OleDbType.VarChar).Value = NoNullText(Values(25))
            .Parameters.Add("@gen_progre", OleDbType.VarChar).Value = NoNullText(Values(26))
            .Parameters.Add("@gen_anno", OleDbType.VarChar).Value = NoNullText(Values(27))
            .Parameters.Add("@con_eff", OleDbType.VarChar).Value = NoNullText(Values(28))
            .Parameters.Add("@con_regist", OleDbType.VarChar).Value = NoNullText(Values(29))
            .Parameters.Add("@con_tipo", OleDbType.VarChar).Value = NoNullText(Values(30))
            .Parameters.Add("@con_numero", OleDbType.VarChar).Value = NoNullText(Values(31))
            .Parameters.Add("@con_progre", OleDbType.VarChar).Value = NoNullText(Values(32))
            .Parameters.Add("@con_anno", OleDbType.VarChar).Value = NoNullText(Values(33))
            .Parameters.Add("@partita", OleDbType.VarChar).Value = NoNullText(Values(34))
            .Parameters.Add("@annotazion", OleDbType.VarChar).Value = NoNullText(Values(35))
            .Parameters.Add("@mutaz_iniz", OleDbType.Double).Value = NoNullDouble(Values(36))
            .Parameters.Add("@mutaz_fine", OleDbType.Double).Value = NoNullDouble(Values(37))
            .Parameters.Add("@prot_notif", OleDbType.VarChar).Value = NoNullText(Values(38))

        End With

        Try
            cmd.Connection = Connect
            cmd.ExecuteNonQuery()
            MyTransaction.Commit()

        Catch ex As Exception
            MsgBox(ex.Message & "<>" & "cuarcuiu")
            MyTransaction.Rollback()

        Finally
            cmd.Dispose()
        End Try

    End Sub


    ''' <summary>Create Identificativi. "cuidenti"</summary>
    ''' <param name="Connect">Connessione del DB</param>
    ''' <param name="Values">Array contenente i dati da memorizzare</param>
    ''' <param name="cmd">paramertro OleDbCommand</param>
    Private Sub Identificativi(Connect As OleDbConnection, ByVal Values() As String, cmd As OleDbCommand)

        Dim MyTransaction As OleDb.OleDbTransaction = Connect.BeginTransaction

        With cmd
            .CommandType = CommandType.Text
            .CommandText = "INSERT INTO cuidenti (codice,sezione,immobile,tipo_imm,progressiv,sez_urbana,foglio,particella,denominato,subalterno,edificiale,flag_impor,[fp],[fp_index])" &
                                         "VALUES (@codice,@sezione,@immobile,@tipo_imm,@progressiv,@sez_urbana,@foglio,@particella,@denominato,@subalterno,@edificiale,@flag_impor,@fp,@fp_index)"
            .Transaction = MyTransaction
            .Parameters.Add("@codice", OleDbType.VarChar).Value = NoNullText(Values(0))
            .Parameters.Add("@sezione", OleDbType.VarChar).Value = NoNullText(Values(1))
            .Parameters.Add("@immobile", OleDbType.Double).Value = NoNullDouble(Values(2))
            .Parameters.Add("@tipo_imm", OleDbType.VarChar).Value = NoNullText(Values(3))
            .Parameters.Add("@progressiv", OleDbType.Double).Value = NoNullDouble(Values(4))
            .Parameters.Add("@sez_urbana", OleDbType.VarChar).Value = NoNullText(Values(6))
            .Parameters.Add("@foglio", OleDbType.VarChar).Value = NoNullText(Values(7))
            .Parameters.Add("@particella", OleDbType.VarChar).Value = NoNullText(Values(8))
            .Parameters.Add("@denominato", OleDbType.Double).Value = NoNullDouble(Values(9))
            .Parameters.Add("@subalterno", OleDbType.VarChar).Value = NoNullText(Values(10))
            .Parameters.Add("@edificiale", OleDbType.VarChar).Value = NoNullText(Values(11))
            .Parameters.Add("@flag_impor", OleDbType.VarChar).Value = NoNullText(Values(12))
            .Parameters.Add("@fp", OleDbType.VarChar).Value = NoNullText(Values(7)) & "_" & NoNullText(Values(8))
            .Parameters.Add("@fp_index", OleDbType.Double).Value = 1 'NoNullDouble(Values(13))
        End With

        Try
            cmd.Connection = Connect
            cmd.ExecuteNonQuery()
            MyTransaction.Commit()

        Catch ex As Exception
            MsgBox(ex.Message & "<>" & "cuidenti")
            MyTransaction.Rollback()

        Finally
            cmd.Dispose()
        End Try

    End Sub


    ''' <summary>Create Identificativi. "CUINDIRI"</summary>
    ''' <param name="Connect">Connessione del DB</param>
    ''' <param name="Values">Array contenente i dati da memorizzare</param>
    ''' <param name="cmd">paramertro OleDbCommand</param>
    Private Sub Indirizzo(Connect As OleDbConnection, ByVal Values() As String, cmd As OleDbCommand)


        Dim MyTransaction As OleDb.OleDbTransaction = Connect.BeginTransaction

        With cmd
            .CommandType = CommandType.Text
            .CommandText = "INSERT INTO cuindiri (codice,sezione,immobile,tipo_imm,progressiv,toponimo,indirizzo,civico1,civico2,civico3,flag_impor)" &
                                         "VALUES (@codice,@sezione,@immobile,@tipo_imm,@progressiv,@toponimo,@indirizzo,@civico1,@civico2,@civico3,@flag_impor)"
            .Transaction = MyTransaction
            .Parameters.Add("@codice", OleDbType.VarChar).Value = NoNullText(Values(0))
            .Parameters.Add("@sezione", OleDbType.VarChar).Value = NoNullText(Values(1))
            .Parameters.Add("@immobile", OleDbType.Double).Value = NoNullDouble(Values(2))
            .Parameters.Add("@tipo_imm", OleDbType.VarChar).Value = NoNullText(Values(3))
            .Parameters.Add("@progressiv", OleDbType.Double).Value = NoNullDouble(Values(4))
            .Parameters.Add("@toponimo", OleDbType.Double).Value = NoNullDouble(Values(6))
            .Parameters.Add("@indirizzo", OleDbType.VarChar).Value = NoNullText(Values(7))
            .Parameters.Add("@civico1", OleDbType.VarChar).Value = NoNullText(Values(8))
            .Parameters.Add("@civico2", OleDbType.VarChar).Value = NoNullText(Values(9))
            .Parameters.Add("@civico3", OleDbType.VarChar).Value = NoNullText(Values(10))
            .Parameters.Add("@flag_impor", OleDbType.VarChar).Value = NoNullText(Values(11))
        End With

        Try
            cmd.Connection = Connect
            cmd.ExecuteNonQuery()
            MyTransaction.Commit()

        Catch ex As Exception
            MsgBox(ex.Message & "<>" & "cuindiri")
            MyTransaction.Rollback()

        Finally
            cmd.Dispose()
        End Try
    End Sub

    ''' <summary>Create UtilitaComuni. "cuutilit"</summary>
    ''' <param name="Connect">Connessione del DB</param>
    ''' <param name="Values">Array contenente i dati da memorizzare</param>
    ''' <param name="cmd">paramertro OleDbCommand</param>
    Private Sub UtilitaComuni(Connect As OleDbConnection, ByVal Values() As String, cmd As OleDbCommand, numeroRiga As Integer)

        Dim MyTransaction As OleDb.OleDbTransaction = Connect.BeginTransaction

        With cmd
            .CommandType = CommandType.Text
            .CommandText = "INSERT INTO CUUTILIT (codice,denominato,flag_impor,foglio,particella,subalterno,immobile,progressiv,sez_urbana,sezione,tipo_imm)" &
                                         "VALUES (@codice, @denominato, @flag_impor, @foglio, @particella ,@subalterno, @immobile, @progressiv, @sez_urbana, @sezione, @tipo_imm)"
            .Transaction = MyTransaction

            .Parameters.Add("@codice", OleDbType.VarChar).Value = Utilita.NoNullText(Values(0))
            .Parameters.Add("@denominato", OleDbType.Double).Value = Utilita.NoNullDouble(Values(1))
            .Parameters.Add("@flag_impor", OleDbType.VarChar).Value = Utilita.NoNullText("") 'Values(2))
            .Parameters.Add("@foglio", OleDbType.VarChar).Value = Utilita.NoNullText(Values(7))
            .Parameters.Add("@particella", OleDbType.VarChar).Value = Utilita.NoNullText(Values(8))
            .Parameters.Add("@subalterno", OleDbType.VarChar).Value = Utilita.NoNullText(Values(10))
            .Parameters.Add("@immobile", OleDbType.Double).Value = Utilita.NoNullDouble(Values(2))
            .Parameters.Add("@progressiv", OleDbType.Double).Value = Utilita.NoNullDouble(Values(4))
            .Parameters.Add("@sez_urbana", OleDbType.VarChar).Value = Utilita.NoNullText("") 'Values(8))
            .Parameters.Add("@sezione", OleDbType.VarChar).Value = Utilita.NoNullText("") 'Values(9))
            .Parameters.Add("@tipo_imm", OleDbType.VarChar).Value = Utilita.NoNullText(Values(3))

        End With

        Try
            cmd.Connection = Connect
            cmd.ExecuteNonQuery()
            MyTransaction.Commit()

        Catch ex As Exception
            MsgBox(ex.Message & "<>" & "cuutilit" & "<>" & numeroRiga)
            MyTransaction.Rollback()

        Finally
            cmd.Dispose()
        End Try
    End Sub

    ''' <summary>Create UtilitaComuni. "CURISERV"</summary>
    ''' <param name="Connect">Connessione del DB</param>
    ''' <param name="Values">Array contenente i dati da memorizzare</param>
    ''' <param name="cmd">paramertro OleDbCommand</param>
    Private Sub Riserve(Connect As OleDbConnection, ByVal Values() As String, cmd As OleDbCommand)

        Dim MyTransaction As OleDb.OleDbTransaction = Connect.BeginTransaction

        With cmd
            .CommandType = CommandType.Text
            .CommandText = "INSERT INTO curiserv (codice,sezione,immobile,tipo_imm,progressiv,riserva,iscrizione,flag_impor)" &
                                         "VALUES (@codice,@sezione,@immobile,@tipo_imm,@progressiv,@riserva,@iscrizione,@flag_impor)"
            .Transaction = MyTransaction
            .Parameters.Add("@codice", OleDbType.VarChar).Value = NoNullText(Values(0))
            .Parameters.Add("@sezione", OleDbType.VarChar).Value = NoNullText(Values(1))
            .Parameters.Add("@immobile", OleDbType.Double).Value = NoNullDouble(Values(2))
            .Parameters.Add("@tipo_imm", OleDbType.VarChar).Value = NoNullText(Values(3))
            .Parameters.Add("@progressiv", OleDbType.Double).Value = NoNullDouble(Values(4))
            .Parameters.Add("@riserva", OleDbType.VarChar).Value = NoNullText(Values(6))
            .Parameters.Add("@iscrizione", OleDbType.VarChar).Value = NoNullText(Values(7))
            .Parameters.Add("@flag_impor", OleDbType.VarChar).Value = NoNullText(Values(8))
        End With

        Try
            cmd.Connection = Connect
            cmd.ExecuteNonQuery()
            MyTransaction.Commit()

        Catch ex As Exception
            MsgBox(ex.Message & "<>" & "curiserv")
            MyTransaction.Rollback()

        Finally
            cmd.Dispose()
        End Try
    End Sub



End Module
