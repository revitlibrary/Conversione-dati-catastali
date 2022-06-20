Imports System.Data.OleDb
Imports System.IO
Imports System.Windows.Forms

Module Titoli
    ''' <summary>Initializes new instance create WriteTitola.</summary>
    ''' <returns>Value True False and numero soggetti</returns>
    Friend Function WriteTitoli(ByVal dbConnection As OleDbConnection, ByVal ConversioneFile As String, ByRef BarProgress As GpeProgressBar) As Integer
        'ByVal NameFiles As String, ByVal Descrizione As String, ByVal PathFileMDB As String, ByRef BarProgress As ProgressBar, ByRef dlgPannello As System.Windows.Forms.Form) As Integer


        Dim ReadLineRows() As String
        Dim objReader As System.IO.StreamReader, NumeroRighe As Double = 0, IndexNumeroRighe As Double = 0
        Dim Folder As New IO.DirectoryInfo(Path.GetDirectoryName(ConversioneFile))
        Dim cmd As New OleDbCommand

        ' dlgPannello.Text = "Conversione dati [" & Descrizione & "]"
        ' 0 CODICE
        ' 1 SEZIONE
        ' 2 SOGGETTO
        ' 3 TIPO_SOG
        ' 4 IMMOBILE
        ' 5 TIPO_IMM
        ' 6 DIRITTO
        ' 7 TITOLO
        ' 8 NUMERATORE
        ' 9 DENOMINATO
        ' 10 REGIME
        ' 11 RIF_REGIME
        ' 12 GEN_VALIDA
        ' 13 GEN_NOTA
        ' 14 GEN_NUMERO
        ' 15 GEN_PROGRE
        ' 16 GEN_ANNO
        ' 17 GEN_REGIST
        ' 18 PARTITA
        ' 19 CON_VALIDA
        ' 20 CON_NOTA
        ' 21 CON_NUMERO
        ' 22 CON_PROGRE
        ' 23 CON_ANNO
        ' 24 CON_REGIST
        ' 25 MUTAZ_INIZ
        ' 26 MUTAZ_FINE
        ' 27 IDENTIFICA
        ' 28 FLAG_IMPOR

        For Each FileTitoli As IO.FileInfo In Folder.GetFiles("*.tit", IO.SearchOption.TopDirectoryOnly)
            Application.DoEvents()
            objReader = New System.IO.StreamReader(FileTitoli.FullName)
            Do While objReader.Peek() <> -1
                objReader.ReadLine()
                NumeroRighe += 1
            Loop

            objReader.Close()

            BarProgress.PositionMin = 1
            BarProgress.PositionMax = NumeroRighe
            BarProgress.Position = 1

            objReader = New System.IO.StreamReader(FileTitoli.FullName)

            Do While objReader.Peek() <> -1
                Application.DoEvents()
                ReadLineRows = Split(objReader.ReadLine, "|")

                Call InsertTitoli(dbConnection, ReadLineRows, cmd)

                IndexNumeroRighe += 1
                BarProgress.Position = IndexNumeroRighe
            Loop

            objReader.Close()
        Next
        cmd.Dispose()

        Return If(NumeroRighe > 0, 1, 0)
    End Function

    Private Sub InsertTitoli(Connect As OleDbConnection, ByVal Values() As String, cmd As OleDbCommand)


        Dim MyTransaction As OleDb.OleDbTransaction = Connect.BeginTransaction

        With cmd
            .CommandType = CommandType.Text
            .CommandText = "INSERT INTO CTTITOLA (codice,sezione,soggetto,tipo_sog,immobile,tipo_imm,diritto,titolo,numeratore,denominato,regime,rif_regime,gen_valida,
                                                  gen_nota,gen_numero,gen_progre,gen_anno,gen_regist,partita,con_valida,con_nota,con_numero,con_progre, con_anno,con_regist,
                                                  mutaz_iniz,mutaz_fine,identifica,flag_impor) " &
                                         "VALUES (@codice,@sezione,@soggetto,@tipo_sog,@immobile,@tipo_imm,@diritto,@titolo,@numeratore,@denominato,@regime,@rif_regime,@gen_valida,
                                                  @gen_nota,@gen_numero,@gen_progre,@gen_anno,@gen_regist,@partita,@con_valida,@con_nota,@con_numero,@con_progre, @con_anno,@con_regist,
                                                  @mutaz_iniz,@mutaz_fine,@identifica,@flag_impor)"
            .Transaction = MyTransaction


            .Parameters.Add("@codice", OleDbType.VarChar).Value = NoNullText(Values(0))
            .Parameters.Add("@sezione", OleDbType.VarChar).Value = NoNullText(Values(1))
            .Parameters.Add("@soggetto", OleDbType.Double).Value = NoNullText(Values(2))
            .Parameters.Add("@tipo_sog", OleDbType.VarChar).Value = NoNullText(Values(3))
            .Parameters.Add("@immobile", OleDbType.Double).Value = NoNullDouble(Values(4))
            .Parameters.Add("@tipo_imm", OleDbType.VarChar).Value = NoNullText(Values(5))
            .Parameters.Add("@diritto", OleDbType.VarChar).Value = NoNullText(Values(6))
            .Parameters.Add("@titolo", OleDbType.VarChar).Value = NoNullText(Values(7))
            .Parameters.Add("@numeratore", OleDbType.Double).Value = NoNullDouble(Values(8))
            .Parameters.Add("@denominato", OleDbType.Double).Value = NoNullDouble(Values(9))
            .Parameters.Add("@regime", OleDbType.VarChar).Value = NoNullText(Values(10))
            .Parameters.Add("@rif_regime", OleDbType.Double).Value = NoNullDouble(Values(11))
            .Parameters.Add("@gen_valida", OleDbType.VarChar).Value = NoNullText(Values(12))
            .Parameters.Add("@gen_nota", OleDbType.VarChar).Value = NoNullText(Values(13))
            .Parameters.Add("@gen_numero", OleDbType.VarChar).Value = NoNullText(Values(14))
            .Parameters.Add("@gen_progre", OleDbType.VarChar).Value = NoNullText(Values(15))
            .Parameters.Add("@gen_anno", OleDbType.VarChar).Value = NoNullText(Values(16))
            .Parameters.Add("@gen_regist", OleDbType.VarChar).Value = NoNullText(Values(17))
            .Parameters.Add("@partita", OleDbType.VarChar).Value = NoNullText(Values(18))
            .Parameters.Add("@con_valida", OleDbType.VarChar).Value = NoNullText(Values(19))
            .Parameters.Add("@con_nota", OleDbType.VarChar).Value = NoNullText(Values(20))
            .Parameters.Add("@con_numero", OleDbType.VarChar).Value = NoNullText(Values(21))
            .Parameters.Add("@con_progre", OleDbType.VarChar).Value = NoNullText(Values(22))
            .Parameters.Add("@con_anno", OleDbType.VarChar).Value = NoNullText(Values(23))
            .Parameters.Add("@con_regist", OleDbType.VarChar).Value = NoNullText(Values(24))
            .Parameters.Add("@muta_iniz", OleDbType.Double).Value = NoNullDouble(Values(25))
            .Parameters.Add("@mutaz_fine", OleDbType.Double).Value = NoNullDouble(Values(26))
            .Parameters.Add("@identifica", OleDbType.Double).Value = NoNullDouble(Values(27))
            .Parameters.Add("@flag_impor", OleDbType.VarChar).Value = NoNullText(Values(28))

        End With

        Try
            cmd.Connection = Connect
            cmd.ExecuteNonQuery()
            MyTransaction.Commit()

        Catch ex As Exception
            MsgBox(ex.Message)
            MyTransaction.Rollback()

        Finally
            cmd.Dispose()
        End Try

    End Sub



End Module
