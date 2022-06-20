Imports System.Data.OleDb
Imports System.IO
Imports System.Windows.Forms


Module StatusTerreni

    Friend Function CatastoTerreni(ByVal dbConnection As OleDbConnection, ByVal ConversioneFile As String, ByRef BarProgress As GpeProgressBar) As String

        Dim ReadLineRows() As String
        Dim objReader As System.IO.StreamReader, NumeroRighe As Double = 0, IndexNumeroRighe As Double = 0
        Dim Folder As New IO.DirectoryInfo(Path.GetDirectoryName(ConversioneFile))
        Dim cmd As New OleDbCommand

        For Each FileTitoli As IO.FileInfo In Folder.GetFiles("*.ter", IO.SearchOption.TopDirectoryOnly)

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
                    Case Is = 1 'Tipo record 1 (Terreno)

                        Call StatusTerreni.Particelle(dbConnection, ReadLineRows, cmd)
                        Exit Select

                    Case Is = 2 'Tipo record 2 (Deduzioni) 
                        Call StatusTerreni.Deduzioni(dbConnection, ReadLineRows, cmd)
                        Exit Select

                    Case Is = 3 'Tipo record 3 (Riserve) 
                        Call StatusTerreni.Riserve(dbConnection, ReadLineRows, cmd)
                        Exit Select

                    Case Is = 4 'Tipo record 4 (Porzioni) 12
                        Call StatusTerreni.Porzioni(dbConnection, ReadLineRows, cmd)
                End Select


                Application.DoEvents()
                IndexNumeroRighe += 1
                BarProgress.Position = IndexNumeroRighe
            Loop

            objReader.Close()
        Next
        cmd.Dispose()
        Return NumeroRighe - IndexNumeroRighe


    End Function


    ''' <summary>Create Identificativi. "CTPARTIC"</summary>
    ''' <param name="Connect">Connessione del DB</param>
    ''' <param name="Values">Array contenente i dati da memorizzare</param>
    ''' <param name="cmd">paramertro OleDbCommand</param>
    Private Sub Particelle(Connect As OleDbConnection, ByVal Values() As String, cmd As OleDbCommand)

        Dim MyTransaction As OleDb.OleDbTransaction = Connect.BeginTransaction

        With cmd
            .CommandType = CommandType.Text
            .CommandText = "INSERT INTO ctpartic (codice,sezione,immobile,tipo_imm,progressiv,foglio,particella,denominato,subalterno,edificiale,qualita,classe,ettari,
                                                  [are],centiare,flag_redd,flag_porz,flag_deduz,dominic_l,agrario_l,dominic_e,agrario_e,gen_eff,gen_regist,gen_tipo,
                                                  gen_numero,gen_progre,gen_anno,con_eff,con_regist,con_tipo,con_numero,con_progre,con_anno,partita,annotazion,
                                                  mutaz_iniz,mutaz_fine)" &
                                         "VALUES (@codice,@sezione,@immobile,@tipo_imm,@progressiv,@foglio,@particella,@denominato,@subalterno,@edificiale,@qualita,
                                                  @classe,@ettari,[@are],@centiare,@flag_redd,@flag_porz,@flag_deduz,@dominic_l,@agrario_l,@dominic_e,@agrario_e,
                                                  @gen_eff,@gen_regist,@gen_tipo,@gen_numero,@gen_progre,@gen_anno,@con_eff,@con_regist,@con_tipo,@con_numero,
                                                  @con_progre,@con_anno,@partita,@annotazion,@mutaz_iniz,@mutaz_fine)"

            .Transaction = MyTransaction
            .Parameters.Add("@codice", OleDbType.VarChar).Value = NoNullText(Values(0))
            .Parameters.Add("@sezione", OleDbType.VarChar).Value = NoNullText(Values(1))
            .Parameters.Add("@immobile", OleDbType.Double).Value = NoNullDouble(Values(2))
            .Parameters.Add("@tipo_imm", OleDbType.VarChar).Value = NoNullText(Values(3))
            .Parameters.Add("@progressiv", OleDbType.Double).Value = NoNullDouble(Values(4))
            .Parameters.Add("@foglio", OleDbType.Double).Value = NoNullDouble(Values(6))
            .Parameters.Add("@particella", OleDbType.VarChar).Value = NoNullText(Values(7))
            .Parameters.Add("@denominato", OleDbType.Double).Value = NoNullDouble(Values(8))
            .Parameters.Add("@subalterno", OleDbType.Double).Value = NoNullDouble(Values(9))
            .Parameters.Add("@edificiale", OleDbType.VarChar).Value = NoNullText(Values(10))
            .Parameters.Add("@qualita", OleDbType.VarChar).Value = NoNullText(Values(11))
            .Parameters.Add("@classe", OleDbType.VarChar).Value = NoNullText(Values(12))
            .Parameters.Add("@ettari", OleDbType.Double).Value = NoNullDouble(Values(13))
            .Parameters.Add("[@are]", OleDbType.Double).Value = NoNullDouble(Values(14))
            .Parameters.Add("@centiare", OleDbType.Double).Value = NoNullDouble(Values(15))
            .Parameters.Add("@flag_redd", OleDbType.VarChar).Value = NoNullText(Values(16))
            .Parameters.Add("@flag_porz", OleDbType.VarChar).Value = NoNullText(Values(17))
            .Parameters.Add("@flag_deduz", OleDbType.VarChar).Value = NoNullText(Values(18))
            .Parameters.Add("@dominic_l", OleDbType.VarChar).Value = NoNullText(Values(19))
            .Parameters.Add("@agrario_l", OleDbType.VarChar).Value = NoNullText(Values(20))
            .Parameters.Add("@dominic_e", OleDbType.VarChar).Value = NoNullText(Values(21))
            .Parameters.Add("@agrario_e", OleDbType.VarChar).Value = NoNullText(Values(22))
            .Parameters.Add("@gen_eff", OleDbType.VarChar).Value = NoNullText(Values(23))
            .Parameters.Add("@gen_regist", OleDbType.VarChar).Value = NoNullText(Values(24))
            .Parameters.Add("@gen_tipo", OleDbType.VarChar).Value = NoNullText(Values(25))
            .Parameters.Add("@gen_numero", OleDbType.VarChar).Value = NoNullText(Values(26))
            .Parameters.Add("@gen_progre", OleDbType.VarChar).Value = NoNullText(Values(27))
            .Parameters.Add("@gen_anno", OleDbType.Double).Value = NoNullDouble(Values(28))
            .Parameters.Add("@con_eff", OleDbType.VarChar).Value = NoNullText(Values(29))
            .Parameters.Add("@con_regist", OleDbType.VarChar).Value = NoNullText(Values(30))
            .Parameters.Add("@con_tipo", OleDbType.VarChar).Value = NoNullText(Values(31))
            .Parameters.Add("@con_numero", OleDbType.VarChar).Value = NoNullText(Values(32))
            .Parameters.Add("@con_progre", OleDbType.VarChar).Value = NoNullText(Values(33))
            .Parameters.Add("@con_anno", OleDbType.Double).Value = NoNullDouble(Values(34))
            .Parameters.Add("@partita", OleDbType.VarChar).Value = NoNullText(Values(35))
            .Parameters.Add("@annotazion", OleDbType.VarChar).Value = NoNullText(Values(36))
            .Parameters.Add("@mutaz_iniz", OleDbType.Double).Value = NoNullDouble(Values(37))
            .Parameters.Add("@mutaz_fine", OleDbType.Double).Value = NoNullDouble(Values(38))

        End With

        Try
            cmd.Connection = Connect
            cmd.ExecuteNonQuery()
            MyTransaction.Commit()

        Catch ex As Exception
            MsgBox(ex.Message & "<>" & "CTPARTIC")
            MyTransaction.Rollback()

        Finally
            cmd.Dispose()
        End Try

    End Sub

    ''' <summary>Create Identificativi. "CTDEDUZI"</summary>
    ''' <param name="Connect">Connessione del DB</param>
    ''' <param name="Values">Array contenente i dati da memorizzare</param>
    ''' <param name="cmd">paramertro OleDbCommand</param>
    Private Sub Deduzioni(Connect As OleDbConnection, ByVal Values() As String, cmd As OleDbCommand)

        Dim MyTransaction As OleDb.OleDbTransaction = Connect.BeginTransaction

        With cmd
            .CommandType = CommandType.Text
            .CommandText = "INSERT INTO CTDEDUZI (codice,sezione,immobile,tipo_imm,progressiv,deduzione,flag_impor)" &
                                         "VALUES (@codice,@sezione,@immobile,@tipo_imm,@progressiv,@deduzione,@flag_impor)"

            .Transaction = MyTransaction
            .Parameters.Add("@codice", OleDbType.VarChar).Value = NoNullText(Values(0))
            .Parameters.Add("@sezione", OleDbType.VarChar).Value = NoNullText(Values(1))
            .Parameters.Add("@immobile", OleDbType.Double).Value = NoNullDouble(Values(2))
            .Parameters.Add("@tipo_imm", OleDbType.VarChar).Value = NoNullText(Values(3))
            .Parameters.Add("@progressiv", OleDbType.Double).Value = NoNullDouble(Values(4))
            .Parameters.Add("@deduzion", OleDbType.VarChar).Value = NoNullText(Values(6))
            .Parameters.Add("@flag_impor", OleDbType.VarChar).Value = NoNullText(Values(7))
        End With

        Try
            cmd.Connection = Connect
            cmd.ExecuteNonQuery()
            MyTransaction.Commit()

        Catch ex As Exception
            MsgBox(ex.Message & "<>" & "CTDEDUZI")
            MyTransaction.Rollback()

        Finally
            cmd.Dispose()
        End Try

    End Sub


    ''' <summary>Create Identificativi. "CTRISERV"</summary>
    ''' <param name="Connect">Connessione del DB</param>
    ''' <param name="Values">Array contenente i dati da memorizzare</param>
    ''' <param name="cmd">paramertro OleDbCommand</param>
    Private Sub Riserve(Connect As OleDbConnection, ByVal Values() As String, cmd As OleDbCommand)

        Dim MyTransaction As OleDb.OleDbTransaction = Connect.BeginTransaction

        With cmd
            .CommandType = CommandType.Text
            .CommandText = "INSERT INTO CTRISERV (codice,sezione,immobile,tipo_imm,progressiv,riserva,iscrizione,flag_impor)" &
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
            MsgBox(ex.Message & "<>" & "CTRISERV")
            MyTransaction.Rollback()

        Finally
            cmd.Dispose()
        End Try

    End Sub

    ''' <summary>Create Identificativi. "CTPORZIO"</summary>
    ''' <param name="Connect">Connessione del DB</param>
    ''' <param name="Values">Array contenente i dati da memorizzare</param>
    ''' <param name="cmd">paramertro OleDbCommand</param>
    Private Sub Porzioni(Connect As OleDbConnection, ByVal Values() As String, cmd As OleDbCommand)

        Dim MyTransaction As OleDb.OleDbTransaction = Connect.BeginTransaction

        With cmd
            .CommandType = CommandType.Text
            .CommandText = "INSERT INTO CTPORZIO (codice,sezione,immobile,tipo_imm,progressiv,porzione,qualita,classe,ettari,[are],centiare,flag_impor)" &
                                         "VALUES (@codice,@sezione,@immobile,@tipo_imm,@progressiv,@porzione,@qualita,@classe,@ettari,[@are],@centiare,@flag_impor)"

            .Transaction = MyTransaction
            .Parameters.Add("@codice", OleDbType.VarChar).Value = NoNullText(Values(0))
            .Parameters.Add("@sezione", OleDbType.VarChar).Value = NoNullText(Values(1))
            .Parameters.Add("@immobile", OleDbType.Double).Value = NoNullDouble(Values(2))
            .Parameters.Add("@tipo_imm", OleDbType.VarChar).Value = NoNullText(Values(3))
            .Parameters.Add("@progressiv", OleDbType.Double).Value = NoNullDouble(Values(4))
            .Parameters.Add("@porzione", OleDbType.VarChar).Value = NoNullText(Values(6))
            .Parameters.Add("@qualita", OleDbType.Double).Value = NoNullDouble(Values(7))
            .Parameters.Add("@classe", OleDbType.VarChar).Value = NoNullText(Values(8))
            .Parameters.Add("@ettari", OleDbType.Double).Value = NoNullDouble(Values(9))
            .Parameters.Add("[@are]", OleDbType.Double).Value = NoNullDouble(Values(10))
            .Parameters.Add("@centiare", OleDbType.Double).Value = NoNullDouble(Values(11))
            .Parameters.Add("@flag_impor", OleDbType.VarChar).Value = NoNullText(Values(12))

        End With

        Try
            cmd.Connection = Connect
            cmd.ExecuteNonQuery()
            MyTransaction.Commit()

        Catch ex As Exception
            MsgBox(ex.Message & "<>" & "CTPORZIO")
            MyTransaction.Rollback()

        Finally
            cmd.Dispose()
        End Try

    End Sub






End Module
