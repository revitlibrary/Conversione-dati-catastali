Imports System.Data.OleDb
Imports System.IO
Imports System.Windows.Forms



Module DatiArchivio

    Friend Function ArchivioDati(ByVal dbConnection As OleDbConnection, ByVal ArchivioTerreni As SetupDati, ArchivioFabbricati As SetupDati) As String


        Call WriteDati(dbConnection, ArchivioTerreni)
        Call WriteDati(dbConnection, ArchivioFabbricati)

        Return ""
    End Function


    Private Sub WriteDati(Connect As OleDbConnection, ByVal xArchivio As SetupDati)

        Dim cmd As New OleDbCommand
        Dim MyTransaction As OleDb.OleDbTransaction = Connect.BeginTransaction

        With cmd
            .CommandType = CommandType.Text
            .CommandText = "INSERT INTO datacrea (comune,datarichiesta,dataelaborazione,tipoestrazione,numrecord,prmarchivio) " &
                                         "VALUES (@comune,@datarichiesta,@dataelaborazione,@tipoestrazione,@numrecord,@prmarchivio)"
            .Transaction = MyTransaction
            .Parameters.Add("@comune", OleDbType.VarChar).Value = NoNullText(xArchivio.Comune)
            .Parameters.Add("@datarichiesta", OleDbType.VarChar).Value = NoNullText(xArchivio.DataRichiesta)
            .Parameters.Add("@dataelaborazione", OleDbType.VarChar).Value = NoNullText(xArchivio.DataEleborazione)
            .Parameters.Add("@tipoestrazione", OleDbType.VarChar).Value = NoNullText(xArchivio.TipoEstrazione)
            .Parameters.Add("@numrecord", OleDbType.Double).Value = NoNullDouble(xArchivio.NumeroRecord)
            .Parameters.Add("@prmarchivio", OleDbType.VarChar).Value = NoNullText(xArchivio.PrmArchivio)

        End With

        MsgBox("3")
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
