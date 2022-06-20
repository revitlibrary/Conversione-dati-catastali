Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.IO


Public Module Soggetti
    ''' <summary>Initializes new instance create WriteSoggetti.</summary>
    ''' <param name="dbConnection">Connessione del DB</param>
    ''' <param name="ConversioneFile">Cartella del file.SOG</param>
    ''' <param name="BarProgress">BarProgress view ProgressBar</param>
    ''' <returns>Value True False and numero soggetti trovati</returns>
    Friend Function WriteSoggetti(ByVal dbConnection As OleDbConnection, ByVal ConversioneFile As String, ByRef BarProgress As GpeProgressBar) As String

        Dim ReadLineRows() As String
        Dim objReader As System.IO.StreamReader, NumeroRighe As Double = 0, IndexNumeroRighe As Double = 0
        Dim NumSoggettiA As Double = 0, NumSoggettiB As Double = 0
        Dim Folder As New IO.DirectoryInfo(Path.GetDirectoryName(ConversioneFile))

        '-------------------- SOG -----------------------------------------------------------
        Dim cmd As New OleDbCommand

        For Each FileSoggettiUrbano As IO.FileInfo In Folder.GetFiles("*.sog", IO.SearchOption.TopDirectoryOnly)
            objReader = New System.IO.StreamReader(FileSoggettiUrbano.FullName) 'PathFolderFabbricati

            Do While objReader.Peek() <> -1
                objReader.ReadLine()
                NumeroRighe += 1
            Loop

            BarProgress.PositionMin = 1
            BarProgress.PositionMax = NumeroRighe
            BarProgress.Position = 1
            objReader.Close()

            objReader = New System.IO.StreamReader(FileSoggettiUrbano.FullName)

            Do While objReader.Peek() <> -1
                Application.DoEvents()
                ReadLineRows = Split(objReader.ReadLine, "|")

                Select Case ReadLineRows(3)
                    Case Is = "P"
                        Call Soggetti.InsertSoggettoP(dbConnection, ReadLineRows, cmd)
                        NumSoggettiA += 1
                        Exit Select

                    Case Else
                        Call Soggetti.InsertSoggettoG(dbConnection, ReadLineRows, cmd)
                        NumSoggettiB += 1
                        Exit Select

                End Select

                IndexNumeroRighe += 1
                BarProgress.Position = IndexNumeroRighe

            Loop
            objReader.Close()
        Next

        cmd.Dispose()

        Return "Persone fisiche:" & NumSoggettiA & " Soggetti Giuridici:" & NumSoggettiB

    End Function


    Private Sub InsertSoggettoG(Connect As OleDbConnection, ByVal Values() As String, cmd As OleDbCommand)


        Dim MyTransaction As OleDb.OleDbTransaction = Connect.BeginTransaction

        With cmd
            .CommandType = CommandType.Text
            .CommandText = "INSERT INTO ctnonfis  (codice,sezione,soggetto,tipo_sog,denominaz,sede,codfiscale) " &
                                          "VALUES (@codice,@sezione,@soggetto,@tipo_sog,@denominaz,@sede,@codfiscale)"
            .Transaction = MyTransaction
            .Parameters.Add("@codice", OleDbType.VarChar).Value = Values(0)
            .Parameters.Add("@sezione", OleDbType.VarChar).Value = If(Values(1).Trim = "", DBNull.Value, Values(1).Trim)
            .Parameters.Add("@soggetto", OleDbType.Double).Value = CDbl(Values(2))
            .Parameters.Add("@tipo_sog", OleDbType.VarChar).Value = Values(3)
            .Parameters.Add("@denominaz", OleDbType.VarChar).Value = Values(4)
            .Parameters.Add("@sede", OleDbType.VarChar).Value = Values(5)
            .Parameters.Add("@codfiscale", OleDbType.VarChar).Value = If(Values(6).Trim = "", DBNull.Value, Values(6).Trim)
        End With

        Try
            cmd.Connection = Connect
            cmd.ExecuteNonQuery()
            MyTransaction.Commit()

        Catch ex As Exception
            '   MsgBox(ex.Message)
            MyTransaction.Rollback()

        Finally
            cmd.Dispose()
        End Try

    End Sub

    Private Sub InsertSoggettoP(Connect As OleDbConnection, ByVal Values() As String, cmd As OleDbCommand)

        Dim MyTransaction As OleDb.OleDbTransaction = Connect.BeginTransaction

        With cmd
            .CommandType = CommandType.Text
            .Transaction = MyTransaction
            .CommandText = "INSERT INTO ctfisica (codice, sezione, soggetto, tipo_sog, cognome, nome, sesso,data, luogo, codfiscale, supplement) " &
                                         "VALUES (@codice,@sezione,@soggetto,@tipo_sog,@cognome,@nome,@sesso,@data,@luogo,@codfiscale,@supplement)"
            .Parameters.Add("@codice", OleDbType.VarChar).Value = Values(0)
            .Parameters.Add("@sezione", OleDbType.VarChar).Value = If(Values(1).Trim = "", DBNull.Value, Values(1).Trim)
            .Parameters.Add("@soggetto", OleDbType.Double).Value = CDbl(Values(2))
            .Parameters.Add("@tipo_sog", OleDbType.VarChar).Value = Values(3)
            .Parameters.Add("@cognome", OleDbType.VarChar).Value = Values(4)
            .Parameters.Add("@nome", OleDbType.VarChar).Value = Values(5)
            .Parameters.Add("@sesso", OleDbType.VarChar).Value = Values(6)
            .Parameters.Add("@data", OleDbType.DBDate).Value = If(IsDate(Values(7).Trim) = False, DBNull.Value, Convert.ToDateTime(Values(7).Trim))
            .Parameters.Add("@luogo", OleDbType.VarChar).Value = If(Values(8).Trim = "", DBNull.Value, Values(8).Trim)
            .Parameters.Add("@codfiscale", OleDbType.VarChar).Value = If(Values(9).Trim = "", DBNull.Value, Values(9).Trim)
            .Parameters.Add("@supplement", OleDbType.VarChar).Value = If(Values(10).Trim = "", DBNull.Value, Values(10).Trim)
        End With

        Try
            cmd.Connection = Connect
            cmd.ExecuteNonQuery()
            MyTransaction.Commit()

        Catch ex As Exception
            '  MsgBox(ex.Message)
            MyTransaction.Rollback()

        Finally
            cmd.Dispose()
        End Try
    End Sub




End Module
