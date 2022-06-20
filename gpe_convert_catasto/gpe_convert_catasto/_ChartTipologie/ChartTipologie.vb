Imports System.Data.OleDb
Imports System.IO
Imports System.Windows.Forms

Module ChartTipologie

    Friend Function CreateChart(ByVal dbConnection As OleDbConnection, ByVal ConversioneFile As String, ByRef BarProgress As GpeProgressBar) As String


        Dim Folder As New IO.DirectoryInfo(Path.GetDirectoryName(ConversioneFile))

        Dim Sqls As String = "TRANSFORM Sum(CUIDENTI.FP_INDEX) AS  TOTALES SELECT CUIDENTI.FP, Count(CUARCUIU.CATEGORIA) AS TOTALE " &
                             "FROM CUIDENTI INNER JOIN CUARCUIU ON CUIDENTI.IMMOBILE = CUARCUIU.IMMOBILE GROUP BY CUIDENTI.FP PIVOT CUARCUIU.CATEGORIA;"

        Dim TmpValues As String = "", CreateTable As String = ""
        Dim cmd_count As New OleDb.OleDbCommand("SELECT COUNT(*) FROM CUIDENTI;", dbConnection), IndexRows As Double = 0
        BarProgress.PositionMin = 0
        BarProgress.PositionMax = cmd_count.ExecuteScalar
        cmd_count.Dispose()

        Dim cmd As New OleDb.OleDbCommand(Sqls, dbConnection) With {
            .Connection = dbConnection,
            .CommandText = Sqls
        }, dr As OleDb.OleDbDataReader, SqlTemp As String = ""
        dr = cmd.ExecuteReader

        dr.Read()

        cmd.Dispose()

        For jj As Integer = 2 To dr.FieldCount - 1
            SqlTemp += ", [" & dr.GetName(jj).ToString & "]"
            TmpValues += "," & (Val(dr(jj).ToString)).ToString
            CreateTable += ", [" & dr.GetName(jj).ToString & "] Integer "
        Next

        With dbConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, "CHART", "TABLE"})
            Using cb As New OleDb.OleDbCommand
                cb.Connection = dbConnection
                If .Rows.Count > 0 Then
                    cb.CommandText = "DROP TABLE CHART"
                    cb.ExecuteNonQuery()
                End If
                cb.CommandText = "CREATE TABLE CHART ([FP] Char(11), [NUM] Integer " & CreateTable & ")"
                cb.ExecuteNonQuery()
            End Using
        End With

        Dim cmdwrite As New OleDbCommand("INSERT INTO [CHART] ([FP], [NUM]" & SqlTemp & ") 
                                                VALUES ( '" & dr(0).ToString & "'," &
                                                             DbValue((Val(dr(1).ToString)).ToString) & TmpValues & ")", dbConnection)
        cmdwrite.ExecuteNonQuery()
        cmdwrite.Dispose()

        While dr.Read
            Application.DoEvents()
            TmpValues = ""
            For jj As Integer = 2 To dr.FieldCount - 1
                TmpValues = TmpValues & "," & DbValue((Val(dr(jj).ToString)).ToString)
            Next
            cmdwrite = New OleDbCommand("INSERT INTO [CHART] ([FP], [NUM]" & SqlTemp & ") 

                                        VALUES ( " & ApiciSI(dr(0).ToString) & "," & DbValue((Val(dr(1).ToString)).ToString) & TmpValues & ")", dbConnection)
            cmdwrite.ExecuteNonQuery()
            IndexRows += 1
            BarProgress.Position = IndexRows
            cmdwrite.Dispose()
        End While
        BarProgress.Position = BarProgress.PositionMax
        dr.Close()

        cmdwrite.Dispose()
        Return ""
    End Function


End Module
