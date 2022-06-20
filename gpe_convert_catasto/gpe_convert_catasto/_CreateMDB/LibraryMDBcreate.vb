Imports System.IO
Imports System.Data.OleDb

Friend Class LibraryMDBcreate

    Public Sub New()
    End Sub



    ''' <summary>Initializes new instance create MDB files.</summary>
    ''' <returns>Value True False is create MDB file</returns>
    Friend Function CreateMDBfiles(ByVal PathFileVisure As String, ByVal FileDataBase As String) As Boolean

        Dim _temp_value As Boolean = False

        If File.Exists(My.Computer.FileSystem.SpecialDirectories.Temp & "\" & FileDataBase) Then
            File.Delete(My.Computer.FileSystem.SpecialDirectories.Temp & "\" & FileDataBase)
        End If

        System.IO.File.Copy(PathFileVisure & "BaseTemplate.accdb", My.Computer.FileSystem.SpecialDirectories.Temp & "\" & FileDataBase, True)
        If File.Exists(My.Computer.FileSystem.SpecialDirectories.Temp & "\" & FileDataBase) = True Then
            _temp_value = CreateTable(My.Computer.FileSystem.SpecialDirectories.Temp & "\" & FileDataBase)
        End If

        Return _temp_value

    End Function

    Private Function CreateTable(ByVal FileDataBase As String) As Boolean
        Dim CreateTableCatasto As New DatiTabelle, _temp_value As Boolean = False ', _temp_CreateTableCatasto As String = ""
        Dim my_dbConnection As New OleDbConnection("Provider=microsoft.ace.oledb.12.0; Data Source=" & FileDataBase)
        Dim my_Command As New OleDbCommand

        my_dbConnection.Open()
        Try
            For Each Sqls As String In CreateTableCatasto.DatiTABELLE
                my_Command = New OleDbCommand(Sqls, my_dbConnection)
                my_Command.ExecuteNonQuery()
                _temp_value = True
            Next

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            my_Command.Dispose()
            my_dbConnection.Close()
        End Try


        Return _temp_value
    End Function



End Class
