
Imports gpe_convert_catasto

Public Class Form1

    Public Sub New()


        ' La chiamata è richiesta dalla finestra di progettazione.
        InitializeComponent()

        ConvertCatasto1.ActualFileMDB = "visure_catasto.accdb"
        ConvertCatasto1.ActualPathFileVisure = My.Computer.FileSystem.SpecialDirectories.Desktop & "\"
        ConvertCatasto1.ActualPathModelloFile = "C:\Users\massimo\Desktop\MDB\addon\"

        ConvertCatasto1.CreateCatastini = False
        ConvertCatasto1.NomeFileTextCatastino = ""

    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
    End Sub


    Private Sub ConvertCatasto1_ChiudiConversione(sender As Object, e As MsgConversione) Handles ConvertCatasto1.ChiudiConversione
        If e.CheckErrore = True Then
            MsgBox(e.MsgErrore)
        End If
        Me.Close()

    End Sub

End Class
