Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.IO
Imports System.Data.OleDb
Imports System.Reflection
Imports System.Drawing

Public Structure SetupDati
    Public Comune As String
    Public DataRichiesta As String
    Public DataEleborazione As String
    Public TipoEstrazione As String
    Public NumeroRecord As Double
    Public PrmArchivio As String

End Structure

Public Class ConvertCatasto

    Private SetupDatiFabbricati As New SetupDati
    Private SetupDatiTerreni As New SetupDati

    Private FileTERRENI As String = ""
        Private FileFABBRICATI As String = ""
        Private PathFileMDB As String = ""

        '   <Browsable(True)>
        '   Public FileMDB As String = ""

        <Browsable(True)>
        Public ActualFileMDB As String = ""
        <Browsable(True)>
        Public ActualPathModelloFile As String = ""
        <Browsable(True)>
        Public ActualPathFileVisure As String = ""
        <Browsable(True)>
        Public CreateCatastini As Boolean = False
        <Browsable(True)>
        Public NomeFileTextCatastino As String = ""
        <Browsable(True)>
        Public AvvenutaCreazioneFiles As Boolean = True

        Public Event ChiudiConversione As EventHandler(Of MsgConversione)


        Public Sub New()
            InitializeComponent()
        End Sub

#Region "  OnSizeChange  "

    Protected Overrides Sub OnSizeChanged(e As EventArgs)
        If Not Me.IsDisposed Then
            If Me.Height <> 194 Then Me.Height = 194
            If Me.Width <> 730 Then Me.Height = 730
            MyBase.OnSizeChanged(e)
        End If
    End Sub

    Public Overloads ReadOnly Property Size As Size
        Get
            Return New Size(730, 194)
        End Get
    End Property

    Protected Overrides ReadOnly Property DefaultSize() As Size
            Get
                Return New Size(730, 194)
            End Get
        End Property


#End Region


    Private Sub Btn_convert_Click(sender As Object, e As EventArgs) Handles btn_convert.Click

        PathFileMDB = My.Computer.FileSystem.SpecialDirectories.Temp & "\"

        If CType(sender, Button).Text = "Chiudi pannello" Then
            btn_convert.Text = "Conversione"
            RaiseEvent ChiudiConversione(Me, New MsgConversione With {.CheckErrore = False, .MsgErrore = ""})
            Exit Sub
        End If

        Using Catasto1
            With Catasto1
                .Filter = "File Terreni|*.ter"
                .DefaultExt = "ter"
                .Title = "Selezionare il file .ter , [TERRENI]"
            End With
            If Catasto1.ShowDialog = Windows.Forms.DialogResult.OK Then
                FileTERRENI = Catasto1.FileName 'Path.GetDirectoryName(Catasto1.FileName)

                Dim objReader As New System.IO.StreamReader(Path.GetDirectoryName(Catasto1.FileName) & "\" & Split(Path.GetFileNameWithoutExtension(Catasto1.FileName), "_").GetValue(0).ToString & ".prm")
                SetupDatiTerreni.Comune = Split(objReader.ReadLine, ":").GetValue(1).ToString.Trim
                SetupDatiTerreni.DataRichiesta = Split(objReader.ReadLine, ":").GetValue(1).ToString.Trim
                SetupDatiTerreni.DataEleborazione = Split(objReader.ReadLine, ":").GetValue(1).ToString.Trim
                SetupDatiTerreni.TipoEstrazione = Split(objReader.ReadLine, ":").GetValue(1).ToString.Trim
                objReader.ReadLine()
                SetupDatiTerreni.NumeroRecord = Convert.ToDouble(Split(objReader.ReadLine, ":").GetValue(1).ToString)
                SetupDatiTerreni.PrmArchivio = "Terreni"
                objReader.Close()
            Else
                RaiseEvent ChiudiConversione(Me, New MsgConversione With {.CheckErrore = True, .MsgErrore = "Non hai seleionato nessu file Terreni (.ter)"})
                Exit Sub
            End If
        End Using

        Using Catasto1
            With Catasto1
                .Filter = "File Fabbricati|*.fab"
                .DefaultExt = "fab"
                .Title = "Selezionare il file .fab , [URBANO]"
            End With

            If Catasto1.ShowDialog = Windows.Forms.DialogResult.OK Then
                FileFABBRICATI = Catasto1.FileName ' Path.GetDirectoryName(Catasto1.FileName)
                Dim objReader As New System.IO.StreamReader(Path.GetDirectoryName(FileFABBRICATI) & "\" & Split(Path.GetFileNameWithoutExtension(FileFABBRICATI), "_").GetValue(0).ToString & ".prm")
                SetupDatiFabbricati.Comune = Split(objReader.ReadLine, ":").GetValue(1).ToString.Trim
                SetupDatiFabbricati.DataRichiesta = Split(objReader.ReadLine, ":").GetValue(1).ToString.Trim
                SetupDatiFabbricati.DataEleborazione = Split(objReader.ReadLine, ":").GetValue(1).ToString.Trim
                SetupDatiFabbricati.TipoEstrazione = Split(objReader.ReadLine, ":").GetValue(1).ToString.Trim
                objReader.ReadLine()
                SetupDatiFabbricati.NumeroRecord = Convert.ToDouble(Split(objReader.ReadLine, ":").GetValue(1).ToString)
                SetupDatiFabbricati.PrmArchivio = "Fabbricati"
                objReader.Close()
            Else
                RaiseEvent ChiudiConversione(Me, New MsgConversione With {.CheckErrore = True, .MsgErrore = "Non hai seleionato nessu file Fabbricati (.fab)"})
                Exit Sub
            End If
        End Using

        Dim CreateMDBFiles As New LibraryMDBcreate

        If CreateMDBFiles.CreateMDBfiles(Me.ActualPathModelloFile, ActualFileMDB) = True Then
            MsgBox("Creazione DB avvenuta con Successo...", MsgBoxStyle.Information, "Creazione db Visure")
            Me.AvvenutaCreazioneFiles = True
        Else
            MsgBox("Problemi con la creazione del DB visure...", MsgBoxStyle.Critical, "Errore Creazione db Visure")
            Me.AvvenutaCreazioneFiles = False
            RaiseEvent ChiudiConversione(Me, New MsgConversione With {.CheckErrore = True, .MsgErrore = "Mancata creazione del DataBase"})
            Exit Sub
        End If

        Dim Write_datiCatastali As String = "" ', NameField As New FieldCatastoInsert
        ProgressBar2.PositionMin = 1
        ProgressBar2.PositionMax = 7
        ProgressBar2.Text = ""

        Dim my_dbConnection As New OleDbConnection("Provider=microsoft.ace.oledb.12.0;Data Source=" & PathFileMDB & ActualFileMDB & ";Mode=Share Exclusive;")
        my_dbConnection.Open()

        Label_x_09.ForeColor = Color.Red
        Write_datiCatastali = Soggetti.WriteSoggetti(my_dbConnection, FileFABBRICATI, ProgressBar1)
        ProgressBar2.Position = 1
        Label_x_09.ForeColor = Color.Green
        label_09.ForeColor = Color.Green
        ProgressBar2.Refresh()

        Label_x_10.ForeColor = Color.Red
        Write_datiCatastali = Soggetti.WriteSoggetti(my_dbConnection, FileTERRENI, ProgressBar1)
        ProgressBar2.Position = 2
        Label_x_10.ForeColor = Color.Green
        label_10.ForeColor = Color.Green
        ProgressBar2.Refresh()

        Label_x_11.ForeColor = Color.Red
        Write_datiCatastali = Titoli.WriteTitoli(my_dbConnection, FileFABBRICATI, ProgressBar1)
        ProgressBar2.Position = 3
        Label_x_11.ForeColor = Color.Green
        label_11.ForeColor = Color.Green
        ProgressBar2.Refresh()

        Label_x_12.ForeColor = Color.Red
        Write_datiCatastali = Titoli.WriteTitoli(my_dbConnection, FileTERRENI, ProgressBar1)
        ProgressBar2.Position = 4
        Label_x_12.ForeColor = Color.Green
        label_12.ForeColor = Color.Green
        ProgressBar2.Refresh()

        Label_x_13.ForeColor = Color.Red
        Write_datiCatastali = StatusTerreni.CatastoTerreni(my_dbConnection, FileTERRENI, ProgressBar1)
        ProgressBar2.Position = 5
        Label_x_13.ForeColor = Color.Green
        label_13.ForeColor = Color.Green
        ProgressBar2.Refresh()

        Label_x_14.ForeColor = Color.Red
        Write_datiCatastali = StatusFabbricati.Fabbricati(my_dbConnection, FileFABBRICATI, ProgressBar1)
        ProgressBar2.Position = 6
        Label_x_14.ForeColor = Color.Green
        label_14.ForeColor = Color.Green
        ProgressBar2.Refresh()

        Label_x_15.ForeColor = Color.Red
        Write_datiCatastali = ChartTipologie.CreateChart(my_dbConnection, PathFileMDB & ActualFileMDB, ProgressBar1)

        ProgressBar2.Position = 7
        ProgressBar2.Refresh()
        Label_x_15.ForeColor = Color.Green
        label_15.ForeColor = Color.Green

        Write_datiCatastali = ArchivioDati(my_dbConnection, SetupDatiFabbricati, SetupDatiTerreni)

        my_dbConnection.Close()
        my_dbConnection.Dispose()

        If File.Exists(Me.ActualPathFileVisure & ActualFileMDB) = True Then
            File.Delete(Me.ActualPathFileVisure & ActualFileMDB)
        End If

        File.Copy(My.Computer.FileSystem.SpecialDirectories.Temp & "\" & ActualFileMDB, Me.ActualPathFileVisure & ActualFileMDB)
        btn_convert.Text = "Chiudi pannello"

    End Sub

End Class



    Public Class MsgConversione
    Inherits EventArgs

    Private _msgErrore As String
    Private _checkerrore As Boolean

    Public Sub New()
    End Sub

    Public Property MsgErrore As String
        Get
            Return _msgErrore
        End Get
        Set(value As String)
            _msgErrore = value
        End Set
    End Property

    Public Property CheckErrore As Boolean
        Get
            Return _checkerrore
        End Get
        Set(value As Boolean)
            _checkerrore = value
        End Set
    End Property

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class