﻿Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Globalization
Imports System.Threading
Imports Skompare.SkompareMain

Public Class FormSkompare

    Dim skompareMain = New SkompareMain

    'Vypisuje data o jednotlivých sešitech do textového pole ve formuláři
    Private Sub ButtonStats(sender As Object, e As EventArgs) Handles BtnStats.Click

        'Vyčištění textboxu
        TBoxStats.Clear()

        'Deklarace pole parametrů
        Dim statsArr()() As String

        'Zkontroluje, zda je vybráno
        If LBoxNewSheets.SelectedIndex >= 0 And
                LBoxOldSheets.SelectedIndex >= 0 Then

            'Získání parametrů (názvy, řádky, sloupce) vybraných listů
            Try
                statsArr = skompareMain.GetSheetParams(LBoxNewSheets.GetItemText(LBoxNewSheets.SelectedItem),
                                                        LBoxOldSheets.GetItemText(LBoxOldSheets.SelectedItem))

            Catch ex As Exception When TypeOf ex Is NullReferenceException _
                                OrElse TypeOf ex Is System.Runtime.InteropServices.COMException
                MsgBox("Select sheets in both workbooks")
                Exit Sub

            End Try

            'Vypsání parametrů do textboxu
            TBoxStats.AppendText("Sheet name:" _
                                + vbTab _
                                + LBoxOldSheets.GetItemText(LBoxOldSheets.SelectedItem) _
                                + vbTab _
                                + LBoxNewSheets.GetItemText(LBoxNewSheets.SelectedItem))
            TBoxStats.AppendText(Environment.NewLine _
                                + "Row count:" _
                                + vbTab _
                                + statsArr(0)(0) _
                                + vbTab _
                                + statsArr(0)(1))
            TBoxStats.AppendText(Environment.NewLine _
                                    + "Column count:" _
                                    + vbTab _
                                    + statsArr(1)(0) _
                                    + vbTab _
                                    + statsArr(1)(1))

        Else

            Throw New System.Exception("Worksheets not selected")
            MsgBox("Select worksheets in both workbooks")

        End If
    End Sub

    'Otevírá dialogové okno pro výběr porovnávaných sešitů
    Private Sub ButtonSelectFile(sender As Object, e As EventArgs) Handles BtnNew.Click, BtnOld.Click

        'Deklarace názvu/cesty nového sešitu
        Dim FileName As String
        'Deklarace listboxu a labelu pro výpis názvů
        Dim Lbox As Object
        Dim nameLbl As Object

        'Otevře dialogové okno pro výběr souboru
        OpenFDNew.Title = "Select file"
        OpenFDNew.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
        OpenFDNew.ShowDialog()

        'Získá cestu vybraného souboru jako String
        FileName = OpenFDNew.FileName

        'Spuštění otevírací funkce
        Try
            skompareMain.OpenExcel(FileName, sender)
        Catch ex As System.Runtime.InteropServices.COMException
            MsgBox("Nebyl vybrán soubor")
            Exit Sub
        End Try

        'Výběr, které objekty se upraví podle stisknutého tlačítka
        If sender Is BtnNew Then 'Stisknuto "nové" tlačítko
            nameLbl = LblNewFileName
            Lbox = LBoxNewSheets
            skompareMain.WriteFileData(skompareMain.NewWb, FileName, LBoxNewSheets, LblNewFileName)
        ElseIf sender Is BtnOld Then 'Stisknuto "staré" tlačítko
            nameLbl = LblOldFileName
            Lbox = LBoxOldSheets
            skompareMain.WriteFileData(skompareMain.OldWb, FileName, LBoxOldSheets, LblOldFileName)
        End If

    End Sub

    'Metoda po zmáčknutí tlačítka pro porovnání
    Private Sub BtnCompare(sender As Object, e As EventArgs) Handles BtnComp.Click

        Dim progressBar As New FormProgBar
        skompareMain.PrLbl = progressBar.LblProgBar
        skompareMain.PrBar = progressBar.ProgBar

        Trace.Listeners.Add(New TextWriterTraceListener("Debug.log", "myListener"))
        Trace.WriteLine("Starting comparing @ " + DateTime.Now.ToString())
        Trace.Indent()

        Try
            'Zákaz přepočítávání
            Trace.WriteLine("Disabling auto-update")
            skompareMain.autoUpdate(False)

            'Získání parametrů (názvy, řádky, sloupce) vybraných listů
            Trace.WriteLine("Getting sheets parameters")
            skompareMain.GetSheetParams(LBoxNewSheets.GetItemText(LBoxNewSheets.SelectedItem),
                                        LBoxOldSheets.GetItemText(LBoxOldSheets.SelectedItem))

            'Zobrazení formuláře s progress barem
            Trace.WriteLine("Showing progress bar")
            'Zobrazí formulář
            progressBar.Show()
            'Nastaví aktuální hodnotu baru na 1 - začátek
            progressBar.ProgBar.Value = 1
            'Nastaví maximum na počet řádků v "novém" listu
            progressBar.ProgBar.Maximum = skompareMain.NewRows
            'Přepíše label
            progressBar.LblProgBar.Text = "Starting"

            'Vytvoření souboru pro zápis výsledků
            progressBar.LblProgBar.Text = "Creating output"
            skompareMain.CreateResult()

            'Spuštění porovnávadla
            progressBar.LblProgBar.Text = "Starting comparison"
            Trace.WriteLine("Starting Comparison")
            skompareMain.Compare()

            'Uložení
            progressBar.LblProgBar.Text = "Saving"
            skompareMain.ResultWb.SaveAs(skompareMain.NewWb.Path & "\compared", FileFormat:=51)

            'Zavření progress baru
            FormProgBar.Hide()

            'Povolení přepočítávání
            Trace.WriteLine("Enabling auto-update")
            skompareMain.autoUpdate(True)

            'Zavření sešitů
            skompareMain.ResultWb.Close(SaveChanges:=True)

            MsgBox("All done")
            Trace.WriteLine("ALL DONE")

            'Přenese formulář do popředí
            Me.Activate()

            'Řešení různých výjimek
        Catch ex As Exception

            Trace.WriteLine(ex.StackTrace _
                            & Environment.NewLine _
                            & ex.Message)
            Trace.WriteLine(ex.InnerException)
            Trace.WriteLine(ex.TargetSite)
            Trace.WriteLine(ex.Source)
            Trace.WriteLine(ex.Data)

            'Nejsou vybrány oba sešity
            If TypeOf ex Is NullReferenceException _
                            OrElse TypeOf ex Is System.Runtime.InteropServices.COMException Then
                Trace.WriteLine("EXCEPTION: " & ex.Message)
                Trace.Flush()
                Exit Sub

                'Při chybě kvůli nepřepsání souboru
            ElseIf TypeOf ex Is System.Runtime.InteropServices.COMException Then
                MsgBox("Compared sheet will not be overwritten")
                Trace.WriteLine("EXCEPTION: " & ex.Message)
                Trace.Flush()
                Exit Sub

                'Ostatní výjimky
            Else

                MsgBox("Exception found: " & ex.Message)
                Trace.WriteLine("EXCEPTION: " & ex.Message)
                Trace.Flush()
                FormProgBar.Hide()
                Exit Sub

            End If

        End Try

        Trace.Unindent()
        Trace.WriteLine("Comparing ended")
        Trace.WriteLine("___________________________________________________")
        Trace.Flush()

    End Sub

    'Metoda při načtení hlavního formuláře
    Private Sub Skompare_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Vytvoření aplikace excel, se ktrou se dále bude pracovat
        skompareMain.XlApp = New Excel.Application

    End Sub

    'Metoda při ukončení hlavního formuláře
    Private Sub Skompare_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If MessageBox.Show("Close?", "Close", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

            'Některé můžou být zavřené
            Try
                skompareMain.OldWb.Close(SaveChanges:=False)
                skompareMain.NewWb.Close(SaveChanges:=False)
                skompareMain.ResultWb.Close(SaveChanges:=False)
            Catch ex As Exception
            End Try

            skompareMain.XlApp.Quit()

        Else
            e.Cancel = True
        End If
    End Sub

    'Mění zadanou barvu podle vepsané hodnoty do textboxu
    Private Sub TBoxColor_TextChanged(sender As Object, e As EventArgs) Handles TBoxColor.TextChanged

        'Rozhodí string na části podle delimiteru
        Dim Colors() As String = Split(TBoxColor.Text, Delimiter:=",")

        Try

            'Udělá z částí stringu integery pro barvy
            Dim red As Integer = Int(Colors(0))
            Dim green As Integer = Int(Colors(1))
            Dim blue As Integer = Int(Colors(2))

            Dim IsLowContrast As Boolean = False

            'Je dostatečný kontrast k černému textu?
            If (red < 200 And green < 200 And blue < 200) _
                Or (red < 150 And green < 150) Then
                IsLowContrast = True
            End If

            'Nastaví barvu textu
            If IsLowContrast Then
                TBoxColor.ForeColor = Color.White
            Else
                TBoxColor.ForeColor = Color.Black
            End If

            'Nastaví barvu pozadí
            TBoxColor.BackColor = Color.FromArgb(red, green, blue)

        Catch ex As Exception

        End Try

    End Sub

    'Výběr barvy z dialogu
    Private Sub BtnColor_Click(sender As Object, e As EventArgs) Handles BtnColor.Click

        If ColorDialog1.ShowDialog() = DialogResult.OK Then

            Dim Highlight As Color = ColorDialog1.Color

            'UZíská jednotlivé barvy
            Dim red As Integer = Highlight.R
            Dim green As Integer = Highlight.G
            Dim blue As Integer = Highlight.B

            Dim IsLowContrast As Boolean = False

            'Je dostatečný kontrast k černému textu?
            If (red < 200 And green < 200 And blue < 200) _
                Or (red < 150 And green < 150) Then
                IsLowContrast = True
            End If

            'Nastaví barvu textu
            If IsLowContrast Then
                TBoxColor.ForeColor = Color.White
            Else
                TBoxColor.ForeColor = Color.Black
            End If

            'Nastaví barvu pozadí
            TBoxColor.BackColor = Highlight
            'Vepíše RGB barvy
            TBoxColor.Text = red & "," & green & "," & blue

        End If

    End Sub

    'Přepínání jazyka UI
    Private Sub BtnLang_Click(sender As Object, e As EventArgs) Handles BtnLang.Click

        'Výběr aktuální nastavené kultury
        Select Case Thread.CurrentThread.CurrentUICulture.Name

            'Aktuálně čeština
            Case "cs-CZ"

                ' Nastaví UI culture na angličtinu (en-US).
                Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")

                'Aktuálně angličtina
            Case "en-US"

                ' Nastaví UI culture na češtinu (cs-CZ).
                Thread.CurrentThread.CurrentUICulture = New CultureInfo("cs-CZ")

        End Select

        'Reinicializuje formulář
        Me.Controls.Clear()
        InitializeComponent()

    End Sub

    'Najde velikost záhlaví a pozici UID kódu podle pojmenovaného rozsahu v sešitu
    Private Sub BtnGetStartPoint_Click(sender As Object, e As EventArgs) Handles BtnGetStartPoint.Click
        skompareMain.GetSheetParams(LBoxNewSheets.GetItemText(LBoxNewSheets.SelectedItem),
                                    LBoxOldSheets.GetItemText(LBoxOldSheets.SelectedItem))
        TBoxColSelect.Text = skompareMain.GetExcelColumnName(skompareMain.NewSheet.Range("UID").Column)
        TBoxStart.Text = skompareMain.NewSheet.Range("Header").Rows.Count + 1
    End Sub
End Class
