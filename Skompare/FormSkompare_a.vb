Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Globalization
Imports System.Threading
Imports Skompare.SkompareMain

Public Class FormSkompare

    'Pro kontrolu viditelnosti panelu "advanced/pokročilé"
    Dim advancedVisibility As Boolean

    ' Vytvoří instanci třídy SkompareMain
    Dim skompareMain = New SkompareMain

    Private Sub BtnTest_Click(sender As Object, e As EventArgs) Handles BtnTest.Click
        skompareMain.AssignSheetsParams(CBoxNewSheets.SelectedItem,
                                        CBoxOldSheets.SelectedItem)
    End Sub


    'Vypisuje data o jednotlivých sešitech do textového pole ve formuláři
    Private Sub ButtonStats(sender As Object, e As EventArgs) Handles BtnStats.Click

        skompareMain.ShowMainParams(TBoxStats)

    End Sub

    'Otevírá dialogové okno pro výběr porovnávaných sešitů
    Private Sub ButtonSelectFile(sender As Object, e As EventArgs) Handles BtnNew.Click, BtnOld.Click

        skompareMain.OpenWorkbook(sender)

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
            skompareMain.GetSheetParams(CBoxNewSheets.GetItemText(CBoxNewSheets.SelectedItem),
                                        CBoxOldSheets.GetItemText(CBoxOldSheets.SelectedItem))

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

            'Zavření progress baru
            FormProgBar.Hide()

            'Povolení přepočítávání
            Trace.WriteLine("Enabling auto-update")
            skompareMain.autoUpdate(True)

            'Zavření sešitů
            skompareMain.OldWb.Close(SaveChanges:=False)
            skompareMain.NewWb.Close(SaveChanges:=False)
            skompareMain.xlApp.visible = True


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
        skompareMain.Application = New Excel.Application

        'Schování panelu "advanced"
        advancedVisibility = False
        PanelBottom.Visible = advancedVisibility
        MyBase.Height -= PanelBottom.Height

    End Sub

    'Metoda při ukončení hlavního formuláře
    Private Sub Skompare_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If MessageBox.Show("Close?", "Close", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

            'Některé můžou být zavřené
            Try
                skompareMain.ResultWb.Close(SaveChanges:=False)
            Catch ex As Exception
            End Try

            skompareMain.Application.Quit()

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
        skompareMain.GetSheetParams(CBoxNewSheets.GetItemText(CBoxNewSheets.SelectedItem),
                                    CBoxOldSheets.GetItemText(CBoxOldSheets.SelectedItem))
        TBoxColSelect1.Text = skompareMain.GetExcelColumnName(skompareMain.NewSheet.Range("UID").Column)
        TBoxStart.Text = skompareMain.NewSheet.Range("Header").Rows.Count + 1
    End Sub

    'Schovává/zobrazuje panel "advanced/pokročilé" pro detailnější nastavení
    Private Sub BtnAdvanced_Click(sender As Object, e As EventArgs) Handles BtnAdvanced.Click

        If advancedVisibility Then
            'Schování panelu "advanced"
            advancedVisibility = False
            PanelBottom.Visible = advancedVisibility
            MyBase.Height -= PanelBottom.Height
        Else
            'Zobrazení panelu "advanced"
            advancedVisibility = True
            PanelBottom.Visible = advancedVisibility
            MyBase.Height += PanelBottom.Height
        End If

    End Sub

End Class

