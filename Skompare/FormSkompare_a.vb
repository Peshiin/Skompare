Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Globalization
Imports System.Threading

Public Class FormSkompare

    'Deklarace aplikace excel
    Dim XlApp As Excel.Application

    'Deklarace sešitů
    Dim NewWb As Excel.Workbook
    Dim OldWb As Excel.Workbook
    Dim ResultWb As Excel.Workbook

    'Deklarace listů
    Dim NewSheet As Excel.Worksheet
    Dim OldSheet As Excel.Worksheet
    Dim NewResSheet As Excel.Worksheet
    Dim OldResSheet As Excel.Worksheet

    'Deklarace parametrů vybraných listů
    'Počty řádků
    Dim NewRows As Int16
    Dim OldRows As Int16
    'Počty sloupců
    Dim NewCols As Int16
    Dim OldCols As Int16
    'Větší počet řádků
    Dim lenRows As Integer
    'Větší počet řádků
    Dim lenCols As Integer
    'Sloupec pro vyhledávání
    Dim ColLookup As Int16

    'Deklarace polí pro porovnání řádků
    Dim NewRowArr As Object(,)
    Dim OldRowArr As Object(,)

    'Deklarace proměnných pro ovládání progress baru a jeho popisku
    Dim PrBar As Object = FormProgBar.ProgBar
    Dim PrLbl As Object = FormProgBar.LblProgBar

    'Získává data o počtech řádků a sloupců v jednotlivých sešitech
    Private Sub GetSheetParams()

        'Zkontroluje, zda je vybráno
        If LBoxNewSheets.SelectedIndex >= 0 And
            LBoxOldSheets.SelectedIndex >= 0 Then

            'Definice listů pro porovnání podle vybraných položek ze seznamů
            NewSheet = NewWb.Worksheets(LBoxNewSheets.GetItemText(LBoxNewSheets.SelectedItem))
            OldSheet = OldWb.Worksheets(LBoxOldSheets.GetItemText(LBoxOldSheets.SelectedItem))

            'Získání parametrů listů
            ''přepočítání UsedRange
            'NewSheet.UsedRange.Calculate()
            'OldSheet.UsedRange.Calculate()
            'řádky
            NewRows = GetLast(NewSheet, order:=Excel.XlSearchOrder.xlByColumns).Row 'NewSheet.UsedRange.Rows.Count
            OldRows = GetLast(OldSheet, order:=Excel.XlSearchOrder.xlByColumns).Row 'OldSheet.UsedRange.Rows.Count
            'sloupce
            NewCols = GetLast(NewSheet, order:=Excel.XlSearchOrder.xlByRows).Column 'NewSheet.UsedRange.Columns.Count
            OldCols = GetLast(OldSheet, order:=Excel.XlSearchOrder.xlByRows).Column 'OldSheet.UsedRange.Columns.Count

        Else

            Throw New System.Exception("Worksheets not selected")
            MsgBox("Select worksheets in both workbooks")

        End If

    End Sub

    'Vypisuje data o jednotlivých sešitech do textového pole ve formuláři
    Private Sub ButtonStats(sender As Object, e As EventArgs) Handles BtnStats.Click
        'Vyčištění textboxu
        TBoxStats.Clear()

        'Získání parametrů (názvy, řádky, sloupce) vybraných listů
        Try
            GetSheetParams()

        Catch ex As Exception When TypeOf ex Is NullReferenceException _
                            OrElse TypeOf ex Is System.Runtime.InteropServices.COMException
            MsgBox("Select sheets in both workbooks")
            Exit Sub

        End Try

        'Vypsání parametrů do textboxu
        TBoxStats.AppendText("Sheet name:" _
                                + vbTab _
                                + OldSheet.Name _
                                + vbTab _
                                + NewSheet.Name)
        TBoxStats.AppendText(Environment.NewLine _
                                + "Row count:" _
                                + vbTab _
                                + OldRows.ToString _
                                + vbTab _
                                + NewRows.ToString)
        TBoxStats.AppendText(Environment.NewLine _
                                + "Column count:" _
                                + vbTab _
                                + OldCols.ToString _
                                + vbTab _
                                + NewCols.ToString)
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
            OpenExcel(FileName, sender)
        Catch ex As System.Runtime.InteropServices.COMException
            MsgBox("Nebyl vybrán soubor")
            Exit Sub
        End Try

        'Výběr, které objekty se upraví podle stisknutého tlačítka
        If sender Is BtnNew Then 'Stisknuto "nové" tlačítko
            nameLbl = LblNewFileName
            Lbox = LBoxNewSheets
            WriteFileData(NewWb, FileName, LBoxNewSheets, LblNewFileName)
        ElseIf sender Is BtnOld Then 'Stisknuto "staré" tlačítko
            nameLbl = LblOldFileName
            Lbox = LBoxOldSheets
            WriteFileData(OldWb, FileName, LBoxOldSheets, LblOldFileName)
        End If

    End Sub

    'Otevírá sešity pro porovnání
    Private Sub OpenExcel(FilePath As String, sender As Object)

        'Výběr, které objekty se upraví podle stisknutého tlačítka
        If sender Is BtnNew Then 'Stisknuto "nové" tlačítko

            'Otevření souboru v aplikaci Excel
            NewWb = XlApp.Workbooks.Open(FilePath, [ReadOnly]:=True)

        ElseIf sender Is BtnOld Then 'Stisknuto "staré" tlačítko

            'Otevření souboru v aplikaci Excel
            OldWb = XlApp.Workbooks.Open(FilePath, [ReadOnly]:=True)

        End If

    End Sub

    'Vypisuje listy sešitů do přehledového okénka
    Sub WriteFileData(Wb As Excel.Workbook, FileName As String, Lbox As Object, nameLbl As Object)

        'Vypsání názvu souboru do formuláře (Dir() vybere pouze název souboru a ne celou cestu)
        nameLbl.Text = Dir(FileName)

        'Vyčištění ListBoxu od popisku
        Lbox.Items.Clear()
        'Vypsání názvů listů ve vybraném sešitu
        For Each sheet In Wb.Worksheets
            Lbox.Items.Add(sheet.Name)
        Next

    End Sub

    'Rozhoduje, jak se vyznačí změna
    Private Sub CompareStyle(NewRng As Excel.Range, NewStr As String, OldStr As String)

        'Jen obarvení
        If RBtnStyle1.Checked Then
            NewRng.Interior.Color = TBoxColor.BackColor
            NewRng.Value = NewStr

            'Obarvení a komentář
        ElseIf RBtnStyle2.Checked Then
            NewRng.Interior.Color = TBoxColor.BackColor
            NewRng.Value = NewStr
            'Vyhazuje výjimku, pokud je komentář prázdný
            If OldStr = "" Then
                NewRng.AddComment("-")
            Else
                NewRng.AddComment(OldStr)
            End If

            'Obarvení a řetězec
        ElseIf RBtnStyle3.Checked Then
            NewRng.Interior.Color = TBoxColor.BackColor
            NewRng.Value = NewStr & " " & TBoxStringStart.Text & OldStr & TBoxStringEnd.Text

            'Jen komentář
        ElseIf RBtnStyle4.Checked Then
            NewRng.Value = NewStr
            'Vyhazuje výjimku, pokud je komentář prázdný
            If OldStr = "" Then
                NewRng.AddComment("-")
            Else
                NewRng.AddComment(OldStr)
            End If

            'Jen řetězec
        ElseIf RBtnStyle5.Checked Then
            NewRng.Value = NewStr & " " & TBoxStringStart.Text & OldStr & TBoxStringEnd.Text

            'Řetězec v komentáři
        ElseIf RBtnStyle6.Checked Then
            NewRng.Value = NewStr
            'Vyhazuje výjimku, pokud je komentář prázdný
            If OldStr = "" Then
                NewRng.AddComment("-")
            Else
                NewRng.AddComment(TBoxStringStart.Text & OldStr & TBoxStringEnd.Text)
            End If

        End If
    End Sub

    'Vyznačuje změny v řádku
    Sub CompareRow(NewA As Array, OldA As Array, NewR As Integer, OldR As Integer)

        'Deklarace pomocných proměnných
        Dim NewVal As String
        Dim OldVal As String

        With NewResSheet.Rows(NewR)

            For Val As Integer = 1 To Math.Min(NewCols, OldCols)

                NewVal = NewA.GetValue(NewR, Val)
                OldVal = OldA.GetValue(OldR, Val)

                If NewVal <> OldVal Then

                    CompareStyle(.Cells(1, Val), NewVal, OldVal)

                End If

            Next

        End With

    End Sub

    'Prochází sešity a porovnává řádky (vyznačení změn v řádku řeší samostatná funkce)
    Sub Compare()

        'Kde je více řádků/sloupců, podle toho se vezme délka pole
        Trace.WriteLine("Getting bigger dimension")
        PrLbl.Text = "Getting bigger dimension"
        lenRows = GetBiggerDim(NewRows, OldRows)
        lenCols = GetBiggerDim(NewCols, OldCols)

        'Vytvoří pole pro porovnávání
        Trace.WriteLine("Getting arrays")
        PrLbl.Text = "Getting arrays"
        NewSheet.UsedRange.Calculate()
        OldSheet.UsedRange.Calculate()
        Dim NewArr As Object(,) = CType(NewSheet.UsedRange.Value, Object(,))
        Dim OldArr As Object(,) = CType(OldSheet.UsedRange.Value, Object(,))

        'Získá číslo sloupce, podle kterého se bude hledat
        Trace.WriteLine("Getting key column")
        PrLbl.Text = "Getting key column"
        ColSelect(TBoxColSelect.Text)

        'Získání startovacího řádku
        PrLbl.Text = "Checking start row input"
        Trace.WriteLine("Checking start row input")
        Dim StartRow As Int16
        'Kontrola na integer
        If Integer.TryParse(TBoxStart.Text, StartRow) = False Then
            MsgBox("Start row entered is not integer")
            Exit Sub
        Else
            PrBar.Value = StartRow - 1
        End If

        'Deklarace pomocných proměnných
        'Trackování, zda byla nalezena shoda
        Dim MatchFound As Boolean
        'Hledaná hodnota (jedinečný kód)
        Dim SearchString As String
        'Index ve "starém" poli, kde je hledaná hodnota
        Dim OldRow As Integer
        'Pomocná proměnná pro hledání duplicit
        Dim i As Integer

        'Získání pole vyhledávaných indexů starého pole
        Dim OldIndArr() As String
        OldIndArr = GetIndArr(OldArr, ColLookup, OldRows)

        'Získání pole pro kontrolu duplicit (0 = index zatím nenalezen)
        Dim Duplicity(OldRows) As Integer
        For i = 1 To OldRows
            Duplicity(i) = 0
        Next

        'Prohledávací cyklus
        PrLbl.Text = "Starting looping"
        Trace.WriteLine("Starting looping")
        'Loop v "nových" datech
        For NewRow = StartRow To NewRows

            'Shoda nenalezena
            MatchFound = False

            'Hledaný jedinečný kód
            SearchString = NewArr(NewRow, ColLookup)

            'Vrátí polohu (řádek) hledaného kódu ve "starém" poli
            OldRow = Array.IndexOf(OldIndArr, SearchString)

            'Ignoruje první výskyt, pokud už byl zaznamenán (pokud jsou duplicitní kódy)
            If OldRow > 0 Then
                If Duplicity(OldRow) = 1 Then
                    OldRow = Array.IndexOf(OldIndArr, SearchString, OldRow + 1)
                End If
            End If

            'Nalezena shoda identifikátoru?
            If OldRow > 0 Then

                'Zaznamená nalezení shody
                MatchFound = True
                Duplicity(OldRow) = 1

                'Porovná buňky v řádku
                CompareRow(NewArr, OldArr, NewRow, OldRow)

            End If


            If MatchFound = False Then

                NewResSheet.Rows(NewRow).EntireRow.Interior.Color = TBoxColor.BackColor

            End If

            PrBar.Value += 1
            PrLbl.Text = "Progress: " _
                        & Math.Round((PrBar.Value - StartRow) / (NewRows - StartRow), 2) * 100 _
                        & "% (" & NewRow & " out of " & NewRows & ")"

        Next

        'Smaže nalezené (zeleně označené) řádky ve "zrušeném" listu
        PrLbl.Text = "Cleaning found rows from Cancelled"
        Trace.WriteLine("Cleaning found rows from Cancelled")
        DeleteRows(OldResSheet, Duplicity)

        'Nastavení zobrazení po dalším otevření sešitu (nebude najeto někam doprostřed listu a nastaví se scroll bar)
        Try
            OldResSheet.Activate()
            OldResSheet.Range("A1").Select()
        Catch ex As Exception
        End Try

    End Sub

    'Metoda pro vymazání nalezených (označeno zeleně) řádek ve "zrušeném" listu
    Sub DeleteRows(sheet As Excel.Worksheet, indexArray() As Integer)

        For i As Integer = indexArray.Length - 1 To TBoxStart.Text Step -1

            If indexArray(i) = 1 Then

                sheet.Rows(i).EntireRow.Delete

            End If

        Next

    End Sub

    'Metoda po zmáčknutí tlačítka pro porovnání
    Private Sub BtnCompare(sender As Object, e As EventArgs) Handles BtnComp.Click

        Trace.Listeners.Add(New TextWriterTraceListener("Debug.log", "myListener"))
        Trace.WriteLine("Starting comparing @ " + DateTime.Now.ToString())
        Trace.Indent()

        Try

            'Zákaz přepočítávání
            Trace.WriteLine("Disabling auto-update")
            autoUpdate(False)

            'Získání parametrů (názvy, řádky, sloupce) vybraných listů
            Trace.WriteLine("Getting sheets parameters")
            GetSheetParams()

            'Zobrazení formuláře s progress barem
            Trace.WriteLine("Showing progress bar")
            ShowForm()

            'Vytvoření souboru pro zápis výsledků
            PrLbl.Text = "Creating output"
            CreateResult()

            'Spuštění porovnávadla
            PrLbl.Text = "Starting comparison"
            Trace.WriteLine("Starting Comparison")
            Compare()

            'Uložení
            PrLbl.Text = "Saving"
            Dim path As String = NewWb.Path

            ResultWb.SaveAs(path & "\compared", FileFormat:=51)

            'Zavření progress baru
            FormProgBar.Hide()

            'Povolení přepočítávání
            Trace.WriteLine("Enabling auto-update")
            autoUpdate(True)

            'Zavření sešitů
            ResultWb.Close(SaveChanges:=True)

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

    'Vrátí pole indexů, podle kterých se vyhledává
    Function GetIndArr(array As Object, col As Integer, len As Integer)

        Dim IndArr(len) As String

        For i = 1 To len

            IndArr(i) = array(i, col)

        Next

        Return IndArr

    End Function

    'Poskytuje hodnotu pro velikost pole (který sešit má víc řádků/sloupců)
    Function GetBiggerDim(x As Integer, y As Integer) As Integer

        If x > y Then
            GetBiggerDim = x
        Else
            GetBiggerDim = y
        End If

        Return GetBiggerDim

    End Function

    'Zobrazí formulář progress baru
    Sub ShowForm()

        'Zobrazí formulář
        FormProgBar.Show()
        'Nastaví aktuální hodnotu baru na 1 - začátek
        PrBar.Value = 1
        'Nastaví maximum na počet řádků v "novém" listu
        PrBar.Maximum = NewRows
        'Přepíše label
        PrLbl.Text = "Starting"

    End Sub

    'Vrací číslo sloupce, podle kterého se vyhledává
    Sub ColSelect(TboxVal As String)

        'Přepis písmene sloupce na číslo
        Dim IntCatch As Integer

        Trace.WriteLine("Is column numeric")
        'Je sloupec zadán jako číslo?
        If IsNumeric(TboxVal) Then

            'Je číslo integer?
            If Integer.TryParse(TboxVal, IntCatch) Then

                ColLookup = TboxVal
                'Trace.WriteLine("Is numeric")

            Else

                MsgBox("Invalid input - Search by column must be integer")
                Trace.WriteLine("Is numeric but not integer")

            End If

        Else

            Try

                'Hodnota není číslo - písmeno se převede na číslo sloupce
                ColLookup = NewSheet.Range(TboxVal & "1").Column
                Trace.WriteLine("Is not numeric and can be turned to column")

            Catch ex As Exception

                MsgBox("Error: " & ex.Message)
                Trace.WriteLine("Is not numeric and cannot be turned to column")

            End Try

        End If

    End Sub

    'Vytváří výstupní soubor
    Sub CreateResult()
        Dim path As String = NewWb.Path

        'Vytvoří výstupní soubor a uloží ho jako formát .xlsx (51), aby se předešlo problémům s kompatibilitou        
        ResultWb = XlApp.Workbooks.Add
        XlApp.ActiveSheet.Name = "NewWbSheet"
        CopyOld(ResultWb, OldWb)

        'Vymazání listu, který se tvoří automaticky s novým sešitem
        XlApp.DisplayAlerts = False
        ResultWb.Sheets("NewWbSheet").Delete
        XlApp.DisplayAlerts = True

        'Zkopírování listů a přiřazení do proměnných
        OldResSheet = ResultWb.Worksheets(OldSheet.Name)
        OldResSheet.Name = "Cancelled"

        NewSheet.Copy(Before:=OldResSheet)
        NewResSheet = XlApp.ActiveSheet
        NewResSheet.Name = OldSheet.Name

    End Sub

    'Metoda při načtení hlavního formuláře
    Private Sub Skompare_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Vytvoření aplikace excel, se ktrou se dále bude pracovat
        XlApp = New Excel.Application

    End Sub

    'Vrátí poslední buňku ve sloupci
    Private Function GetLast(ws As Excel.Worksheet, order As Excel.XlSearchOrder) As Excel.Range
        GetLast = ws.Cells.Find(What:="*",
                                  After:=ws.Cells(1, 1),
                                  LookIn:=Excel.XlFindLookIn.xlFormulas,
                                  LookAt:=Excel.XlLookAt.xlPart,
                                  SearchOrder:=order,
                                  SearchDirection:=Excel.XlSearchDirection.xlPrevious,
                                  MatchCase:=False)
    End Function

    'Metoda při ukončení hlavního formuláře
    Private Sub Skompare_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If MessageBox.Show("Close?", "Close", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

            'Některé můžou být zavřené
            Try
                OldWb.Close(SaveChanges:=False)
                NewWb.Close(SaveChanges:=False)
                ResultWb.Close(SaveChanges:=False)
            Catch ex As Exception
            End Try

            XlApp.Quit()

        Else
            e.Cancel = True
        End If
    End Sub

    'Metoda pro kopírování listů do výstupního souboru
    Sub CopyOld(res As Excel.Workbook, old As Excel.Workbook)

        Dim oldSheets As Excel.Sheets = old.Worksheets()
        Dim x As Integer = 1

        For Each sheet As Excel.Worksheet In oldSheets

            sheet.Copy(After:=res.Worksheets(x))
            x += 1

        Next

    End Sub

    'Metoda pro nastavení automatického přepočítávání/updatů sešitu
    Sub autoUpdate(auto As Boolean)

        If auto = True Then

            'Povolení přepočítávání, updateování apod. sešitu během výpočtu
            With XlApp
                .Calculation = Excel.XlCalculation.xlCalculationAutomatic
                .ScreenUpdating = True
                .DisplayStatusBar = True
                .EnableEvents = True
            End With

        ElseIf auto = False Then

            'Zákaz přepočítávání, updateování apod. sešitu během výpočtu
            With XlApp
                .Calculation = Excel.XlCalculation.xlCalculationManual
                .ScreenUpdating = False
                .DisplayStatusBar = False
                .EnableEvents = False
            End With

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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    End Sub

    'Převádí číslo sloupce na písmeno
    Private Function GetExcelColumnName(columnNumber As Integer) As String

        Dim columnName As String = String.Empty
        Dim modulo As Integer

        While columnNumber > 0
            modulo = (columnNumber - 1) Mod 26
            columnName = Convert.ToChar(65 + modulo).ToString() & columnName
            columnNumber = CInt((columnNumber - modulo) / 26)
        End While

        Return columnName
    End Function

    'Najde velikost záhlaví a pozici UID kódu podle pojmenovaného rozsahu v sešitu
    Private Sub BtnGetStartPoint_Click(sender As Object, e As EventArgs) Handles BtnGetStartPoint.Click
        GetSheetParams()
        TBoxColSelect.Text = GetExcelColumnName(NewSheet.Range("UID").Column)
        TBoxStart.Text = NewSheet.Range("Header").Rows.Count
    End Sub
End Class

'Třída pro volání konzole
Public Class Win32

    <DllImport("kernel32.dll")> Public Shared Function AllocTrace() As Boolean

    End Function
    <DllImport("kernel32.dll")> Public Shared Function FreeTrace() As Boolean

    End Function

End Class

