Imports Excel = Microsoft.Office.Interop.Excel

Public Class SkompareMain

    'Deklarace aplikace excel
    Public XlApp As Excel.Application

    'Deklarace sešitů
    Public NewWb As Excel.Workbook
    Public OldWb As Excel.Workbook
    Public ResultWb As Excel.Workbook

    'Deklarace listů
    Public NewSheet As Excel.Worksheet
    Public OldSheet As Excel.Worksheet
    Public NewResSheet As Excel.Worksheet
    Public OldResSheet As Excel.Worksheet

    'Deklarace parametrů vybraných listů
    'Počty řádků
    Public NewRows As Integer
    Public OldRows As Integer
    'Počty sloupců
    Public NewCols As Integer
    Public OldCols As Integer
    'Větší počet řádků
    Public lenRows As Integer
    'Větší počet řádků
    Public lenCols As Integer
    'Sloupec pro vyhledávání
    Public ColLookup As Int16

    'Deklarace polí pro porovnání řádků
    Public NewRowArr As Object(,)
    Public OldRowArr As Object(,)

    'Deklarace proměnných pro ovládání progress baru a jeho popisku
    Public PrBar As Object = FormProgBar.ProgBar
    Public PrLbl As Object = FormProgBar.LblProgBar

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

    'Metoda pro nalezení duplicitních jedinečných kódů
    Public Function CheckDuplicities() As Boolean
        Return CheckDuplicities = False
    End Function

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

    'Prochází sešity a porovnává řádky (vyznačení změn v řádku řeší samostatná funkce)
    Public Sub Compare()

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
        ColSelect(FormSkompare.TBoxColSelect.Text)

        'Získání startovacího řádku
        PrLbl.Text = "Checking start row input"
        Trace.WriteLine("Checking start row input")
        Dim StartRow As Int16
        'Kontrola na integer
        If Integer.TryParse(FormSkompare.TBoxStart.Text, StartRow) = False Then
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

                NewResSheet.Rows(NewRow).EntireRow.Interior.Color = FormSkompare.TBoxColor.BackColor

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

    'Rozhoduje, jak se vyznačí změna
    Private Sub CompareStyle(NewRng As Excel.Range, NewStr As String, OldStr As String)

        'Jen obarvení
        If FormSkompare.RBtnStyle1.Checked Then
            NewRng.Interior.Color = FormSkompare.TBoxColor.BackColor
            NewRng.Value = NewStr

            'Obarvení a komentář
        ElseIf FormSkompare.RBtnStyle2.Checked Then
            NewRng.Interior.Color = FormSkompare.TBoxColor.BackColor
            NewRng.Value = NewStr
            'Vyhazuje výjimku, pokud je komentář prázdný
            If OldStr = "" Then
                NewRng.AddComment("-")
            Else
                NewRng.AddComment(OldStr)
            End If

            'Obarvení a řetězec
        ElseIf FormSkompare.RBtnStyle3.Checked Then
            NewRng.Interior.Color = FormSkompare.TBoxColor.BackColor
            NewRng.Value = NewStr & " " _
                & FormSkompare.TBoxStringStart.Text _
                & OldStr _
                & FormSkompare.TBoxStringEnd.Text _

            'Jen komentář
        ElseIf FormSkompare.RBtnStyle4.Checked Then
            NewRng.Value = NewStr
            'Vyhazuje výjimku, pokud je komentář prázdný
            If OldStr = "" Then
                NewRng.AddComment("-")
            Else
                NewRng.AddComment(OldStr)
            End If

            'Jen řetězec
        ElseIf FormSkompare.RBtnStyle5.Checked Then
            NewRng.Value = NewStr & " " _
                & FormSkompare.TBoxStringStart.Text _
                & OldStr _
                & FormSkompare.TBoxStringEnd.Text _

            'Řetězec v komentáři
        ElseIf FormSkompare.RBtnStyle6.Checked Then
            NewRng.Value = NewStr
            'Vyhazuje výjimku, pokud je komentář prázdný
            If OldStr = "" Then
                NewRng.AddComment("-")
            Else
                NewRng.AddComment(FormSkompare.TBoxStringStart.Text _
                                  & OldStr _
                                  & FormSkompare.TBoxStringEnd.Text)
            End If

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

    'Metoda pro vymazání nalezených (označeno zeleně) řádek ve "zrušeném" listu
    Sub DeleteRows(sheet As Excel.Worksheet, indexArray() As Integer)

        For i As Integer = indexArray.Length - 1 To FormSkompare.TBoxStart.Text Step -1

            If indexArray(i) = 1 Then

                sheet.Rows(i).EntireRow.Delete

            End If

        Next

    End Sub

    'Poskytuje hodnotu pro velikost pole (který sešit má víc řádků/sloupců)
    Function GetBiggerDim(x As Integer, y As Integer) As Integer

        If x > y Then
            GetBiggerDim = x
        Else
            GetBiggerDim = y
        End If

        Return GetBiggerDim

    End Function

    'Převádí číslo sloupce na písmeno
    Public Function GetExcelColumnName(columnNumber As Integer) As String

        Dim columnName As String = String.Empty
        Dim modulo As Integer

        While columnNumber > 0
            modulo = (columnNumber - 1) Mod 26
            columnName = Convert.ToChar(65 + modulo).ToString() & columnName
            columnNumber = CInt((columnNumber - modulo) / 26)
        End While

        Return columnName
    End Function

    'Vrátí pole indexů, podle kterých se vyhledává
    Function GetIndArr(array As Object, col As Integer, len As Integer)

        Dim IndArr(len) As String

        For i = 1 To len

            IndArr(i) = array(i, col)

        Next

        Return IndArr

    End Function

    'Vrátí poslední buňku ve sloupci/řádku
    Private Function GetLast(ws As Excel.Worksheet, order As Excel.XlSearchOrder) As Excel.Range
        GetLast = ws.Cells.Find(What:="*",
                                  After:=ws.Cells(1, 1),
                                  LookIn:=Excel.XlFindLookIn.xlFormulas,
                                  LookAt:=Excel.XlLookAt.xlPart,
                                  SearchOrder:=order,
                                  SearchDirection:=Excel.XlSearchDirection.xlPrevious,
                                  MatchCase:=False)
    End Function

    'Získává data o počtech řádků a sloupců v jednotlivých sešitech
    Public Function GetSheetParams(newSheetName As String, oldSheetName As String)

        'Přiřazením globálním proměnným
        OldSheet = OldWb.Worksheets(oldSheetName)
        NewSheet = NewWb.Worksheets(newSheetName)

        'Deklarace pole
        Dim returnArr()() As String = New String(1)() {}
        returnArr(0) = New String(1) {}
        returnArr(1) = New String(1) {}

        returnArr(0)(0) = CStr(GetLast(OldSheet, order:=Excel.XlSearchOrder.xlByColumns).Row)
        returnArr(0)(1) = CStr(GetLast(NewSheet, order:=Excel.XlSearchOrder.xlByColumns).Row)

        returnArr(1)(0) = CStr(GetLast(OldSheet, order:=Excel.XlSearchOrder.xlByRows).Column)
        returnArr(1)(1) = CStr(GetLast(NewSheet, order:=Excel.XlSearchOrder.xlByRows).Column)

        OldRows = returnArr(0)(0)
        NewRows = returnArr(0)(1)

        OldCols = returnArr(1)(0)
        NewCols = returnArr(1)(1)

        Return returnArr

    End Function

    'Otevírá sešity pro porovnání
    Public Sub OpenExcel(FilePath As String, sender As Object)

        'Výběr, které objekty se upraví podle stisknutého tlačítka
        If sender Is FormSkompare.BtnNew Then 'Stisknuto "nové" tlačítko

            'Otevření souboru v aplikaci Excel
            NewWb = XlApp.Workbooks.Open(FilePath, [ReadOnly]:=True)

        ElseIf sender Is FormSkompare.BtnOld Then 'Stisknuto "staré" tlačítko

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

End Class
