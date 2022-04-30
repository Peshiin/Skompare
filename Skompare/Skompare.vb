Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Globalization
Imports System.Threading
Imports System.Diagnostics
Imports System.Runtime.InteropServices

Public Class SkompareMain

    '###############################################################
    '           Properties
    '###############################################################

    'Deklarace aplikace excel
    Private XlApp As Excel.Application
    Property Application As Excel.Application
        Get
            Return XlApp
        End Get
        Set(value As Excel.Application)
            XlApp = value
        End Set
    End Property

    'Deklarace sešitů
    Private NewWb As Excel.Workbook
    Private OldWb As Excel.Workbook
    Private ResultWb As Excel.Workbook
    Private Wbs = New Excel.Workbook() {NewWb, OldWb, ResultWb}
    Property Workbooks As Excel.Workbook()
        Get
            Return Wbs
        End Get
        Set(value() As Excel.Workbook)
            Wbs = value
            NewWb = value(1)
            OldWb = value(2)
            NewWb = value(3)
        End Set
    End Property

    'Deklarace listů
    Private NewSheet As Excel.Worksheet
    Private OldSheet As Excel.Worksheet
    Private NewResSheet As Excel.Worksheet
    Private OldResSheet As Excel.Worksheet

    'Deklarace parametrů vybraných listů
    'Počty řádků
    Private NewRows As Integer
    Private OldRows As Integer
    'Počty sloupců
    Private NewCols As Integer
    Private OldCols As Integer
    'Větší počet řádků
    Private lenRows As Integer
    'Větší počet řádků
    Private lenCols As Integer

    'Sloupce pro vyhledávání
    Private SearchKeysCols(2) As Integer

    'Declaration of the row where the comparing shall start (to ignore header)
    Private startRow As Integer

    'Deklarace polí porovnávaných rozsahů
    Private NewArr As Object(,)
    Private OldArr As Object(,)

    'Declaration of search arrays
    Private NewSearchArr As String()
    Private OldSearchArr As String()

    'Deklarace polí pro porovnání řádků
    Private NewRowArr As Object(,)
    Private OldRowArr As Object(,)

    'Deklarace proměnných pro ovládání progress baru a jeho popisku
    Private PrBar As Object = FormProgBar.ProgBar
    Private PrLbl As Object = FormProgBar.LblProgBar

    '###############################################################
    '           Methods
    '###############################################################


    '           Opening File
    '###############################################################
    'Open workbook respective to the old/new button
    Public Sub OpenWorkbook(sender As Object)

        'Gets path of the opening file via file dialog
        Dim FilePath = GetFilePathFD(FormSkompare.OpenFD)

        If FilePath Is Nothing Then
            MessageBox.Show("Nebyl vybrán soubor")
            Exit Sub
        End If

        'Assigns opened workbook to the class variable
        'Writes the name of the respective file to the UI
        'Writes the name of sheets to the UI
        If sender Is FormSkompare.BtnNew Then

            NewWb = XlApp.Workbooks.Open(FilePath, [ReadOnly]:=True)
            FormSkompare.LblNewFileName.Text = Dir(FilePath)
            WriteWorksheetsToUI(NewWb, FormSkompare.CBoxNewSheets)

        ElseIf sender Is FormSkompare.BtnOld Then

            OldWb = XlApp.Workbooks.Open(FilePath, [ReadOnly]:=True)
            FormSkompare.LblOldFileName.Text = Dir(FilePath)
            WriteWorksheetsToUI(NewWb, FormSkompare.CBoxOldSheets)

        End If


    End Sub

    'Opens file dialog and returns selected file path
    Private Function GetFilePathFD(fd As FileDialog) As String

        'Otevře dialogové okno pro výběr souboru
        fd.Title = "Select file"
        fd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"

        If fd.ShowDialog() = DialogResult.OK Then

            'Získá cestu vybraného souboru jako String
            GetFilePathFD = fd.FileName
            Return GetFilePathFD

        Else
            Return Nothing
        End If

    End Function

    'Lists names of worksheets to the UI
    Private Sub WriteWorksheetsToUI(wb As Excel.Workbook, cBox As ComboBox)

        'Clears the cBox
        cBox.Items.Clear()

        'Writes names of all worksheets in respective workbook to the cBox
        For Each ws As Excel.Worksheet In wb.Worksheets

            cBox.Items.Add(ws.Name)

            'Set selected item to something so the cBox doesn't appear empty
            If cBox.SelectedItem Is Nothing Then
                cBox.SelectedItem = ws.Name
            End If

        Next

    End Sub


    '           Data extraction
    '###############################################################
    'Compare function initiation
    Public Sub CompareInit()

        'Tries to assign sheet parameters if workbooks are assigned
        If NewWb IsNot Nothing And
            OldWb IsNot Nothing Then

            AssignSheetsParams(FormSkompare.CBoxNewSheets.SelectedItem,
                               FormSkompare.CBoxOldSheets.SelectedItem)
        Else
            MessageBox.Show("Nejsou vybrány soubory pro porovnání")
            Exit Sub
        End If

        'Checks the inputs
        If CheckInput() = False Then
            Exit Sub
        End If

        'Assigns sheets variables
        AssignSheetsParams(FormSkompare.CBoxNewSheets.SelectedItem,
                           FormSkompare.CBoxOldSheets.SelectedItem)

        'Assigns starting row
        startRow = FormSkompare.TBoxStart.Text

        'Assigns columns to search by
        SearchKeysCols(0) = ColSelect(FormSkompare.TBoxColSelect1.Text)
        If FormSkompare.TBoxColSelect2.Enabled Then
            SearchKeysCols(1) = ColSelect(FormSkompare.TBoxColSelect2.Text)
        End If
        If FormSkompare.TBoxColSelect3.Enabled Then
            SearchKeysCols(2) = ColSelect(FormSkompare.TBoxColSelect3.Text)
        End If

        'Assigns sheets arrays
        GetSheetArrays()

        'Assigns search arrays
        NewSearchArr = GetSearchArray(NewArr)
        OldSearchArr = GetSearchArray(OldArr)



    End Sub

    'Writes main parameters to the tBox
    Public Sub ShowMainParams(tBox As RichTextBox)

        'Tries to assign sheet parameters if workbooks are assigned
        If NewWb IsNot Nothing And
            OldWb IsNot Nothing Then

            AssignSheetsParams(FormSkompare.CBoxNewSheets.SelectedItem,
                               FormSkompare.CBoxOldSheets.SelectedItem)
        Else
            MessageBox.Show("Nejsou vybrány soubory pro porovnání")
            Exit Sub
        End If

        If CheckInput() Then

            'Clears tBox from previous data
            tBox.Clear()

            tBox.AppendText(vbTab _
                                + "Nový sešit" _
                                + vbTab _
                                + "Starý sešit")

            tBox.AppendText(Environment.NewLine _
                                + "Sheet name:" _
                                + vbTab _
                                + NewSheet.Name _
                                + vbTab _
                                + OldSheet.Name)

            tBox.AppendText(Environment.NewLine _
                                + "Row count:" _
                                + vbTab _
                                + CStr(NewRows) _
                                + vbTab _
                                + CStr(OldRows))

            tBox.AppendText(Environment.NewLine _
                                    + "Column count:" _
                                    + vbTab _
                                    + CStr(NewCols) _
                                    + vbTab _
                                    + CStr(OldCols))

        End If

    End Sub

    'Assigning "new" and "old" sheets to variables and setting their lenghts
    Private Sub AssignSheetsParams(newSheetName As String, oldSheetName As String)

        If NewWb IsNot Nothing And
            OldWb IsNot Nothing Then

            'Assigning sheets to variables
            NewSheet = NewWb.Sheets(newSheetName)
            OldSheet = OldWb.Sheets(oldSheetName)

            'Getting number of rows and columns in "new" sheet
            NewRows = GetLast(NewSheet, Excel.XlSearchOrder.xlByColumns).Row
            NewCols = GetLast(NewSheet, Excel.XlSearchOrder.xlByRows).Column

            'Getting number of rows and columns in "old" sheet
            OldRows = GetLast(OldSheet, Excel.XlSearchOrder.xlByColumns).Row
            OldCols = GetLast(OldSheet, Excel.XlSearchOrder.xlByRows).Column

            'Getting the bigger number of rows
            lenRows = GetBiggerDim(NewRows, OldRows)
            lenCols = GetBiggerDim(NewCols, OldCols)

        End If

    End Sub

    'Checks if all the key data are filled
    Private Function CheckInput() As Boolean

        'Is Excel application assigned?
        If XlApp Is Nothing Then
            MessageBox.Show("Není přiřazena aplikace Excel")
            Return False

            'Is "new" workbook assigned?
        ElseIf NewWb Is Nothing Then
            MessageBox.Show("Nebyl přiřazen ""nový"" sešit Excel")
            Return False
            'Is "old" workbook assigned?
        ElseIf OldWb Is Nothing Then
            MessageBox.Show("Nebyl přiřazen ""starý"" sešit Excel")
            Return False

            'Is "new" worksheet assigned?
        ElseIf NewSheet Is Nothing Then
            MessageBox.Show("Nebyl přiřazen ""nový"" list Excel")
            Return False
            'Is "old" worksheet assigned?
        ElseIf OldSheet Is Nothing Then
            MessageBox.Show("Nebyl přiřazen ""starý"" list Excel")
            Return False

            'Do both sheets have the same number of columns?
        ElseIf NewCols <> OldCols Then
            MessageBox.Show("Počty sloupců v porovnávaných listech se liší")
            Return False

        Else
            Return True
        End If

    End Function

    'Gets array of indexes to search by
    Public Function GetSearchArray(inputArr As Object(,))

        Dim len As Integer

        'Defines length of the return array
        len = UBound(inputArr, 1)

        Dim returnArr(len) As String

        For row As Integer = startRow To len

            For Each key In SearchKeysCols

                If key > 0 Then

                    returnArr(row) &= inputArr(row, key)

                End If

            Next

        Next

            Return returnArr

    End Function

    'Gets array of values from range to be compared
    Public Sub GetSheetArrays()

        Dim lastCell As String

        'Assigns "new" array of compared values
        lastCell = CStr(GetExcelColumnName(lenCols)) & CStr(NewRows)
        NewArr = CType(NewSheet.Range("A1", lastCell).Value, Object(,))

        'Assigns "old" array of compared values
        lastCell = CStr(GetExcelColumnName(lenCols)) & CStr(OldRows)
        OldArr = CType(NewSheet.Range("A1", lastCell).Value, Object(,))

    End Sub











    'Allow/disable auto updates of Excel App
    Sub autoUpdate(auto As Boolean)

        If auto = True Then

            'Allow auto updates
            With XlApp
                .Calculation = Excel.XlCalculation.xlCalculationAutomatic
                .ScreenUpdating = True
                .DisplayStatusBar = True
                .EnableEvents = True
            End With

        ElseIf auto = False Then

            'Disable auto updates
            With XlApp
                .Calculation = Excel.XlCalculation.xlCalculationManual
                .ScreenUpdating = False
                .DisplayStatusBar = False
                .EnableEvents = False
            End With

        End If

    End Sub

    'Vrací číslo sloupce, podle kterého se vyhledává
    Private Function ColSelect(TboxVal As String) As Integer

        'Přepis písmene sloupce na číslo
        Dim IntCatch As Integer

        Trace.WriteLine("Is column numeric")
        'Je sloupec zadán jako číslo?
        If IsNumeric(TboxVal) Then

            'Je číslo integer?
            If Integer.TryParse(TboxVal, IntCatch) Then

                Return TboxVal

            Else

                MsgBox("Invalid input - Search by column must be integer")
                Trace.WriteLine("Is numeric but not integer")
                Return Nothing

            End If

        Else

            Try

                'Hodnota není číslo - písmeno se převede na číslo sloupce
                Return NewSheet.Range(TboxVal & "1").Column
                Trace.WriteLine("Is not numeric and can be turned to column")

            Catch ex As Exception

                MsgBox("Error: " & ex.Message)
                Trace.WriteLine("Is not numeric and cannot be turned to column")
                Return Nothing

            End Try

        End If

    End Function

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
        Dim NewArr As Object(,) = CType(NewSheet.UsedRange.Value, Object(,))
        Dim OldArr As Object(,) = CType(OldSheet.UsedRange.Value, Object(,))

        'Získá číslo sloupce, podle kterého se bude hledat
        Trace.WriteLine("Getting key column")
        PrLbl.Text = "Getting key column"
        'ColLookupPrim = ColSelect(FormSkompare.TBoxColSelect1.Text)
        'ColLookupSec = ColSelect(FormSkompare.TBoxColSelect2.Text)

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
        'Seznam položek s duplicitním UID
        Dim duplicityArr As String()

        'Získání pole vyhledávaných indexů starého pole
        Dim OldIndArr() As String
        'OldIndArr = GetIndArr(OldArr, ColLookupPrim, OldRows)
        'Deklarace pole se sekundárním vyhledávacím indexem
        Dim OldIndArrSecondary() As String
        'OldIndArrSecondary = GetIndArr(OldArr, ColLookupSec, OldRows)

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
            'SearchString = NewArr(NewRow, ColLookupPrim)

            'Vrátí polohu (řádek) hledaného kódu ve "starém" poli
            OldRow = Array.IndexOf(OldIndArr, SearchString)

            'Nalezena shoda identifikátoru?
            If OldRow > 0 Then

                'Kontroluje duplicitu
                If Duplicity(OldRow) = 1 Then

                    'Získá pole čísel řádků se stejným SearchString
                    duplicityArr = GetDuplicityList(OldIndArr, SearchString)
                    'Nastaví SearchString dle sekundárního klíče
                    'SearchString = NewArr(NewRow, ColLookupSec)

                    'Prochází pole duplicit  
                    For Each element As String In duplicityArr
                        Trace.WriteLine(OldIndArrSecondary(element) & " " & SearchString & " " & element)
                        If OldIndArrSecondary(element) = SearchString Then
                            OldRow = element
                            Duplicity(OldRow) = 1
                        End If
                    Next

                End If

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

    'Metoda pro nalezení duplicitních jedinečných kódů
    Public Function GetDuplicityList(arr As Array, str As String) As String()

        Dim row = 0
        Dim duplicityList As New List(Of String)

        For Each element In arr
            If CStr(element) = str Then
                duplicityList.Add(row)
            End If
            row += 1
        Next

        GetDuplicityList = duplicityList.ToArray

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


        ElseIf sender Is FormSkompare.BtnOld Then 'Stisknuto "staré" tlačítko

            'Otevření souboru v aplikaci Excel
            OldWb = XlApp.Workbooks.Open(FilePath, [ReadOnly]:=True)

        End If

    End Sub

    'Vypisuje listy sešitů do přehledového okénka
    Sub WriteFileData(Wb As Excel.Workbook, FileName As String, Cbox As Object, nameLbl As Object)

        'Vypsání názvu souboru do formuláře (Dir() vybere pouze název souboru a ne celou cestu)
        nameLbl.Text = Dir(FileName)

        'Vyčištění ListBoxu od popisku
        Cbox.Items.Clear()
        'Vypsání názvů listů ve vybraném sešitu
        For Each sheet In Wb.Worksheets
            Cbox.Items.Add(sheet.Name)
        Next

        'Nastaví do comboboxu hodnotu prvního listu
        Dim sheetOne As Excel.Worksheet = Wb.Worksheets(1)
        Cbox.SelectedIndex = Cbox.FindStringExact(sheetOne.Name)

    End Sub

End Class
