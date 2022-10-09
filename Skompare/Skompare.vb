Imports Excel = Microsoft.Office.Interop.Excel
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Globalization
Imports System.Threading
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports Microsoft.Vbe.Interop
Imports System.Reflection

Public Class SkompareMain

    '###############################################################
    '           Properties
    '###############################################################

    'Deklarace aplikace excel
    Public XlApp As Excel.Application

    'Deklarace cest k sešitům
    Private NewPath As String
    Private OldPath As String

    'Deklarace sešitů
    Private NewWb As Excel.Workbook
    Private OldWb As Excel.Workbook
    Private ResultWb As Excel.Workbook

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
    'Stores if there is different nubmer of columns in the two workbooks
    Private differentCols As Boolean = False

    'Sloupce pro vyhledávání
    Private SearchKeysCols(2) As Integer

    'Declaration of the row where the comparing shall start (to ignore header)
    Private startRow As Integer

    'Defines which style of comparing is used
    Private compStyle As String

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
    Private prBar As Object
    Private prLbl As Object

    'Declaration of color to highlight changes
    Private highlight As Color
    'Start and end strings for marking changes
    Private startStr As String
    Private endStr As String


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

            'Closes previously opened workbook
            Try
                If NewWb IsNot Nothing Then
                    NewWb.Close()
                End If
            Catch
            End Try

            NewPath = FilePath
            NewWb = XlApp.Workbooks.Open(FilePath, [ReadOnly]:=True)
            FormSkompare.LblNewFileName.Text = Dir(FilePath)
            WriteWorksheetsToUI(NewWb, FormSkompare.CBoxNewSheets)

        ElseIf sender Is FormSkompare.BtnOld Then

            Try
                If OldWb IsNot Nothing Then
                    'Closes previously opened workbook
                    OldWb.Close()
                End If
            Catch
            End Try

            OldPath = FilePath
            OldWb = XlApp.Workbooks.Open(FilePath, [ReadOnly]:=True)
            FormSkompare.LblOldFileName.Text = Dir(FilePath)
            WriteWorksheetsToUI(OldWb, FormSkompare.CBoxOldSheets)

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

    'Gets the number of starting row and does necessary checks
    Private Function GetStartRow(input As String) As Integer

        'Kontrola na integer
        If Integer.TryParse(input, GetStartRow) = False Then
            MessageBox.Show("Zadaná hodnota počátečního řádku musí být typu integer")
            Exit Function
        Else
            Return GetStartRow
        End If
    End Function

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
        Try

            If NewWb IsNot Nothing And
            OldWb IsNot Nothing Then

                'Assigning sheets to variables
                NewSheet = NewWb.Sheets(newSheetName)
                OldSheet = OldWb.Sheets(oldSheetName)

                'Getting number of rows and columns in "new" sheet
                NewRows = GetLast(NewSheet, Excel.XlSearchOrder.xlByColumns)
                NewCols = GetLast(NewSheet, Excel.XlSearchOrder.xlByRows)

                'Getting number of rows and columns in "old" sheet
                OldRows = GetLast(OldSheet, Excel.XlSearchOrder.xlByColumns)
                OldCols = GetLast(OldSheet, Excel.XlSearchOrder.xlByRows)

                'Getting the bigger number of rows
                lenRows = GetBiggerDim(NewRows, OldRows)
                lenCols = GetBiggerDim(NewCols, OldCols)

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try

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

        ElseIf startRow > NewRows Then
            MessageBox.Show("Zadaný počáteční řádek je vyšší než počet řádků v sešitu.")
            Return False

            'Is any style defining radio button checked?
        Else
            Dim styleChecked = False

            For Each control In FormSkompare.GBoxCompareStyle.Controls.OfType(Of RadioButton)
                If control.Checked Then
                    styleChecked = True
                End If
            Next

            If styleChecked = False Then
                MessageBox.Show("Zaškrtněte způsob označování změn.")
            End If

            Return styleChecked

        End If

        'Do both sheets have the same number of columns?
        If NewCols <> OldCols Then

            MessageBox.Show("Počty sloupců v porovnávaných listech se liší")
            differentCols = True
            Return True

        Else
            differentCols = False
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
        OldArr = CType(OldSheet.Range("A1", lastCell).Value, Object(,))

    End Sub

    'Gets number of the input column
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

    'Returns bigger value from the two input
    Function GetBiggerDim(x As Integer, y As Integer) As Integer

        If x > y Then
            GetBiggerDim = x
        Else
            GetBiggerDim = y
        End If

        Return GetBiggerDim

    End Function

    'Gets letter of a column from input number
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

    'Gets last cell in column/row
    Private Function GetLast(ws As Excel.Worksheet, order As Excel.XlSearchOrder) As Integer
        Dim last As Excel.Range

        last = ws.Cells.Find(What:="*",
                                  After:=ws.Range("A1"),'Cells(1, 1),
                                  LookIn:=Excel.XlFindLookIn.xlValues,'Excel.XlFindLookIn.xlFormulas,
                                  LookAt:=Excel.XlLookAt.xlPart,
                                  SearchOrder:=order,
                                  SearchDirection:=Excel.XlSearchDirection.xlPrevious,
                                  MatchCase:=False)

        'checks if "last" has any value
        If last Is Nothing Then
            Throw New Exception("Vybraný list (" & ws.Name & ") je pravděpodobně prázdný.")
        End If

        'Looking for last row
        If order = Excel.XlSearchOrder.xlByColumns Then
            If last.Row < ws.UsedRange.Rows.Count Then
                GetLast = ws.UsedRange.Rows.Count
                Exit Function
            Else
                GetLast = last.Row
                Exit Function
            End If
            'looking for last column
        ElseIf order = Excel.XlSearchOrder.xlByRows Then
            If last.Column < ws.UsedRange.Columns.Count Then
                GetLast = ws.UsedRange.Columns.Count
                Exit Function
            Else
                GetLast = last.Column
                Exit Function
            End If
        End If

        Return Nothing

    End Function




    '           Comparing
    '###############################################################

    'Compare function initialization
    Public Sub CompareInit()

        'Checks if Debug.log exists and deletes it
        Dim debugFilePath As String = My.Application.Info.DirectoryPath & "\\Debug.log"
        If System.IO.File.Exists(debugFilePath) Then
            Try
                My.Computer.FileSystem.DeleteFile(debugFilePath)
            Catch
            End Try
        End If

        'Initializes tracing for debugging
        Dim debug As New TextWriterTraceListener(My.Application.Info.DirectoryPath & "\\Debug.log", "myListener")
        Trace.Listeners.Add(debug)
        Trace.WriteLine("Starting comparing @ " + DateTime.Now.ToString())
        Trace.Indent()

        'Checks open workbooks and opens them from NewPath/OldPath
        NewWb = XlApp.Workbooks.Open(NewPath, [ReadOnly]:=True)
        OldWb = XlApp.Workbooks.Open(OldPath, [ReadOnly]:=True)

        'Initialize progress bar form
        Dim prBarForm = New FormProgBar
        prBarForm.Show()
        prBar = prBarForm.ProgBar
        prLbl = prBarForm.LblProgBar
        prBarForm.LblProgBar.Text = "Inicializace porovnání"

        Try

            'Tries to assign sheet parameters if workbooks are assigned
            If NewWb IsNot Nothing And
                    OldWb IsNot Nothing Then

                AssignSheetsParams(FormSkompare.CBoxNewSheets.SelectedItem,
                                   FormSkompare.CBoxOldSheets.SelectedItem)

                'Assigns starting row
                startRow = GetStartRow(FormSkompare.TBoxStart.Text)

                'Checks the number of columns and lets the user to change them
                CheckColumns()
            Else
                MessageBox.Show("Nejsou vybrány soubory pro porovnání")
                prBarForm.Dispose()
                Exit Sub
            End If

            'Checks the inputs
            If CheckInput() = False Then
                prBarForm.Dispose()
                Exit Sub
            End If

            'Assigns comparing style to the name of the checked radio button
            compStyle = FormSkompare.GBoxCompareStyle.Controls.OfType(Of RadioButton).
                            Where(Function(r) r.Checked = True).
                            FirstOrDefault().Name

            'Assigns highlighting color and strings
            highlight = FormSkompare.TBoxColor.BackColor
            startStr = FormSkompare.TBoxStringStart.Text
            endStr = FormSkompare.TBoxStringEnd.Text

            'Assigns columns to search by
            'Goes through all the control elements with "ColSelect" tag in FormSkompare
            '!!!!!  does not necessarily find ColSelect1 as first   !!!!!
            Dim i As Integer = 0
            For Each control In FormSkompare.GBoxStatsDiff.Controls
                If control.Tag = "ColSelect" Then
                    If control.Enabled Then
                        SearchKeysCols(i) = ColSelect(control.Text)
                        i += 1
                    End If
                End If
            Next

            'Assigns sheets arrays
            GetSheetArrays()

            'Assigns search arrays
            NewSearchArr = GetSearchArray(NewArr)
            OldSearchArr = GetSearchArray(OldArr)

            'Sets initial value and boundaries of the progress bar
            ProgressBarInit(prBarForm)

            'Sets auto updating of the XlApp to false
            autoUpdate(False)

            'Creates "result" workbook to where the actual comparing will be done
            CreateResult()

            'Removes background color from "Cancelled"
            'RemoveBackground(OldResSheet)

            'Comparison itself
            Compare()

            'Hides progress bar form
            prBarForm.Dispose()

            'Allows auto updating
            autoUpdate(True)

            'Closes the originals and shows the result
            XlApp.Visible = True
            NewWb.Close(SaveChanges:=False)
            OldWb.Close(SaveChanges:=False)

            FormSkompare.Activate()


        Catch ex As Exception

            Trace.WriteLine(ex.StackTrace _
                            & Environment.NewLine _
                            & ex.Message)
            Trace.Flush()

            prBarForm.Dispose()

        End Try

        Trace.Unindent()
        Trace.WriteLine("Comparing ended")
        Trace.WriteLine("___________________________________________________")
        Trace.Flush()

    End Sub

    'Goes through the worksheets and compares rows
    Private Sub Compare()

        'Deklarace pomocných proměnných
        'Trackování, zda byla nalezena shoda
        Dim MatchFound As Boolean
        'Hledaná hodnota (jedinečný kód)
        Dim SearchString As String
        'Index ve "starém" poli, kde je hledaná hodnota
        Dim OldRow As Integer
        'Seznam položek s duplicitním UID
        Dim duplicityFound As Boolean = False

        'Získání pole pro kontrolu duplicit (0 = index zatím nenalezen)
        Dim Duplicity(lenRows) As Integer
        For Each element In Duplicity
            Duplicity(element) = 0
        Next

        'Prohledávací cyklus
        prLbl.Text = "Starting looping"
        Trace.WriteLine("Starting looping")
        'Loop v "nových" datech
        For NewRow = startRow To NewRows

            'Shoda nenalezena
            MatchFound = False

            'Hledaný jedinečný kód
            SearchString = NewSearchArr(NewRow)

            'Vrátí polohu (řádek) hledaného kódu ve "starém" poli
            OldRow = Array.IndexOf(OldSearchArr, SearchString)

            'Nalezena shoda identifikátoru?
            If OldRow > 0 Then

                'Kontroluje duplicitu
                If Duplicity(OldRow) = 1 Then

                    If duplicityFound = False Then

                        MessageBox.Show("Nalezena duplicita zadaných vyhledávacích klíčů" &
                                        Environment.NewLine &
                                        "Skript proběhne s předpokladem max. dvou duplicit." &
                                        Environment.NewLine &
                                        "Pokud je předpokládané množství duplicit více, ošetřete vhodným výběrem klíčů.")

                        duplicityFound = True

                    End If

                    OldRow = Array.IndexOf(OldSearchArr, SearchString, OldRow + 1)

                End If

                If OldRow > 0 Then

                    'Zaznamená nalezení shody
                    MatchFound = True
                    Duplicity(OldRow) = 1

                    'Porovná buňky v řádku
                    CompareRow(NewRow, OldRow)

                End If

            End If

            If MatchFound = False Then

                NewResSheet.Rows(NewRow).EntireRow.Interior.Color = highlight

            End If

            prBar.Value += 1
            prLbl.Text = "Progress: " _
                        & Math.Round((prBar.Value - startRow) / (NewRows - startRow), 2) * 100 _
                        & "% (" & NewRow & " out of " & NewRows & ")"

        Next

        'Smaže nalezené (zeleně označené) řádky ve "zrušeném" listu
        prLbl.Text = "Cleaning found rows from Cancelled"
        Trace.WriteLine("Cleaning found rows from Cancelled")
        DeleteRows(OldResSheet, Duplicity)

        'Nastavení zobrazení po dalším otevření sešitu (nebude najeto někam doprostřed listu a nastaví se scroll bar)
        Try
            OldResSheet.Activate()
            OldResSheet.Range("A1").Select()
        Catch ex As Exception
        End Try

    End Sub

    'Compares values in single rows
    Sub CompareRow(NewR As Integer, OldR As Integer)

        'Deklarace pomocných proměnných
        Dim NewVal As String
        Dim OldVal As String

        Try

            With NewResSheet.Rows(NewR)

                For col As Integer = 1 To lenCols

                    NewVal = NewArr.GetValue(NewR, col)
                    OldVal = OldArr.GetValue(OldR, col)

                    If String.Compare(NewVal, OldVal, True, CultureInfo.InvariantCulture) Then

                        CompareStyle(.Cells(1, col), NewVal, OldVal)

                    End If
                Next

            End With

        Catch
        End Try

    End Sub

    'Defines how the differences found shall be highlighted
    Private Sub CompareStyle(NewRng As Excel.Range, NewStr As String, OldStr As String)

        'Sets range format to Text
        NewRng.NumberFormat = "@"

        'Jen obarvení
        If compStyle = "RBtnStyle1" Then
            NewRng.Interior.Color = highlight
            NewRng.Value = NewStr

            'Obarvení a komentář
        ElseIf compStyle = "RBtnStyle2" Then

            NewRng.Interior.Color = highlight
            NewRng.Value = NewStr

            'Deletes existing comment if it exists
            If NewRng.Comment IsNot Nothing Then
                NewRng.Comment.Delete()
            End If

            'If the old value was nothing, dash will be written in the comment
            If OldStr = "" Then
                NewRng.AddComment("-")
            Else
                NewRng.AddComment(OldStr)
            End If

            'Obarvení a řetězec
        ElseIf compStyle = "RBtnStyle3" Then
            NewRng.Interior.Color = highlight
            NewRng.Value = NewStr & " " _
                & startStr _
                & OldStr _
                & endStr _

            'Jen komentář
        ElseIf compStyle = "RBtnStyle4" Then

            NewRng.Value = NewStr

            'Deletes existing comment if it exists
            If NewRng.Comment IsNot Nothing Then
                NewRng.Comment.Delete()
            End If

            'If the old value was nothing, dash will be written in the comment
            If OldStr = "" Then
                NewRng.AddComment("-")
            Else
                NewRng.AddComment(OldStr)
            End If

            'Jen řetězec
        ElseIf compStyle = "RBtnStyle5" Then
            NewRng.Value = NewStr & " " _
                & startStr _
                & OldStr _
                & endStr _

            'Řetězec v komentáři
        ElseIf compStyle = "RBtnStyle6" Then

            NewRng.Value = NewStr

            'Deletes existing comment if it exists
            If NewRng.Comment IsNot Nothing Then
                NewRng.Comment.Delete()
            End If

            'If the old value was nothing, dash will be written in the comment
            If OldStr = "" Then
                NewRng.AddComment("-")
            Else
                NewRng.AddComment(startStr _
                                  & OldStr _
                                  & endStr)
            End If

        End If
    End Sub





    '           Others
    '###############################################################
    'Initializes progress bar form based on inputs
    Private Sub ProgressBarInit(prBar As FormProgBar)

        prBar.ProgBar.Minimum = 1
        prBar.ProgBar.Maximum = lenRows + 1
        prBar.ProgBar.Value = startRow

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

    'Copies modules and macros from oldWb
    Sub CopyMacros(res As Excel.Workbook, old As Excel.Workbook)

        Dim dest As VBComponent

        Dim found As Boolean
        found = False

        Try
            'Iterate existing workbook And copy over the code modules
            For Each source As VBComponent In old.VBProject.VBComponents

                'Do we have any code lines in the code module to copy?
                If source.CodeModule.CountOfLines > 0 Then

                    'We need to check whether we already have a code module with that name in our workbook
                    'This will be for the sheets And workbook And we assume that we have already copied the sheets accordingly
                    For Each destNew As VBComponent In res.VBProject.VBComponents

                        If destNew.Name = source.CodeModule.Name Then

                            destNew.CodeModule.InsertLines(1, source.CodeModule.Lines(1, source.CodeModule.CountOfLines))
                            found = True

                            'We've found the matching codemodule so lets exit
                            Exit For

                        Else 'we have To create the code Module

                            found = False 'Set found To False so we can add the codemodule. 

                        End If

                    Next

                    If (found = False) Then

                        dest = res.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule)
                        dest.CodeModule.AddFromString(source.CodeModule.Lines(1, source.CodeModule.CountOfLines))

                        dest.Name = source.Name
                        Marshal.ReleaseComObject(dest)
                        Marshal.ReleaseComObject(source)
                        dest = Nothing

                    End If
                End If
            Next

        Catch ex As Exception
            'Known problem with security settings
            If ex.HResult = -2146827284 Then
                MessageBox.Show("Je zakázán programový přístup k VBA projektu. " &
                                "Pro zapnutí v aplikaci Excel:" &
                                Environment.NewLine &
                                "Soubor -> Možnosti -> Centrum zabezpečení -> Nastavení centra zabezpečení -> Nastavení maker -> Důvěřovat přístupu k objektovému modelu projektu VBA" &
                                Environment.NewLine & Environment.NewLine &
                                "Porovnání teď proběhne bez přenosu maker do výsledného sešitu.")
            Else

                MessageBox.Show(ex.Message & Environment.NewLine & ex.HResult)

            End If

        End Try

    End Sub


    'Copy formulas (that means also values) from dest sheet to tar sheet
    Sub CopySheetFormulas(dest As Excel.Worksheet, source As Excel.Worksheet)

        Try
            Dim rng As Excel.Range

            For Each cell As Excel.Range In source.UsedRange.Cells

                rng = dest.Range(GetExcelColumnName(cell.Column) & cell.Row)
                rng.Formula = cell.FormulaLocal

            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message & Environment.NewLine &
                            ex.StackTrace & Environment.NewLine)
        End Try
    End Sub

    'Creates "result" workbook
    Sub CreateResult()

        XlApp.DisplayAlerts = False

        'Vytvoří výstupní soubor   
        ResultWb = XlApp.Workbooks.Add
        XlApp.ActiveSheet.Name = "NewWbSheet"

        'Copies newSheet to ResultWb
        ResultWb.Worksheets.Add(After:=ResultWb.Sheets(1))
        NewResSheet = ResultWb.ActiveSheet
        NewResSheet.Name = OldSheet.Name
        CopySheetFormulas(ResultWb.ActiveSheet, NewSheet)

        'Goes through sheets in OldWb and copies them into ResultWb with exception of compared sheet
        For Each sheet As Excel.Worksheet In OldWb.Worksheets

            If sheet.Name <> OldSheet.Name Then
                ResultWb.Worksheets.Add(After:=ResultWb.Sheets(ResultWb.Sheets.Count))
                ResultWb.ActiveSheet.Name = sheet.Name

                CopySheetFormulas(ResultWb.ActiveSheet, sheet)

            End If

        Next

        'Copies oldSheet to ResultWb and renames to "Cancelled"
        ResultWb.Worksheets.Add(After:=ResultWb.Sheets(NewResSheet.Name))
        OldResSheet = ResultWb.ActiveSheet
        OldResSheet.Name = "Cancelled"
        CopySheetFormulas(OldResSheet, OldSheet)

        'Deletes sheet that gets automatically created when creating new workbook
        ResultWb.Sheets("NewWbSheet").Delete

        'Copies VBA project
        CopyMacros(ResultWb, OldWb)

        XlApp.DisplayAlerts = True

        ResultWb.Unprotect()
        NewResSheet.Unprotect()
        OldResSheet.Unprotect()

    End Sub

    'Deletes "found" rows from "Cancelled" sheet in "result" workbook
    Sub DeleteRows(sheet As Excel.Worksheet, indexArray() As Integer)

        For i As Integer = indexArray.Length - 1 To startRow Step -1

            If indexArray(i) = 1 Then

                sheet.Rows(i).EntireRow.Delete

            End If

        Next

    End Sub

    'Checks the number of columns in both sheets and lets the user to unite this
    Private Sub CheckColumns()

        If NewCols <> OldCols Then

            If MessageBox.Show("Rozdílný počet sloupců ve vybraných listech." &
                                Environment.NewLine &
                               "Chcete se pokusit o úpravu?",
                               "Close",
                               MessageBoxButtons.YesNo,
                               MessageBoxIcon.Question) = System.Windows.Forms.DialogResult.Yes Then

                MessageBox.Show("Následně se otevře aplikace pro přidání sloupců" &
                                Environment.NewLine &
                                "Přidejte sloupce tam, kde chybí, ale" &
                                Environment.NewLine &
                                Environment.NewLine &
                                "!!! NEZAVÍREJTE OKNO EXCELU !!!")

                XlApp.Visible = True

                '262144 makes the message box TopMost
                MsgBox("Hotovo?" & Environment.NewLine & "Doplnili jste všechny sloupce na správná místa?", 262144)

                XlApp.Visible = False

                AssignSheetsParams(FormSkompare.CBoxNewSheets.SelectedItem,
                                   FormSkompare.CBoxOldSheets.SelectedItem)

            End If

        End If

    End Sub

    'Deletes background color from OldResSheet
    Private Sub RemoveBackground(ws As Excel.Worksheet)

        For row As Integer = startRow To GetLast(ws, Excel.XlSearchOrder.xlByColumns)

            ws.Range("A" & row).EntireRow.Interior.Color = Excel.XlColorIndex.xlColorIndexNone

        Next

    End Sub

    'Gets starting points of tagged document
    Public Sub GetStart()

        If OldWb IsNot Nothing And NewWb IsNot Nothing Then
            AssignSheetsParams(FormSkompare.CBoxNewSheets.SelectedItem,
                            FormSkompare.CBoxOldSheets.SelectedItem)
        Else
            MessageBox.Show("Vyberte oba sešity i listy")
            Exit Sub
        End If

        Try

            FormSkompare.TBoxColSelect1.Text = GetExcelColumnName(NewSheet.Range("UID").Column)


            FormSkompare.ChBoxColSelect2.Checked = True
            FormSkompare.TBoxColSelect2.Enabled = True
            FormSkompare.TBoxColSelect2.ForeColor = SystemColors.WindowText
            FormSkompare.TBoxColSelect2.Text = GetExcelColumnName(NewSheet.Range("KKS_1").Column)


            FormSkompare.ChBoxColSelect3.Checked = True
            FormSkompare.TBoxColSelect3.Enabled = True
            FormSkompare.TBoxColSelect3.ForeColor = SystemColors.WindowText
            FormSkompare.TBoxColSelect3.Text = GetExcelColumnName(NewSheet.Range("KKS_2").Column)



            FormSkompare.TBoxStart.Text = NewSheet.Range("Header").Rows.Count + 1

        Catch

        End Try
    End Sub

End Class
