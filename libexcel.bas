Attribute VB_Name = "libexcel"
'@Folder("VBA.Extensions.libexcel")
'@ModuleDescription("This module was designed for Excel apps to extend functionality & to speed up developement. Don't import to other Office apps.")
Option Explicit

' References:
' [x] libcore (module)
' source: https://github.com/opi1101/VBA-extensions

Enum XlCellValueType
  Value
  Value2
  Text
End Enum

'@Description("Provides access to Excel application properties.")
Sub ApplicationSet(Optional bScreen As Variant, Optional bEvents As Variant, Optional bAlerts As Variant, _
Optional xlCalc As Variant, Optional bStatusBar As Variant)

  If (Not IsMissing(bScreen)) Then
    If VarType(bScreen) <> vbBoolean Then _
      Err.Raise 13, "libexcel.ApplicationSet", _
        StringMultiline("bScreen argument must be type of Boolean.", "Type: " & TypeName(bScreen))
    Application.ScreenUpdating = bScreen
  End If

  If (Not IsMissing(bEvents)) Then
    If VarType(bEvents) <> vbBoolean Then _
      Err.Raise 13, "libexcel.ApplicationSet", _
        StringMultiline("bEvents argument must be type of Boolean.", "Type: " & TypeName(bEvents))
    Application.EnableEvents = bEvents
  End If

  If (Not IsMissing(bAlerts)) Then
    If VarType(bAlerts) <> vbBoolean Then _
      Err.Raise 13, "libexcel.ApplicationSet", _
        StringMultiline("bAlerts argument must be type of Boolean.", "Type: " & TypeName(bAlerts))
    Application.DisplayAlerts = bAlerts
  End If

  If (Not IsMissing(bStatusBar)) Then
    If VarType(bStatusBar) <> vbBoolean Then _
      Err.Raise 13, "libexcel.ApplicationSet", _
        StringMultiline("bStatusBar argument must be type of Boolean.", "Type: " & TypeName(bStatusBar))
    Application.DisplayStatusBar = bStatusBar
  End If
  
  If (Not IsMissing(xlCalc)) Then
    If VarType(xlCalc) <> vbLong Then _
      Err.Raise 13, "libexcel.ApplicationSet", _
        StringMultiline("xlCalc argument must be a member of XlCalculation enum.", "Type: " & TypeName(xlCalc))
    Application.Calculation = xlCalc
  End If
End Sub

'@Description("Resets Excel application properties to their default values.")
Sub ApplicationReset()
  ApplicationSet bScreen:=True, bEvents:=True, bAlerts:=True, xlCalc:=XlCalculation.xlCalculationAutomatic, bStatusBar:=True
End Sub

'@Description("Sets Excel application properties for resource intensive tasks.")
Sub ApplicationPrepare()
  ApplicationSet bScreen:=False, bEvents:=False, bAlerts:=True, xlCalc:=XlCalculation.xlCalculationManual, bStatusBar:=True
End Sub

'@Description("Displays/hides ribbon, workbook tabs, formulabar, statusbar...")
Sub ApplicationToggleFullscreen(Optional Win As Excel.Window)
  If (Win Is Nothing) Then Set Win = ActiveWindow
  With Win
    .DisplayWorkbookTabs = Not .DisplayWorkbookTabs
    .DisplayRuler = Not .DisplayRuler
    .DisplayHeadings = Not .DisplayHeadings
  End With

  With Application
    .DisplayStatusBar = Not .DisplayStatusBar
    .DisplayScrollBars = Not .DisplayScrollBars
    .DisplayFormulaBar = Not .DisplayFormulaBar
    .CommandBars("Full Screen").Visible = Not .CommandBars("Full Screen").Visible
    .CommandBars("Worksheet Menu Bar").Enabled = Not .CommandBars("Worksheet Menu Bar").Enabled
    .ExecuteExcel4Macro "show.toolbar(""Ribbon""," & CStr(Not .ExecuteExcel4Macro("get.toolbar(7, ""Ribbon"")")) & ")"
  End With
End Sub

'@Description("Inserts a 2D array to target range. TargetCell is considered to be the desired range's top-left cell.")
Sub ArrayToRange(Arr2D As Variant, TargetCell As Excel.Range)
Dim lErr As Long
  If (TargetCell Is Nothing) Then _
    Err.Raise 91, "libexcel.ArrayToRange", "TargetCell argument is not set to an insance of an object."
  If (Not ArrayIsDimmed(Arr2D)) Then _
    Err.Raise 13, "libexcel.ArrayToRange", _
    StringMultiline("Arr2D is not assigned or not an array type.", "Type: " & TypeName(Arr2D))
  If RangeIsMultiArea(TargetCell) Then _
    Err.Raise 5, "libexcel.ArrayToRange", StringMultiline("Range with multiple areas is not supported.", _
    "Areas count: " & TargetCell.Areas.Count, _
    "TargetCell: " & TargetCell.Address)
  If (ArrayDimensionCount(Arr2D)) <> 2 Then _
    Err.Raise 13, "libexcel.ArrayToRange", _
    StringMultiline("Arr2D parameter must be a two dimensional array.", "Dimensions: " & ArrayDimensionCount(Arr2D))
  
  On Error Resume Next
  TargetCell.Cells(1, 1).Resize(UBound(Arr2D, 1), UBound(Arr2D, 2)).Value = Arr2D
  lErr = Err.Number
  On Error GoTo 0
  Select Case True
    Case lErr = 438
      Err.Raise 438, "libexcel.ArrayToRange", StringMultiline("Cannot insert objects in range.", _
      "TargetCell: " & TargetCell.Address)
    Case lErr <> 0
      Err.Raise lErr, "libexcel.ArrayToRange", StringMultiline("An error has occured while inserting array in range.", _
      "Error: " & Error(lErr), "TargetCell: " & TargetCell.Address)
  End Select
End Sub

'@Description("Inserts a 1D array to a row. Array will be inserted next to TargetCell (inclusive).")
Sub ArrayToRow(Arr1D As Variant, TargetCell As Excel.Range)
Dim lCount As Long, lErr As Long

  If (Not ArrayIsDimmed(Arr1D)) Then _
    Err.Raise 13, "libexcel.ArrayToRow", _
    StringMultiline("Array is not assigned or not an array type.", "Type: " & TypeName(Arr1D))
  If (Not ArrayIs1D(Arr1D)) Then _
    Err.Raise 5, "libexcel.ArrayToRow", _
    StringMultiline("Array with " & ArrayDimensionCount(Arr1D) & " is not supported.", "Dimensions: " & ArrayDimensionCount(Arr1D))
  
  lCount = ArrayDimensionLength(Arr1D)
  On Error Resume Next
  TargetCell.Cells(1, 1).Resize(ColumnSize:=lCount).Value = Arr1D
  lErr = Err.Number
  On Error GoTo 0
  Select Case True
    Case lErr = 438
      Err.Raise 438, "libexcel.ArrayToRow", StringMultiline("Cannot insert objects in row.", _
      "TargetCell: " & TargetCell.Address)
    Case lErr <> 0
      Err.Raise lErr, "libexcel.ArrayToRow", StringMultiline("An error has occured while inserting array in row.", _
      "Error: " & Error(lErr), "TargetCell: " & TargetCell.Address)
  End Select
End Sub

'@Description("Inserts a 1D array to a column. Array will be inserted under TargetCell (inclusive).")
Sub ArrayToColumn(Arr1D As Variant, TargetCell As Excel.Range)
Dim lCount As Long, lErr As Long

  If (Not ArrayIsDimmed(Arr1D)) Then _
    Err.Raise 13, "libexcel.ArrayToColumn", _
    StringMultiline("Array is not assigned or not an array type.", "Type: " & TypeName(Arr1D))
  If (Not ArrayIs1D(Arr1D)) Then _
    Err.Raise 5, "libexcel.ArrayToColumn", _
    StringMultiline("Array with " & ArrayDimensionCount(Arr1D) & " is not supported.", "Dimensions: " & ArrayDimensionCount(Arr1D))
    
  lCount = ArrayDimensionLength(Arr1D)
  On Error Resume Next
  TargetCell.Cells(1, 1).Resize(RowSize:=lCount).Value = Application.Transpose(Arr1D)
  lErr = Err.Number
  On Error GoTo 0
  Select Case True
    Case lErr = 438
      Err.Raise 438, "libexcel.ArrayToColumn", StringMultiline("Cannot insert objects in column.", _
      "TargetCell: " & TargetCell.Address)
    Case lErr <> 0
      Err.Raise lErr, "libexcel.ArrayToColumn", StringMultiline("An error has occured while inserting array in column.", _
      "Error: " & Error(lErr), "TargetCell: " & TargetCell.Address)
  End Select
End Sub

'@Description("Inserts a dictionary to a range. Key values will be inserted in TargetCell's column, Items will be inserted next to TargetCell's column.")
Sub DictionaryToRange(Dict As Object, TargetCell As Excel.Range)
Dim lErr As Long

  On Error Resume Next
  ArrayToColumn Dict.Keys(), TargetCell
  ArrayToColumn Dict.Items(), TargetCell.Offset(, 1)
  lErr = Err.Number
  On Error GoTo 0
  Select Case True
    Case lErr = 438
      Err.Raise 438, "libexcel.DictionaryToRange", StringMultiline("Cannot insert objects in range.", _
      "TargetCell: " & TargetCell.Address)
    Case lErr <> 0
      Err.Raise lErr, "libexcel.DictionaryToRange", StringMultiline("An error has occured while inserting dictionary in range.", _
      "Error: " & Error(lErr), "TargetCell: " & TargetCell.Address)
  End Select
End Sub

'@Description("Returns the specified Workbook object if it's already opened by the current application. Returns Nothing if it is not open.")
Function Workbooks2(ByVal Path As String) As Workbook
Dim Wbk As Workbook

  On Error Resume Next
  Path = PathToUNC(Path)
  Set Wbk = Workbooks(Dir(Path))
  On Error GoTo 0
  If (Not Wbk Is Nothing) Then _
    If LCase$(PathToUNC(Wbk.FullName)) = LCase$(Path) Then Set Workbooks2 = Wbk
  On Error GoTo 0
End Function

'@Description("Returns a Workbook object or Nothing. Opens the specified Workbook, does not prompt or raise an error if fails. ")
Function WorkbookOpenInstant(ByVal Path As String, ReadOnly As Boolean, _
Optional sPassword As Variant, Optional sWritePassword As Variant) As Excel.Workbook
Dim b As Boolean

  b = Application.DisplayAlerts
  Application.DisplayAlerts = False
  On Error Resume Next
  Select Case True
    Case (Not IsMissing(sPassword)), (Not IsMissing(sWritePassword))
      Set WorkbookOpenInstant = Workbooks.Open(Filename:=Path, ReadOnly:=ReadOnly, _
      Password:=sPassword, WriteResPassword:=sWritePassword)
    Case (Not IsMissing(sPassword))
      Set WorkbookOpenInstant = Workbooks.Open(Filename:=Path, ReadOnly:=ReadOnly, Password:=sPassword)
    Case (Not IsMissing(sWritePassword))
      Set WorkbookOpenInstant = Workbooks.Open(Filename:=Path, ReadOnly:=ReadOnly, WriteResPassword:=sWritePassword)
  End Select
  On Error GoTo 0
  Application.DisplayAlerts = b
End Function

'@Description("Returns a fully qualified macroname. Use with Application.Run to call other workbook's subroutines/functions.")
Function WorkbookMacroName(ByVal MacroName As String, Optional Wbk As Workbook) As String
  If (Wbk Is Nothing) Then Set Wbk = ThisWorkbook
  WorkbookMacroName = "'" & Wbk.FullName & "'!" & MacroName
End Function

'@Description("Returns an Excel workbook's lockfile path. Raises an error if Path file is unavailable.")
Function WorkbookLockfilePath(ByVal Path As String) As String
Dim sName As String

  If StringIsEmptyOrWhitespace(Path) Then _
    Err.Raise 53, "libexcel.WorkbookLockfilePath", StringMultiline("Filepath not provided.", "Filepath: " & Path)

  On Error Resume Next
  sName = Dir(Path, vbNormal)
  On Error GoTo 0
  
  If StrPtr(sName) = 0 Then _
    Err.Raise 53, "libexcel.WorkbookLockfilePath", StringMultiline("File not found or unavailable.", "Filepath: " & Path)

  WorkbookLockfilePath = PathCombine(PathParentDirectory(Path), "~$" & sName)
End Function

'@Description("Returns the owner of a workbook's lockfile in domain\username format. Raises an error if Path file is unavailable.")
Function WorkbookLockedBy(ByVal Path As String) As String
  On Error Resume Next
  WorkbookLockedBy = FileOwner(WorkbookLockfilePath(Path))
End Function

'@Description("Clears all filters in the specified workbook object.")
Sub WorkbookClearAllFilters(Optional Wbk As Excel.Workbook)
Dim Sht As Excel.Worksheet

  If Wbk Is Nothing Then Set Wbk = ThisWorkbook
  For Each Sht In Wbk.Worksheets
    WorksheetClearAllFilters Sht, Wbk
  Next Sht
End Sub

'@Description("Sets a workbook's BuiltinDocumentProperties.")
Sub WorkbookSetBuiltinDocumentProperties(Wbk As Excel.Workbook, Optional sTitle As String, Optional sAuthor As String, Optional sAppName As String, _
Optional sCompany As String, Optional sComments As String, Optional sKeywords As String, Optional sCategory As String)
  With Wbk.BuiltinDocumentProperties
    If sTitle <> vbNullString Then .Item("Title") = sTitle
    If sAuthor <> vbNullString Then .Item("Author") = sAuthor
    If sAppName <> vbNullString Then .Item("Application Name") = sAppName
    If sCompany <> vbNullString Then .Item("Company") = sCompany
    If sComments <> vbNullString Then .Item("Comments") = sComments
    If sKeywords <> vbNullString Then .Item("Keywords") = sKeywords
    If sCategory <> vbNullString Then .Item("Category") = sCategory
  End With
End Sub

'@Description("Returns the specified Worksheet object. Raises an error if Worksheet does not exist.")
Function Worksheets2(Sht As Variant, Optional Wbk As Excel.Workbook) As Excel.Worksheet
Dim lErr As Long

  If (Wbk Is Nothing) Then Set Wbk = ThisWorkbook
  Select Case VarType(Sht)
    Case vbString, vbInteger, vbLong
      On Error Resume Next
      Set Worksheets2 = Wbk.Worksheets(Sht)
      If Err.Number = 9 Then
        On Error GoTo 0
        Err.Raise 9, "libexcel.Worksheets2", StringMultiline("Worksheet not found.", "Workbook: " & Wbk.Name, "Sheet: " & Sht, "Type: " & TypeName(Sht))
      End If
      On Error GoTo 0
    Case vbObject
      If (Not TypeOf Sht Is Worksheet) Then _
        Err.Raise 13, "libexcel.Worksheets2", _
          StringMultiline("Sht argument must be type of Worksheet, String, Integer or Long.", "Type: " & TypeName(Sht))
      Set Worksheets2 = Sht
    Case Else
      Err.Raise 13, "libexcel.Worksheets2", _
        StringMultiline("Sht argument must be type of Worksheet, String, Integer or Long.", "Type: " & TypeName(Sht))
  End Select
End Function

'@Description("Deletes specified worksheet, does not prompt user for deletion.")
Sub WorksheetDeleteInstant(Sht As Variant, Optional Wbk As Excel.Workbook)
Dim b As Boolean

  If (Wbk Is Nothing) Then Set Wbk = ThisWorkbook
  b = Wbk.Application.DisplayAlerts
  Wbk.Application.DisplayAlerts = False
  Worksheets2(Sht, Wbk).Delete
  Wbk.Application.DisplayAlerts = b
End Sub

'@Description("Returns True if specified Worksheet exists in a Workbook.")
Function WorksheetExists(Sht As Variant, Optional Wbk As Excel.Workbook) As Boolean
Dim shtTmp As Excel.Worksheet

  If (Wbk Is Nothing) Then Set Wbk = ThisWorkbook
  On Error Resume Next
  Set shtTmp = Worksheets2(Sht, Wbk)
  On Error GoTo 0
  WorksheetExists = (Not shtTmp Is Nothing)
End Function

'@Description("Clears all filters on a worksheet.")
Sub WorksheetClearAllFilters(Sht As Variant, Optional Wbk As Excel.Workbook)
  If (Wbk Is Nothing) Then Set Wbk = ThisWorkbook
  With Worksheets2(Sht, Wbk)
    On Error Resume Next
    .ShowAllData
    .AutoFilter.ShowAllData
  End With
End Sub

'@Description("Removes empty rows on a given worksheet.")
Sub WorksheetDeleteEmptyRows(Sht As Variant, Optional Wbk As Excel.Workbook)
  RangeDeleteEmptyRows Worksheets2(Sht, Wbk).Cells
End Sub

'@Description("Returns the extended worksheet UsedRange property from cell A1.")
Function WorksheetUsedRange2(Sht As Variant, Optional Wbk As Excel.Workbook) As Excel.Range
  If (Wbk Is Nothing) Then Set Wbk = ThisWorkbook
  Set WorksheetUsedRange2 = Sht.Range("A1", Sht.UsedRange)
End Function

'@Description("Returns the last used cell in a given column. Uses End(xlUp) method.")
Function WorksheetColumnLastCell(Column As Variant, Sht As Variant, Optional Wbk As Excel.Workbook) As Excel.Range
  If (Wbk Is Nothing) Then Set Wbk = ThisWorkbook
  Select Case VarType(Column)
    Case vbLong, vbInteger, vbString
    Case Else
      Err.Raise 13, "libexcel.WorksheetColumnLastCell", _
        StringMultiline("Column argument must be type of String, Integer or Long.", "Type: " & TypeName$(Column))
  End Select
  With Worksheets2(Sht, Wbk)
    Set WorksheetColumnLastCell = .Cells(.Rows.Count, Column).End(xlUp)
  End With
End Function

'@Description("Returns the last used cell in a given row. Uses End(xlToLeft) method.")
Function WorksheetRowLastCell(Row As Variant, Sht As Variant, Optional Wbk As Excel.Workbook) As Excel.Range
  If Wbk Is Nothing Then Set Wbk = ThisWorkbook
  Select Case VarType(Row)
    Case vbLong, vbInteger, vbString
    Case Else
      Err.Raise 13, "libexcel.WorksheetRowLastCell", _
        StringMultiline("Row argument must be type of String, Integer or Long.", "Type: " & TypeName$(Row))
  End Select
  With Worksheets2(Sht, Wbk)
    Set WorksheetRowLastCell = .Cells(Row, .Columns.Count).End(xlToLeft)
  End With
End Function

'@Description("Returns the first cell on a sheet containing a value. Returns Nothing if no such value was found.")
Function WorksheetFindValue(What As Variant, Sht As Variant, Optional Wbk As Excel.Workbook) As Excel.Range
  Set WorksheetFindValue = RangeFindValue(What, Worksheets2(Sht, Wbk).Cells)
End Function

'@Description("Returns the last used row on a worksheet. Uses Range.Find method.")
Function WorksheetFindLastRow(Sht As Variant, Optional Wbk As Excel.Workbook) As Excel.Range
  Set WorksheetFindLastRow = RangeFindLastRow(Worksheets2(Sht, Wbk).Cells)
End Function

'@Description("Returns the last used column on a worksheet. Uses Range.Find method.")
Function WorksheetFindLastColumn(Sht As Variant, Optional Wbk As Excel.Workbook) As Excel.Range
  Set WorksheetFindLastColumn = RangeFindLastColumn(Worksheets2(Sht, Wbk).Cells)
End Function

'@Description("Returns the last used cell on a worksheet. Uses Range.Find method.")
Function WorksheetFindLastCell(Sht As Variant, Optional Wbk As Excel.Workbook) As Excel.Range
  Set WorksheetFindLastCell = RangeFindLastCell(Worksheets2(Sht, Wbk).Cells)
End Function

'@Description("Returns a worksheet's parent workbook object.")
Function WorksheetWorkbook(Sht As Variant, Optional Wbk As Excel.Workbook) As Excel.Workbook
  Set WorksheetWorkbook = Worksheets2(Sht, Wbk).Parent
End Function

'@Description("Returns True if each cell in a range is empty. Works if range has multiple areas.")
Function RangeIsEmpty(Target As Excel.Range) As Boolean
  If (Target Is Nothing) Then _
    Err.Raise 91, "libexcel.RangeIsEmpty", "Target argument is not set to an insance of an object."
  RangeIsEmpty = (WorksheetFunction.CountA(Target) = 0)
End Function

'@Description("Returns True if a range has multiple areas (non-contiguous range).")
Function RangeIsMultiArea(Target As Excel.Range) As Boolean
  If (Target Is Nothing) Then _
    Err.Raise 91, "libexcel.RangeIsMultiArea", "Target argument is not set to an insance of an object."
  RangeIsMultiArea = (Target.Areas.Count > 1)
End Function

'@Description("Returns a range total row count. Works if range has multiple areas.")
Function RangeRowsCount(Target As Excel.Range) As Long
Dim rn As Range
Dim d As Object

  Set d = NewDictionaryObject
  For Each rn In Target.Rows
    If (Not d.Exists(rn.Row)) Then d.Add rn.Row, 0
  Next rn
  RangeRowsCount = d.Count
End Function

'@Description("Returns a range total column count. Works if range has multiple areas.")
Function RangeColumnsCount(Target As Excel.Range) As Long
Dim rn As Range
Dim d As Object

  Set d = NewDictionaryObject
  For Each rn In Target.Columns
    If (Not d.Exists(rn.Column)) Then d.Add rn.Column, 0
  Next rn
  RangeColumnsCount = d.Count
End Function

'@Description("Returns a range where header row(s) from the top are excluded from table range.")
Function RangeDataBody(Target As Excel.Range, Optional HeaderRowCount As Long = 1) As Excel.Range
  If (Target Is Nothing) Then _
    Err.Raise 91, "libexcel.RangeDataBody", "Target argument is not set to an insance of an object."
  If RangeIsMultiArea(Target) Then _
    Err.Raise 5, "libexcel.RangeDataBody", StringMultiline("Range with multiple areas is not supported.", _
    "Areas count: " & Target.Areas.Count, _
    "Target: " & Target.Address)
  If HeaderRowCount < 1 Then _
    Err.Raise 9, "libexcel.RangeDataBody", StringMultiline("HeaderRowCount must be greater zero.", _
    "HeaderRowCount: " & HeaderRowCount)
  If HeaderRowCount >= Target.Rows.Count Then _
    Err.Raise 9, "libexcel.RangeDataBody", StringMultiline("HeaderRowCount must be greater than Target total row count.", _
    "HeaderRowCount: " & HeaderRowCount, "Target rows count: " & Target.Rows.Count, "Target: " & Target.Address)

  With Target
    Set RangeDataBody = .Offset(HeaderRowCount).Resize(.Rows.Count - HeaderRowCount)
  End With
End Function

'@Description("Returns the first cell in a range containing a value. Returns Nothing if no such value was found.")
Function RangeFindValue(What As Variant, Target As Excel.Range) As Excel.Range
  If (Target Is Nothing) Then _
    Err.Raise 91, "libexcel.RangeFindValue", "Target argument is not set to an insance of an object."
  If VarType(What) = vbObject Then _
    Err.Raise 438, "libexcel.RangeFindValue", "What argument cannot be object type."
    
  On Error Resume Next
  Set RangeFindValue = Target.Find(What, Target.Cells(1, 1), _
    XlFindLookIn.xlFormulas, XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious, False)
  On Error GoTo 0
End Function

'@Description("Returns the last row in a range containing any value. Returns Nothing if no such cell was found. Works if range has multiple areas.")
Function RangeFindLastRow(Target As Excel.Range) As Excel.Range
  If (Target Is Nothing) Then _
    Err.Raise 91, "libexcel.RangeFindLastRow", "Target argument is not set to an insance of an object."

  On Error Resume Next
  Set RangeFindLastRow = Target.Find("*", Target.Cells(1, 1), _
    XlFindLookIn.xlFormulas, XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious, False)
  On Error GoTo 0
  
  If (Not RangeFindLastRow Is Nothing) Then _
    Set RangeFindLastRow = Target.Rows(RangeFindLastRow.Rows(1).Row - Target.Rows(1).Row + 1)
End Function

'@Description("Returns the last column in a range containing any value. Returns Nothing if no such cell was found. Works if range has multiple areas.")
Function RangeFindLastColumn(Target As Excel.Range) As Excel.Range
  If (Target Is Nothing) Then _
    Err.Raise 91, "libexcel.RangeFindLastColumn", "Target argument is not set to an insance of an object."

  On Error Resume Next
  Set RangeFindLastColumn = Target.Find("*", Target.Cells(1, 1), _
    XlFindLookIn.xlFormulas, XlLookAt.xlPart, XlSearchOrder.xlByColumns, XlSearchDirection.xlPrevious, False)
  On Error GoTo 0
  
  If (Not RangeFindLastColumn Is Nothing) Then _
    Set RangeFindLastColumn = Target.Columns(RangeFindLastColumn.Columns(1).Column - Target.Columns(1).Column + 1)
End Function

'@Description("Returns the last cell with any value in a given range. Returns Nothing if no such cell was found. Works if range has multiple areas.")
Function RangeFindLastCell(Target As Excel.Range) As Excel.Range
Dim lRow As Long, lCol As Long

  If (Target Is Nothing) Then _
    Err.Raise 91, "libexcel.RangeFindLastCell", "Target argument is not set to an insance of an object."
  
  On Error Resume Next
  lRow = Target.Find("*", Target.Cells(1, 1), _
    XlFindLookIn.xlFormulas, XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious, False).Row
  lCol = Target.Find("*", Target.Cells(1, 1), _
    XlFindLookIn.xlFormulas, XlLookAt.xlPart, XlSearchOrder.xlByColumns, XlSearchDirection.xlPrevious, False).Column
  On Error GoTo 0
  
  If (lRow < 1) Or (lCol < 1) Then Exit Function
  Set RangeFindLastCell = Target.Worksheet.Cells(lRow, lCol)
End Function

'@Description("Removes empty rows in a given range.")
Sub RangeDeleteEmptyRows(Target As Excel.Range)
Dim lArea As Long, lRow As Long

  If (Target Is Nothing) Then _
    Err.Raise 91, "libexcel.RangeDeleteEmptyRows", "Target argument is not set to an insance of an object."
    
  For lArea = Target.Areas.Count To 1 Step -1
    With Target.Areas(lArea)
      For lRow = .Rows.Count To 1 Step -1
        If WorksheetFunction.CountA(.Rows(lRow)) = 0 Then .Rows(lRow).Delete
      Next lRow
    End With
  Next lArea
End Sub

'@Description("Returns the resized range. RowSize represents the row count to cut from the top, ColumnSize determines how many columns will be cut from the left.")
Function RangeResize2(Target As Excel.Range, Optional RowSize As Long, Optional ColumnSize As Long) As Excel.Range
  If (Target Is Nothing) Then _
    Err.Raise 91, "libexcel.RangeResize2", "Target argument is not set to an insance of an object."
  If RangeIsMultiArea(Target) Then _
    Err.Raise 5, "libexcel.RangeResize2", StringMultiline("Range with multiple areas is not supported.", _
    "Areas count: " & Target.Areas.Count, _
    "Target: " & Target.Address)
  If RowSize >= Target.Rows.Count Then _
    Err.Raise 9, "libexcel.RangeResize2", StringMultiline("RowSize must be less than Target total row count.", _
    "RowSize: " & RowSize, "Target: " & Target.Rows.Count)
  If ColumnSize >= Target.Columns.Count Then _
    Err.Raise 9, "libexcel.RangeResize2", StringMultiline("ColumnSize must be less than Target total column count.", _
    "ColumnSize: " & ColumnSize, "Target: " & Target.Columns.Count)
    
  With Target
    Set RangeResize2 = .Offset(RowSize, ColumnSize).Resize(.Rows.Count - RowSize, .Columns.Count - ColumnSize)
  End With
End Function

'@Description("Returns a dictionary with unique key values from the given range.")
Function RangeToDictionary(Target As Excel.Range, Optional ByVal KeyColumn As Variant, Optional ByVal ItemColumn As Variant, _
Optional CellValue As XlCellValueType = XlCellValueType.Value, Optional LoopStep As Long = 1, Optional YieldAt As Long) As Object
Dim rn As Range
  
  If (Target Is Nothing) Then _
    Err.Raise 91, "libexcel.RangeToDictionary", "Target argument is not set to an insance of an object."
  If RangeIsMultiArea(Target) Then _
    Err.Raise 5, "libexcel.RangeToDictionary", StringMultiline("Range with multiple areas is not supported.", _
    "Areas count: " & Target.Areas.Count, _
    "Target: " & Target.Address)
  If LoopStep < 1 Then _
    Err.Raise 9, "libexcel.RangeToDictionary", StringMultiline("LoopStep argument must be bigger than zero.", _
    "LoopStep: " & LoopStep)
  If YieldAt < 0 Then _
    Err.Raise 9, "libexcel.RangeToDictionary", StringMultiline("YieldAt argument must be equal or greater than zero.", _
    "YieldAt: " & YieldAt)
  If (Not VariantIsValidColumnID(KeyColumn)) Then _
    Err.Raise 13, "libexcel.RangeToDictionary", StringMultiline("KeyColumn argument is not a valid column identifier.", _
    "KeyColumn: " & CStr(KeyColumn), "Type: " & TypeName$(KeyColumn))
  KeyColumn = rn.Column
  If (Not VariantIsValidColumnID(ItemColumn)) Then _
    Err.Raise 13, "libexcel.RangeToDictionary", StringMultiline("ItemColumn argument is not a valid column identifier.", _
    "ItemColumn: " & CStr(ItemColumn), "Type: " & TypeName$(ItemColumn))
  ItemColumn = rn.Column

  On Error GoTo 0
  Select Case CellValue
    Case XlCellValueType.Value
      Set RangeToDictionary = ArrayToDictionary(Target.Value, CLng(KeyColumn), CLng(ItemColumn), LoopStep, YieldAt)
    Case XlCellValueType.Value2
      Set RangeToDictionary = ArrayToDictionary(Target.Value2, CLng(KeyColumn), CLng(ItemColumn), LoopStep, YieldAt)
    Case XlCellValueType.Text
      Err.Raise 5, "libexcel.RangeToDictionary", StringMultiline("Range text value is not supported.", _
      "Target: " & Target.Address)
  End Select
End Function

'@Description("Returns a dictionary where unique key values are the specified column values or row numbers, items are the row values as an array.")
Function RangeRowsToDictionary(Target As Excel.Range, Optional ByVal KeyColumn As Variant, _
Optional CellValue As XlCellValueType = XlCellValueType.Value, Optional LoopStep As Long = 1, Optional YieldAt As Long) As Object
Dim rn As Range
Dim d As Object
Dim lStep As Long, lCount As Long
Dim bColVal As Boolean
Dim Key As Variant

  If (Target Is Nothing) Then _
    Err.Raise 91, "libexcel.RangeRowsToDictionary", "Target argument is not set to an insance of an object."
  If LoopStep < 1 Then _
    Err.Raise 9, "libexcel.RangeRowsToDictionary", StringMultiline("LoopStep argument must be bigger than zero.", _
    "LoopStep: " & LoopStep)
  If YieldAt < 0 Then _
    Err.Raise 9, "libexcel.RangeRowsToDictionary", StringMultiline("YieldAt argument must be equal or greater than zero.", _
    "YieldAt: " & YieldAt)

  bColVal = (Not IsMissing(KeyColumn)) ' Dict. key is based on a row's column value

  If bColVal Then _
    If (Not VariantIsValidColumnID(KeyColumn)) Then _
      Err.Raise 13, "libexcel.RangeRowsToDictionary", StringMultiline("KeyColumn argument is not a valid column identifier.", _
      "KeyColumn: " & CStr(KeyColumn), "Type: " & TypeName$(KeyColumn))
  
  Set d = NewDictionaryObject
  For Each rn In Target.Rows
    lCount = lCount + 1
    lStep = lStep + 1
    If lStep = LoopStep Then
      lStep = 0
      If bColVal Then
        Key = rn.Columns(KeyColumn).Value
      Else
        Key = rn.Row
      End If
      If (Not d.Exists(Key)) Then
        Select Case CellValue
          Case XlCellValueType.Value
            d.Add Key, rn.Value
          Case XlCellValueType.Value2
            d.Add Key, rn.Value2
          Case XlCellValueType.Text
            Err.Raise 5, "libexcel.RangeRowsToDictionary", StringMultiline("Range text value is not supported.", _
            "Target: " & Target.Address)
        End Select
        End If
    End If
    If YieldAt > 0 Then _
      If lCount Mod YieldAt = 0 Then DoEvents
  Next rn
  Set RangeRowsToDictionary = d
End Function

'@Description("Returns a row's values as a 1D array.")
Function RangeRowToArray1D(TargetRow As Excel.Range, Optional CellValue As XlCellValueType = XlCellValueType.Value) As Variant
  If (TargetRow Is Nothing) Then _
    Err.Raise 91, "libexcel.RangeRowToArray1D", "TargetRow argument is not set to an insance of an object."
  If RangeIsMultiArea(TargetRow) Then _
    Err.Raise 5, "libexcel.RangeRowToArray1D", StringMultiline("Range with multiple areas is not supported.", _
    "Areas count: " & TargetRow.Areas.Count, _
    "TargetRow: " & TargetRow.Address)
  If TargetRow.Rows.Count <> 1 Then _
    Err.Raise 5, "libexcel.RangeRowToArray1D", StringMultiline("TargetRow with multiple rows is not supported.", _
    "Row count: " & TargetRow.Rows.Count, _
    "TargetRow: " & TargetRow.Address)

  Select Case CellValue
    Case XlCellValueType.Value
      RangeRowToArray1D = Application.Transpose(Application.Transpose(TargetRow.Value))
    Case XlCellValueType.Value2
      RangeRowToArray1D = Application.Transpose(Application.Transpose(TargetRow.Value2))
    Case XlCellValueType.Text
      Err.Raise 5, "libexcel.RangeRowToArray1D", StringMultiline("Range text value is not supported.", _
      "TargetRow: " & TargetRow.Address)
  End Select
End Function

'@Description("Returns a column's values as a 1D array.")
Function RangeColumnToArray1D(TargetCol As Excel.Range, Optional CellValue As XlCellValueType = XlCellValueType.Value) As Variant
  If (TargetCol Is Nothing) Then _
    Err.Raise 91, "libexcel.RangeColumnToArray1D", "TargetCol argument is not set to an insance of an object."
  If RangeIsMultiArea(TargetCol) Then _
    Err.Raise 5, "libexcel.RangeColumnToArray1D", StringMultiline("Range with multiple areas is not supported.", _
    "Areas count: " & TargetCol.Areas.Count, _
    "TargetCol: " & TargetCol.Address)
  If TargetCol.Columns.Count <> 1 Then _
    Err.Raise 5, "libexcel.RangeColumnToArray1D", StringMultiline("TargetCol with multiple columns is not supported.", _
    "Column count: " & TargetCol.Columns.Count, _
    "TargetCol: " & TargetCol.Address)

  Select Case CellValue
    Case XlCellValueType.Value
      RangeColumnToArray1D = Application.Transpose(TargetCol.Value)
    Case XlCellValueType.Value2
      RangeColumnToArray1D = Application.Transpose(TargetCol.Value2)
    Case XlCellValueType.Text
      Err.Raise 5, "libexcel.RangeColumnToArray1D", StringMultiline("Range text value is not supported.", _
      "TargetCol: " & TargetCol.Address)
  End Select
End Function

'@Description("Returns True if both parameter range's first row (header) has equal values. Uses lower case comparsion.")
Function RangeHeadersAreEqual(Table1 As Excel.Range, Table2 As Excel.Range) As Boolean
Dim h1, h2

  On Error Resume Next
  h1 = Join(RangeRowToArray1D(Table1.Rows(1)))
  h2 = Join(RangeRowToArray1D(Table2.Rows(1)))
  On Error GoTo 0
  RangeHeadersAreEqual = (LCase(h1) = LCase(h2))
End Function

'@Description("Centers range on screen. Make sure worksheet is visible or workbook structure protection is turned off.")
Sub RangeCenterScreen(Target As Excel.Range)
  If (Not Target.Worksheet Is ActiveSheet) Then
    If WorksheetWorkbook(Target.Worksheet).ProtectStructure Then _
      Err.Raise 5, "libexcel.RangeCenterScreen", _
      StringMultiline("Cannot center range, worksheet is not visible and workbook structure is protected.", _
      "Workbook: " & WorksheetWorkbook(Target.Worksheet).Name, "Sheet: " & Target.Worksheet.Name, "Range: " & Target.Address)
    Target.Worksheet.Visible = True
    Target.Worksheet.Activate
  End If
  Application.Goto Reference:=Target, Scroll:=True
  With ActiveWindow
    .SmallScroll Up:=.VisibleRange.Rows.Count / 2, ToLeft:=.VisibleRange.Columns.Count / 2
  End With
End Sub

'@Description("Returns True if argument is a valid column.")
Function VariantIsValidColumnID(Col As Variant) As Boolean
Dim rn As Range

  On Error Resume Next
  Set rn = Columns(Col)
  VariantIsValidColumnID = (Err.Number = 0)
End Function

'@Description("Sets the visibility of the gridlines on a worksheet.")
Sub WindowsSetGridLines(Sht As Variant, Display As Boolean, Optional Wbk As Excel.Workbook)
Dim vw As WorksheetView

  For Each vw In WorksheetWorkbook(Sht, Wbk).Windows(1).SheetViews
    If vw.Sheet.Name = Worksheets2(Sht, Wbk).Name Then
      vw.DisplayGridlines = Display
      Exit Sub
    End If
  Next vw
  Set vw = Nothing
End Sub
