Attribute VB_Name = "libcore"
'@Folder("VBA.Extensions.libcore")
'@ModuleDescription("This module was designed for VBA apps to extend functionality & to speed up developement.")
Option Explicit

' References: none
' source: https://github.com/opi1101/VBA-extensions

'@Description("Returns true, if parameter Application is MS Excel application.")
Function ApplicationIsExcel(App As Application) As Boolean
  On Error Resume Next
  ApplicationIsExcel = (InStr(App.Name, "Excel") > 0)
End Function

'@Description("Returns true, if parameter Application is MS Access application.")
Function ApplicationIsAccess(App As Application) As Boolean
  On Error Resume Next
  ApplicationIsAccess = (InStr(App.Name, "Access") > 0)
End Function

'@Description("Returns true, if the argument has lower bound (=array, declared and assigned).")
Function ArrayIsDimmed(Arr As Variant) As Boolean
Dim x As Long

  On Error Resume Next
  x = LBound(Arr)
  ArrayIsDimmed = (Err.Number = 0)
End Function

'@Description("Returns the number of dimensions of an array. Raises an error if the argument is not an array type.")
Function ArrayDimensionCount(Arr As Variant) As Long
Dim iCount As Integer
Dim lSize As Long

  If (Not ArrayIsDimmed(Arr)) Then _
    Err.Raise 13, "libcore.ArrayDimensionCount", _
    StringMultiline("Array is not assigned or not an array type.", "Type: " & TypeName(Arr))

  On Error Resume Next
  Do
    iCount = iCount + 1
    lSize = UBound(Arr, iCount)
  Loop Until Err.Number <> 0
  On Error GoTo 0
  ArrayDimensionCount = (iCount - 1)
End Function

'@Description("Returns the number of elements in the array's specific dimension. TargetDim can be omitted if array is 1D.")
Function ArrayDimensionLength(Arr As Variant, Optional ByVal TargetDim As Long) As Long
Dim lDim As Long

  If (Not ArrayIsDimmed(Arr)) Then _
    Err.Raise 13, "libcore.ArrayDimensionLength", _
    StringMultiline("Array is not assigned or not an array type.", "Type: " & TypeName(Arr))

  lDim = ArrayDimensionCount(Arr)
  If lDim = 1 Then TargetDim = 1
  
  If TargetDim < 1 Or TargetDim > lDim Then _
    Err.Raise 9, "libcore.ArrayDimensionLength", _
    StringMultiline("TargetDim is out of bounds.", _
    "Array dimension count: " & lDim, "TargetDim: " & TargetDim)
  ArrayDimensionLength = Abs(LBound(Arr, TargetDim) - UBound(Arr, TargetDim)) + 1
End Function

'@Description("Returns True if an array has only one dimension. Raises an error if the argument is not an array type.")
Function ArrayIs1D(Arr As Variant) As Boolean
  If (Not ArrayIsDimmed(Arr)) Then _
    Err.Raise 13, "libcore.ArrayIs1D", StringMultiline("Array is not assigned or not an array type.", _
      "Type: " & TypeName(Arr))
  ArrayIs1D = (ArrayDimensionCount(Arr) = 1)
End Function

'@Description("Returns True if an array has exactly two dimensions. Raises an error if the argument is not an array type.")
Function ArrayIs2D(Arr As Variant) As Boolean
  If (Not ArrayIsDimmed(Arr)) Then _
    Err.Raise 13, "libcore.ArrayIs2D", StringMultiline("Array is not assigned or not an array type.", _
      "Type: " & TypeName(Arr))
  ArrayIs2D = (ArrayDimensionCount(Arr) = 2)
End Function

'@Description("Resizes the specified dimension of a 1D or 2D array. Resizes the original array, keeps values with ReDim Preserve keywords.")
Sub ArrayResizeDimension(Arr As Variant, TargetDim As Integer, DimNewSize As Long)
Dim lDim As Long
Dim bIsExcel As Boolean
Dim xlApp As Object

  On Error GoTo ERRH
  bIsExcel = ApplicationIsExcel(Application)
  lDim = ArrayDimensionCount(Arr)
  Select Case lDim
    Case 1 ' 1D array
      If DimNewSize < LBound(Arr) Then _
        Err.Raise 9, "libcore.ArrayResizeDimension", StringMultiline("DimNewSize cannot be less than the lower bound of the array.", _
        "Array lower bound: " & LBound(Arr), _
        "DimNewSize: " & DimNewSize)
      ReDim Preserve Arr(LBound(Arr) To DimNewSize)
    Case 2 ' 2D array
      Select Case TargetDim
        Case 1
          If bIsExcel Then
            Set xlApp = Application
          Else
            Set xlApp = NewExcelObject
          End If
          If DimNewSize < LBound(Arr, 1) Then _
            Err.Raise 9, "libcore.ArrayResizeDimension", StringMultiline("DimNewSize cannot be less than the lower bound of the 1st dimension.", _
            "Array 1st dim. lower bound: " & LBound(Arr, 1), _
            "DimNewSize: " & DimNewSize)
          Arr = xlApp.Transpose(Arr)
          ReDim Preserve Arr(LBound(Arr, 1) To UBound(Arr, 1), LBound(Arr, 2) To DimNewSize)
          Arr = xlApp.Transpose(Arr)
          If (Not bIsExcel) Then xlApp.Quit
        Case 2
          If DimNewSize < LBound(Arr, 2) Then _
            Err.Raise 9, "libcore.ArrayResizeDimension", StringMultiline("DimNewSize cannot be less than the lower bound of the 2nd dimension.", _
            "Array 2nd dim. lower bound: " & LBound(Arr, 2), _
            "DimNewSize: " & DimNewSize)
          ReDim Preserve Arr(LBound(Arr, 1) To UBound(Arr, 1), LBound(Arr, 2) To DimNewSize)
        Case Else
          Err.Raise 9, "libcore.ArrayResizeDimension", StringMultiline("TargetDim argument is out of bounds.", _
            "Array dimension count: " & lDim, _
            "Target dimension: " & TargetDim)
      End Select
    Case Else ' More than 2 dimensions
      Err.Raise 5, "libcore.ArrayResizeDimension", "Array with " & lDim & " dimensions is not supported."
  End Select
Exit Sub
ERRH:
  If (Not bIsExcel) Then xlApp.Quit
  Err.Raise Err.Number, "libcore.ArrayResizeDimension", Err.Description
End Sub

'@Description("Fills a dictionary with 1D or 2D array values. Key- and ValueColumns must be members of the 2nd dimension of a 2D array.")
Function ArrayToDictionary(Arr As Variant, Optional KeyColumn As Long, Optional ItemColumn As Long, _
  Optional LoopStep As Long = 1, Optional YieldAt As Long) As Object
Dim d As Object
Dim lDim As Long, x As Long

  If (Not ArrayIsDimmed(Arr)) Then _
    Err.Raise 13, "libcore.ArrayToDictionary", StringMultiline("Array is not assigned or not an array type.", _
    "Type: " & TypeName(Arr))
  If LoopStep < 1 Then _
    Err.Raise 9, "libcore.ArrayToDictionary", StringMultiline("LoopStep argument must be bigger than zero.", _
    "LoopStep: " & LoopStep)
  If YieldAt < 0 Then _
    Err.Raise 9, "libcore.ArrayToDictionary", StringMultiline("YieldAt argument must be equal or greater than zero.", _
    "YieldAt: " & YieldAt)

  lDim = ArrayDimensionCount(Arr)
  Set d = NewDictionaryObject
  
  Select Case lDim
    Case 1
      For x = LBound(Arr) To UBound(Arr) Step LoopStep
        Select Case True
          Case IsEmpty(Arr(x)), IsError(Arr(x)), Arr(x) = vbNullString
          Case Else
            If (Not d.Exists(Arr(x))) Then d.Add Arr(x), x
        End Select
        If YieldAt > 0 Then _
          If x Mod YieldAt = 0 Then DoEvents
      Next x
    Case 2
      Select Case True
        Case KeyColumn < LBound(Arr, 2), KeyColumn > UBound(Arr, 2)
          Err.Raise 9, "libcore.ArrayToDictionary", StringMultiline("KeyColumn is outside the bounds of the array's 2nd dimension.", _
            "KeyColumn: " & KeyColumn, "2nd dimension LBound-UBound: " & LBound(Arr, 2) & "-" & UBound(Arr, 2))
        Case ItemColumn < LBound(Arr, 2), ItemColumn > UBound(Arr, 2)
          Err.Raise 9, "libcore.ArrayToDictionary", StringMultiline("ItemColumn is outside the bounds of the array's 2nd dimension.", _
            "ItemColumn: " & ItemColumn, "2nd dimension LBound-UBound: " & LBound(Arr, 2) & "-" & UBound(Arr, 2))
      End Select
      For x = LBound(Arr, 1) To UBound(Arr, 1) Step LoopStep
        Select Case True
          Case IsEmpty(Arr(x, KeyColumn)), IsError(Arr(x, KeyColumn)), Arr(x, KeyColumn) = vbNullString
          Case Else
            If (Not d.Exists(Arr(x, KeyColumn))) Then d.Add Arr(x, KeyColumn), Arr(x, ItemColumn)
        End Select
        If YieldAt > 0 Then _
          If x Mod YieldAt = 0 Then DoEvents
      Next x
    Case Else
      Err.Raise 5, "libcore.ArrayToDictionary", "Arrays with " & lDim & " dimensions is not supported."
  End Select
  Set ArrayToDictionary = d
End Function

'@Description("Writes an array to a csv file. Changes/adds extension to Filepath argument.")
Sub ArrayToCSV(Arr As Variant, ByVal Filepath As String, Optional Delimiter As String = ";")
Dim x As Long, y As Long, lDim As Long, lFile As Long
Dim s As String

  If (Not ArrayIsDimmed(Arr)) Then _
    Err.Raise 13, "libcore.ArrayToCSV", _
    StringMultiline("Array is not assigned or not an array type.", "Type: " & TypeName(Arr))
  
  lDim = ArrayDimensionCount(Arr)
  On Error GoTo ERRH
  Select Case lDim
    Case 1
      s = Join(Arr, Delimiter & vbNewLine)
    Case 2
      For x = LBound(Arr, 1) To UBound(Arr, 1)
        For y = LBound(Arr, 2) To UBound(Arr, 2)
          s = s & Replace(CStr(Arr(x, y)), vbNewLine, " ") & Delimiter
          If y = UBound(Arr, 2) Then s = s & vbNewLine
        Next y
      Next x
    Case Else
      Err.Raise 5, "libcore.ArrayToCSV", "Array with " & lDim & " dimensions is not supported."
  End Select
  
  On Error GoTo ERR_FILE
  lFile = FreeFile
  Filepath = PathChangeExtension(Filepath, ".csv")
  Open Filepath For Output As #lFile
  Print #lFile, s
  Close #lFile

Exit Sub
ERRH:
  Select Case Err.Number
    Case 438
      If lDim = 1 Then
        Err.Raise Err.Number, "libcore.ArrayToCSV", StringMultiline("Cannot write objects to a csv file.", _
        "Error: " & Err.Description)
      ElseIf lDim = 2 Then
        Err.Raise Err.Number, "libcore.ArrayToCSV", StringMultiline("Cannot write objects to a csv file.", _
        "Error: " & Err.Description, "Item: Array(" & x & "," & y & ")")
      End If
    Case Else
      Err.Raise Err.Number, "libcore.ArrayToCSV", StringMultiline("An error occured while processing the array.", _
      "Error: " & Err.Description)
  End Select
Exit Sub
ERR_FILE:
  Err.Raise Err.Number, "libcore.ArrayToCSV", StringMultiline("An error occured while creating the csv file.", _
  "Error: " & Err.Description)
End Sub

'@Description("Returns the owner of a file in domain\username format. Raises an error if the file is unavailable or doesn't exist.")
Function FileOwner(ByVal Filepath As String) As String
Dim secUtil As Object, secDescr As Object, fso As Object

  Set secUtil = NewClassReference("ADsSecurityUtility")
  Set fso = NewFileSystemObject

  If (Not fso.FileExists(Filepath)) Then _
    Err.Raise 53, "libcore.FileOwner", StringMultiline("File not found or unavailable.", "Filepath: " & Filepath)

  On Error Resume Next
  Set secDescr = secUtil.GetSecurityDescriptor(CVar(Filepath), 1, 1)
  FileOwner = secDescr.Owner
End Function

Function FileSelect(Optional sTitle As Variant, Optional sButtonName As Variant, Optional sDefaultPath As Variant, _
Optional bMultiSelect As Variant, Optional sExtensions As Variant) As String()
Dim fd As Office.FileDialog
Dim bHasDialog As Boolean
Dim xlApp As Object
Dim v() As String
Dim x As Long

  bHasDialog = (ApplicationIsExcel(Application) Or ApplicationIsAccess(Application))
  If bHasDialog Then
    Set xlApp = Application
  Else
    Set xlApp = NewExcelObject
  End If
  Set fd = xlApp.FileDialog(msoFileDialogFilePicker)
  
  If (Not IsMissing(bMultiSelect)) Then
    If VarType(bMultiSelect) <> vbBoolean Then _
      Err.Raise 13, "libcore.FileSelect", _
        StringMultiline("bMultiSelect argument must be type of Boolean.", "Type: " & TypeName(bMultiSelect))
    fd.AllowMultiSelect = bMultiSelect
  End If
  
  If (Not IsMissing(sTitle)) Then
    If VarType(sTitle) <> vbString Then _
      Err.Raise 13, "libcore.FileSelect", _
        StringMultiline("sTitle argument must be type of String.", "Type: " & TypeName(sTitle))
    fd.Title = sTitle
  End If
  
  If (Not IsMissing(sButtonName)) Then
    If VarType(sButtonName) <> vbString Then _
      Err.Raise 13, "libcore.FileSelect", _
        StringMultiline("sButtonName argument must be type of String.", "Type: " & TypeName(sButtonName))
    fd.ButtonName = sButtonName
  End If
  
  If (Not IsMissing(sDefaultPath)) Then
    If VarType(sDefaultPath) <> vbString Then _
      Err.Raise 13, "libcore.FileSelect", _
        StringMultiline("sDefaultPath argument must be type of String.", "Type: " & TypeName(sDefaultPath))
    fd.InitialFileName = sDefaultPath
  End If
  
  If (Not IsMissing(sExtensions)) Then
    If VarType(sExtensions) <> vbString Then _
      Err.Raise 13, "libcore.FileSelect", _
        StringMultiline("sExtensions argument must be type of String.", "Type: " & TypeName(sExtensions))
    fd.Filters.Clear
    fd.Filters.Add "Custom files only", sExtensions
  End If
  
  With fd
    .Show
    If .SelectedItems.Count = 0 Then Exit Function
    ReDim v(1 To .SelectedItems.Count)
    For x = 1 To .SelectedItems.Count
      v(x) = .SelectedItems(x)
    Next x
    FileSelect = v
  End With
  
  If (Not bHasDialog) Then xlApp.Quit
  Set fd = Nothing
End Function

'@Description("Returns the owner of a directory in domain\username format. Raises an error if the directory is unavailable or doesn't exist.")
Function DirectoryOwner(ByVal Dirpath As String) As String
Dim secUtil As Object, secDescr As Object, fso As Object

  Set secUtil = NewClassReference("ADsSecurityUtility")
  Set fso = NewFileSystemObject

  If (Not fso.FolderExists(Dirpath)) Then _
    Err.Raise 53, "libcore.DirectoryOwner", StringMultiline("Directory not found or unavailable.", "Path: " & Dirpath)

  On Error Resume Next
  Set secDescr = secUtil.GetSecurityDescriptor(CVar(Dirpath), 1, 1)
  DirectoryOwner = secDescr.Owner
End Function

'@Description("Creates all the folders and subfolders in the specified path. Raises an error if the path is an existing file.")
Function DirectoryCreate(ByVal Dirpath As String) As Boolean
Static fso As Object

  If (fso Is Nothing) Then Set fso = NewFileSystemObject
  With fso
    Select Case True
      Case .FileExists(Dirpath)
        Err.Raise 9, "libcore.DirectoryCreate", StringMultiline("Failed to create directory. DirPath argument is an existing file.", _
          "DirPath: " & Dirpath)
      Case .FolderExists(Dirpath)
        DirectoryCreate = True
        Exit Function
      Case DirectoryCreate(.GetParentFolderName(Dirpath))
        On Error Resume Next
        DirectoryCreate = (Not .CreateFolder(Dirpath) Is Nothing)
        On Error GoTo 0
    End Select
  End With
End Function

'@Description("Returns the folder path selected by the user. Returns vbNullString if no folder was selected.")
Function DirectorySelect(Optional sTitle As Variant, Optional sButtonName As Variant, Optional sDefaultPath As Variant) As String
Dim fd As Office.FileDialog
Dim bHasDialog As Boolean
Dim xlApp As Object

  bHasDialog = (ApplicationIsExcel(Application) Or ApplicationIsAccess(Application))
  If bHasDialog Then
    Set xlApp = Application
  Else
    Set xlApp = NewExcelObject
  End If
  Set fd = xlApp.FileDialog(msoFileDialogFolderPicker)
  
  If (Not IsMissing(sTitle)) Then
    If VarType(sTitle) <> vbString Then _
      Err.Raise 13, "libcore.DirectorySelect", _
        StringMultiline("sTitle argument must be type of String.", "Type: " & TypeName(sTitle))
    fd.Title = sTitle
  End If
  
  If (Not IsMissing(sButtonName)) Then
    If VarType(sButtonName) <> vbString Then _
      Err.Raise 13, "libcore.DirectorySelect", _
        StringMultiline("sButtonName argument must be type of String.", "Type: " & TypeName(sButtonName))
    fd.ButtonName = sButtonName
  End If
  
  If (Not IsMissing(sDefaultPath)) Then
    If VarType(sDefaultPath) <> vbString Then _
      Err.Raise 13, "libcore.DirectorySelect", _
        StringMultiline("sDefaultPath argument must be type of String.", "Type: " & TypeName(sDefaultPath))
    fd.InitialFileName = sDefaultPath
  End If
  
  With fd
    .Show
    If .SelectedItems.Count = 0 Then Exit Function
    DirectorySelect = .SelectedItems(1)
  End With
  
  If (Not bHasDialog) Then xlApp.Quit
  Set fd = Nothing
End Function

Function DoublesAreEqual(Double1 As Double, Double2 As Double, Optional EqualDigits As Integer = 8) As Boolean
  DoublesAreEqual = (Round(Double1, EqualDigits) = Round(Double2, EqualDigits))
End Function

'@Description("Returns a new object instance reference. Raises an error if class is not available.")
Function NewClassReference(ByVal ClassName As String) As Object
  On Error Resume Next
  Set NewClassReference = CreateObject(ClassName)
  On Error GoTo 0
  
  If (NewClassReference Is Nothing) Then _
    Err.Raise 429, "libcore.NewClassReference", _
    StringMultiline("Failed to create object instance. The class isn't registered, " & _
    "or DLL required by the object is unavailable.", "Class: " & ClassName)
End Function

'@Description("Returns a new Scripting.FileSystemObject instance reference")
Function NewFileSystemObject() As Object
  Set NewFileSystemObject = NewClassReference("Scripting.FileSystemObject")
End Function

'@Description("Returns a new Scripting.Dictionary instance reference")
Function NewDictionaryObject() As Object
  Set NewDictionaryObject = NewClassReference("Scripting.Dictionary")
End Function

'@Description("Returns a new VBScript.RegExp instance reference")
Function NewRegExpObject() As Object
  Set NewRegExpObject = NewClassReference("VBScript.RegExp")
End Function

'@Description("Returns the currently running Excel application object, or starts a new one.")
Function NewExcelObject() As Object
  On Error Resume Next
  Set NewExcelObject = GetObject(, "Excel.Application")
  If Err.Number <> 0 Then
    On Error GoTo 0
    Set NewExcelObject = NewClassReference("Excel.Application")
  End If
End Function

'@Description("Returns True if parameter object references are all set. Raises an error if parameter is not an object type or reference is not set.")
Function ObjectsAreAssigned(ParamArray Objects() As Variant) As Boolean
Dim x As Long

  For x = LBound(Objects) To UBound(Objects)
    Select Case True
      Case VarType(Objects(x)) <> VbVarType.vbObject
        Err.Raise 13, "libcore.ObjectsAreAssigned", StringMultiline("Type mismatch in parameter array.", _
        "Type: " & TypeName(Objects(x)), _
        "Index: " & x)
      Case (Objects(x) Is Nothing)
        Err.Raise 91, "libcore.ObjectsAreAssigned", StringMultiline("Object reference not set in parameter array.", _
        "Type: " & TypeName(Objects(x)), _
        "Index: " & x)
    End Select
  Next x
  ObjectsAreAssigned = True
End Function

'@Description("Returns UNC path for the specified network path. Returns vbNullString if network drive or path is unavailable on the machine.")
Function PathToUNC(ByVal NetworkPath As String) As String
Dim netwrk As Object, drves As Object
Dim s As String
Dim x As Long

  Set netwrk = NewClassReference("WScript.Network")
  Set drves = netwrk.EnumNetworkDrives
  s = Left$(NetworkPath, 2)
  For x = 0 To drves.Count - 1 Step 2
    If LCase$(drves.Item(x)) = LCase$(s) Then
      PathToUNC = drves.Item(x + 1) & Mid$(NetworkPath, 3)
      Exit For
    End If
  Next x
  Set netwrk = Nothing
  Set drves = Nothing
End Function

'@Description("Returns the extension (including the period '.') of the specified path string. Returns vbNullString if no extension was found.")
Function PathExtension(ByVal Path As String) As String
Dim lidx As Long

  lidx = InStrRev(Path, ".")
  If lidx > 0 Then _
    PathExtension = Right$(Path, Len(Path) - lidx + 1)
End Function

'@Description("Returns the specified path's parent directory.")
Function PathParentDirectory(ByVal Path As String) As String
  On Error Resume Next
  PathParentDirectory = Left$(Path, InStrRev(Path, "\") - 1)
End Function

'@Description("Returns a path where paramarray values are joined by backslash character(\).")
Function PathCombine(ParamArray Args() As Variant) As String
  PathCombine = Join(Args, "\")
End Function

'@Description("Returns the new filepath, with the new extension. Perion character is acceppted or can be omitted.")
Function PathChangeExtension(ByVal Path As String, ByVal NewExtension As String) As String
Dim lidx As Long

  If Left$(NewExtension, 1) <> "." Then NewExtension = "." & NewExtension
  lidx = InStrRev(Path, ".")
  Select Case lidx
    Case 0
      PathChangeExtension = Path & NewExtension
    Case Else
      PathChangeExtension = Left$(Path, lidx - 1) & NewExtension
  End Select
End Function

'@Description("Returns a new string where paramarray values are joined by the given delimiter.")
Function StringCombine(Delimiter As String, ParamArray Args() As Variant) As String
  On Error GoTo ERRH
  StringCombine = Join(Args, Delimiter)
Exit Function
ERRH:
  Select Case Err.Number
    Case 438
      Err.Raise Err.Number, "libcore.StringCombine", StringMultiline("An object was passed as an argument. Cannot convert object to string.", _
      "Error: " & Err.Description)
    Case Else
      Err.Raise Err.Number, "libcore.StringCombine", StringMultiline("An error occured while creating a combined string.", _
      "Error: " & Err.Description)
  End Select
End Function

'@Description("Returns whether a specified string is empty, or consists only of white-space characters.")
Function StringIsEmptyOrWhitespace(ByVal Text As Variant) As Boolean
  Select Case True
    Case IsError(Text)
    Case IsEmpty(Text), Text = vbNullString, Len(StringRemoveChars(Text, "\s")) = 0
      StringIsEmptyOrWhitespace = True
  End Select
End Function

'@Description("Replaces placeholders in a string with the parameter array values in order. Usage: StringInterpolate("Hello {0}! I love {1}.", "World", "VBA"))
Function StringInterpolate(ByVal Text As String, ParamArray Values() As Variant) As String
Dim x As Long

  For x = LBound(Values) To UBound(Values)
    Select Case VariantCanParseToString(Values(x))
      Case False
        Err.Raise 13, "libcore.StringInterpolate", StringMultiline("Type mismatch in parameter array.", _
        "Expected type(s): Numeric/String", _
        "Argument type: " & TypeName(Values(x)), _
        "Index: " & x)
      Case True
        Text = Replace(Text, "{" & x & "}", CStr(Values(x)))
    End Select
  Next x
  StringInterpolate = Text
End Function

'@Description("Returns a new string where each argument is concatenated by a new line.")
Function StringMultiline(ParamArray Lines() As Variant) As String
  On Error GoTo ERRH
  StringMultiline = Join(Lines, vbNewLine)
Exit Function
ERRH:
  Select Case Err.Number
    Case 438
      Err.Raise Err.Number, "libcore.StringMultiline", StringMultiline("An object was passed as an argument. Cannot convert object to string.", _
      "Error: " & Err.Description)
    Case Else
      Err.Raise Err.Number, "libcore.StringMultiline", StringMultiline("An error occured while creating a multiline string.", _
      "Error: " & Err.Description)
  End Select
End Function

'@Description("Writes a string to a file. Any data will be overwritten. File will be created if it doesn't exists.")
Sub StringToFileOverwrite(ByVal Text As String, Filepath As String)
Dim lFile As Long
  
  On Error GoTo ERR_FILE
  lFile = FreeFile
  Open Filepath For Output As #lFile
  Print #lFile, Text
  Close #lFile

Exit Sub
ERR_FILE:
  Err.Raise Err.Number, "libcore.StringToFileOverwrite", StringMultiline("An error occured while writing the file.", _
  "Error: " & Err.Description)
End Sub

'@Description("Appends a string to the end of the file. File will be created if it doesn't exists.")
Sub StringToFileAppend(ByVal Text As String, Filepath As String)
Dim lFile As Long
  
  On Error GoTo ERR_FILE
  lFile = FreeFile
  Open Filepath For Append As #lFile
  Print #lFile, Text
  Close #lFile

Exit Sub
ERR_FILE:
  Err.Raise Err.Number, "libcore.StringToFileAppend", StringMultiline("An error occured while writing the file.", _
  "Error: " & Err.Description)
End Sub

'@Description("Returns SHA256Managed hash for the specified string.")
Function StringEncryptSHA256(ByVal Text As String) As String
Dim enc As Object, prov As Object
Dim hash() As Byte
Dim x As Integer

  Set enc = NewClassReference("System.Text.UTF8Encoding")
  Set prov = NewClassReference("System.Security.Cryptography.SHA256Managed")
  hash = prov.ComputeHash_2(enc.Getbytes_4(Text))
  For x = LBound(hash) To UBound(hash)
    StringEncryptSHA256 = StringEncryptSHA256 & Hex(hash(x) / 16) & Hex(hash(x) Mod 16)
  Next x
End Function

'@Description("Returns a new string, where the specified characters have been removed.")
Function StringRemoveChars(ByVal Text As String, Optional sRegexpPattern As Variant, Optional RegExpObj As Object) As String
Dim rgx As Object

  If (Not IsMissing(sRegexpPattern)) Then
    If VarType(sRegexpPattern) <> vbString Then _
      Err.Raise 13, "libcore.StringRemoveChars", _
        StringMultiline("sRegexpPattern argument must be type of String.", "Type: " & TypeName$(sRegexpPattern))
  Else
    sRegexpPattern = "[^a-z0-9ˆı¸˚˙Û·È]" ' Welcome from Hungary :)
  End If
  
  Set rgx = RegExpObj
  If (rgx Is Nothing) Then Set rgx = NewRegExpObject
  
  If LCase$(TypeName$(rgx)) <> "iregexp2" Then _
    Err.Raise 13, "libcore.StringRemoveChars", _
    StringMultiline("RegExpObj argument must be type of VBScript.RegExp.", "Type: " & TypeName$(rgx))
  
  If (RegExpObj Is Nothing) Then
    With rgx
      .Global = True
      .IgnoreCase = True
      .MultiLine = True
      .Pattern = sRegexpPattern
    End With
  End If
  
  StringRemoveChars = rgx.Replace(Text, vbNullString)
End Function

'@Description("Returns True if data type can be parsed as a string.")
Function VariantCanParseToString(Var As Variant) As Boolean
  Select Case VarType(Var)
    Case VbVarType.vbError, VbVarType.vbObject, VbVarType.vbUserDefinedType, VbVarType.vbArray
    Case Else
      VariantCanParseToString = True
  End Select
End Function

'@Description("Returns a Windows account username. If parameters are omitted, aims for the current user's username. Returns vbNullString if fails.")
Function WindowsUserName(Optional ByVal sDomain As Variant, Optional ByVal sUserName As Variant) As String
Dim o As Object

  If IsMissing(sDomain) Then
    sDomain = Environ("userdomain")
  Else
    If VarType(sDomain) <> vbString Then _
      Err.Raise 13, "libcore.WindowsUserName", _
        StringMultiline("sDomain argument must be type of String.", "Type: " & TypeName(sDomain))
  End If
  
  If IsMissing(sUserName) Then
    sUserName = Environ("username")
  Else
    If VarType(sUserName) <> vbString Then _
      Err.Raise 13, "libcore.WindowsUserName", _
        StringMultiline("sUserName argument must be type of String.", "Type: " & TypeName(sUserName))
  End If
  
  On Error Resume Next
  Set o = GetObject("WinNT://" & sDomain & "/" & sUserName & ",user")
  WindowsUserName = o.FullName
End Function
