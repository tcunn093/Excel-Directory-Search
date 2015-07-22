VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Const containsFieldRange As String = "C2"
Const doesNotContainFieldRange As String = "C3"

Const xofyRange As String = "C1"
Const titleRange As String = "B1"
Const fileCountRange As String = "A3"
Const fileCountCol As Long = 1
Const baseRow As Long = 4
Const baseCol As Long = 2
Const namesRow As Long = baseRow
Const namesCol As Long = baseCol
Const addressesRow = baseRow
Const addressesCol = baseCol + 1

Private Sub SpreadsheetBeautification()

ActiveWindow.DisplayGridlines = False

With Cells.Font

    .Color = RGB(255, 255, 255)
    .Size = 8
    .Name = "Lucida Console"

End With

Cells(namesRow, namesCol).Interior.Color = RGB(0, 0, 0)

Range(xofyRange).Value = ""
Cells(namesRow, namesCol).Value = "File Names"
Cells(addressesRow, addressesCol).Value = "File Location"

Range(containsFieldRange).Offset(0, -1).Value = "Search Term"
Range(doesNotContainFieldRange).Offset(0, -1).Value = "Term to Ignore"


With Rows(4).EntireRow

With .Interior
.Color = RGB(100, 100, 100)
End With

.HorizontalAlignment = xlCenter

End With

With Range(containsFieldRange & ":" & doesNotContainFieldRange).EntireRow
.Interior.Color = RGB(153, 153, 153)
.Borders.Color = RGB(100, 100, 100)

End With

With Range(fileCountRange)

With .Offset(-1, 0)
.Value = "Total Number of Files"
.Interior.Color = RGB(153, 153, 153)
.Borders.Color = RGB(100, 100, 100)
.HorizontalAlignment = xlCenter
End With

With .Offset(1, 0)
.Interior.Color = RGB(100, 100, 100)
.Borders.LineStyle = xlNone
.HorizontalAlignment = xlCenter
End With

.Interior.Color = RGB(153, 153, 153)
.Borders.Color = RGB(100, 100, 100)
.HorizontalAlignment = xlCenter
.EntireColumn.ColumnWidth = 25

End With

With Range(containsFieldRange & ":" & doesNotContainFieldRange)
    
.Interior.Color = RGB(255, 255, 255)
.Borders.Color = RGB(100, 100, 100)
.Font.Color = RGB(0, 0, 0)
.Offset(0, -1).HorizontalAlignment = xlRight

End With

With Range(titleRange)

.Value = "Directory Search 1.0"
.Font.Size = 16
.HorizontalAlignment = xlCenter
.EntireRow.Interior.Color = RGB(100, 100, 100)

End With

Columns(namesCol).ColumnWidth = 45
Columns(addressesCol).ColumnWidth = 90

For i = 1 To 2
Range("B4:C4").AutoFilter Field:=i, Visibledropdown:=False
Next i

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

Dim cef As String: cf = Range(containsFieldRange)
Dim dncf As String: dncf = Range(doesNotContainFieldRange)
Dim tableRange As Range


If Not Intersect(Target, Range(containsFieldRange & ":" & doesNotContainFieldRange)) Is Nothing Then
Set tableRange = ListObjects("FilesAndFolders").Range
    If cf <> "" And dncf <> "" Then

        tableRange.AutoFilter Field:=2, _
        Criteria1:="=*" & cf & "*", Operator:=xlAnd, _
        Criteria2:="<>*" & dncf & "*"

    ElseIf cf <> "" And dncf = "" Then
        tableRange.AutoFilter Field:=2, _
        Criteria1:="=*" & cf & "*"

    ElseIf cf = "" And dncf <> "" Then
        tableRange.AutoFilter Field:=2, _
        Criteria1:="<>*" & dncf & "*"

    Else:

        tableRange.AutoFilter Field:=2

    End If

Range(xofyRange).Value = tableRange.SpecialCells(xlCellTypeVisible).Count / 2 - 1 _
& " of " & tableRange.Rows.Count - 1 & " Records"

End If

End Sub


Sub main()

Application.ScreenUpdating = False
Dim counter As Long: counter = 1

rootDir = Application.ActiveWorkbook.Path + "\"

If ListObjects.Count = 1 Then
    ListObjects("FilesAndFolders").Delete
End If

Cells.Interior.Color = RGB(0, 0, 0)

Set FilesAndFolders = New Collection

Call IterateThroughFolder(rootDir)

While counter < FilesAndFolders.Count

    Call IterateThroughFolder(FilesAndFolders.Item(counter).Address)
    DoEvents
    Application.StatusBar = "Found " & FilesAndFolders.Count & " files and folders"
    counter = counter + 1

Wend

ReDim names(FilesAndFolders.Count)
ReDim addresses(FilesAndFolders.Count)

For counter = 1 To FilesAndFolders.Count
    
    names(counter - 1) = FilesAndFolders.Item(counter).Name
    addresses(counter - 1) = FilesAndFolders.Item(counter).Address
    
Next counter

Call pasteArray(namesRow, namesCol, names)
Call pasteArray(addressesRow, addressesCol, addresses)

Call rangeToTable(baseRow, baseCol, "FilesAndFolders")

Range(fileCountRange).Value = ListObjects("FilesAndFolders").Range.Rows.Count

Call SpreadsheetBeautification

ListObjects("FilesAndFolders").ListColumns(1).Name = "File Names"
ListObjects("FilesAndFolders").ListColumns(2).Name = "File Addresses"
ListObjects("FilesAndFolders").TableStyle = ""

Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

Set KeyCells = Range(Cells(namesRow + 1, namesCol), Cells(lastRow(addressesCol), lastColumn(addressesRow)))
If Intersect(Target, KeyCells) Is Nothing Then
Else
Shell "C:\Windows\explorer.exe /select, " & Cells(Selection.row, addressesCol).Value, vbNormalFocus
End If

End Sub
