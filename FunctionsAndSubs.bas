Attribute VB_Name = "FunctionsAndSubs"
Public rootDir As String
Public FilesAndFolders As Collection
Public f As FileData 'File or Folder data
Public varDirectory As String
Public names() As Variant
Public addresses() As Variant
Public fileCount As Long

Sub rangeToTable(ByVal startrow As Long, ByVal startcol As Long, ByVal tablename As String)

ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(startrow, startcol), Cells(lastRow(startcol), lastColumn(startrow))), , xlNo).Name = tablename

End Sub

Sub IterateThroughFolder(folder As String)
Dim i As Integer

varDirectory = Dir(folder, vbDirectory)

While varDirectory <> ""

    If InStrRev(folder + varDirectory, ".") <> Len(folder + varDirectory) Then
    
        Set f = New FileData

        With f
            f.Name = varDirectory
            f.Address = folder + varDirectory
        End With

        FilesAndFolders.Add f
        fileCount = fileCount + 1
    End If
    
    varDirectory = Dir

Wend

End Sub

Sub Auto_Open()

Call Sheet1.main

End Sub


Function isFile(s As String)
'file is true
'folder is false
On Error GoTo Errhandler:
If InStr(s, ".") = 0 And (GetAttr(s) = 16 Or GetAttr(s) = 17) Then
    isFile = False
Else: isFile = True

Exit Sub

Errhandler:
isFile = False

End If

End Function


Sub pasteArray(ByVal row As Long, ByVal col As Long, a() As Variant)

Range(Cells(row, col), Cells(row + UBound(a) - 1, col)).Value = Application.Transpose(a)

End Sub

Function lastRow(ByVal c As Long) As Long

lastRow = Cells(Rows.Count, c).End(xlUp).row

End Function

Function lastColumn(ByVal r As Long) As Long

lastColumn = Cells(r, Columns.Count).End(xlToLeft).Column

End Function

Sub clearRange(ByVal sr As Long, ByVal er As Long, ByVal sc As Long, ByVal ec As Long)

Range(Cells(sr, sc), Cells(er, ec)).Clear

End Sub

