Attribute VB_Name = "Export"
Sub ExportAsPO_Select()

' This macro exports a single column as a monolingual PO file with row number as msgid and cell value as msgstr.
' msgid "row_number_as_ad_hoc_ID"
' msgstr "text"
' The column for export can be selected arbitrarily.

Dim PO_Contents, TextValue As String
Dim RowNo, LastRow, SourceColumnNo As Integer
Dim SourceColumn As Range
Dim PO_FilePath, FileName As String
Dim UTFStream, BinaryStream As Object
Dim fso

' Get the column with source text
' Error handling:
' https://msdn.microsoft.com/en-us/library/office/ff839468.aspx?f=255&MSPPError=-2147217396
' http://stackoverflow.com/questions/19609479/handle-cancellation-of-inputbox-to-select-range
On Error Resume Next
Set SourceColumn = Application.InputBox("Select any cell in the column with source text." & _
                vbNewLine & vbNewLine & "Text values in this column will be written as msgstr, and row number as msgid." & vbNewLine, "Select Source", "A1", Type:=8)
On Error GoTo SelectionCanceled

SourceColumnNo = SourceColumn.column

' Find the column limit
LastRow = ActiveSheet.Columns(SourceColumnNo).Find(What:="*", _
    SearchOrder:=xlByRows, _
    SearchDirection:=xlPrevious).Row

PO_Contents = ""
TextValue = ""
' Run through data in range
For RowNo = 1 To LastRow
    ' Read source text from Column B (2)
    TextValue = ActiveSheet.Cells(RowNo, SourceColumnNo).Value
    ' Escape quotes to form valid PO values
    TextValue = Replace(TextValue, """", "\""")
    ' Build PO contents
    PO_Contents = PO_Contents & "msgid """ & RowNo & """" & vbCrLf & "msgstr """ & TextValue & """" & vbCrLf & vbCrLf
Next

'Set PO file name and path
Set fso = CreateObject("Scripting.FileSystemObject")
FileName = fso.GetBaseName(ActiveWorkbook.Name)
PO_FilePath = Application.GetSaveAsFilename(FileName, "(*.po),*.po", , "Save As...")
If PO_FilePath = vbNullString Then
    Exit Sub
End If

' Write PO_Contents into file
Set UTFStream = CreateObject("adodb.stream")
UTFStream.Type = 2 ' Put the text into the buffer as adTypeText
UTFStream.Mode = adModeReadWrite
UTFStream.Charset = "UTF-8"
UTFStream.Open
UTFStream.WriteText PO_Contents
UTFStream.Position = 3 'skip BOM

Set BinaryStream = CreateObject("adodb.stream")
BinaryStream.Type = 1 ' set the stream type as adTypeBinary
BinaryStream.Mode = adModeReadWrite
BinaryStream.Open

UTFStream.CopyTo BinaryStream 'Strips BOM (first 3 bytes)
UTFStream.flush
UTFStream.Close

BinaryStream.SaveToFile PO_FilePath, 2
BinaryStream.flush
BinaryStream.Close

SelectionCanceled:
Set SourceColumn = Nothing

End Sub