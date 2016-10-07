Attribute VB_Name = "Export"
Sub ExportAsPOmono()

Dim PO_Contents, msgidValue, msgstrValue As String
Dim RowNo, LastRow, SourceColumnNo As Integer
Dim WriteContext As Boolean
Dim SourceColumn As Range
Dim PO_FilePath, FileName As String
Dim UTFStream, BinaryStream As Object
Dim fso


'Error handling:
'     https://msdn.microsoft.com/en-us/library/office/ff839468.aspx?f=255&MSPPError=-2147217396
'     http://stackoverflow.com/questions/19609479/handle-cancellation-of-inputbox-to-select-range
On Error Resume Next

'Check for Context to write as msgid
If MsgBox("Does Column A contain Context or IDs to be written as msgid values?" & vbNewLine & vbNewLine _
        & "If no, row numbers will be used for msgid." & vbNewLine, _
        vbYesNo, "Checking for Context") = vbYes Then
    WriteContext = True
Else
    WriteContext = False
End If

'Get the column with source text
Set SourceColumn = Application.InputBox("Which column contains source texts to be written as msgstr values?" & _
                vbNewLine & vbNewLine & "Select any cell in this column." & vbNewLine & vbNewLine, "Select Source", "B1", Type:=8)
On Error GoTo SelectionCanceled
SourceColumnNo = SourceColumn.column
PO_Contents = ""

'Find the column limit
LastRow = ActiveSheet.Columns(SourceColumnNo).Find(What:="*", _
    SearchOrder:=xlByRows, _
    SearchDirection:=xlPrevious).Row

'Run through data in range
For RowNo = 1 To LastRow
    'Get msgid based on Context check
    If WriteContext = False Then
        msgidValue = RowNo
    Else
        msgidValue = ActiveSheet.Cells(RowNo, 1).Value
    End If
    'Get source text from Column B (2)
    msgstrValue = ActiveSheet.Cells(RowNo, SourceColumnNo).Value
    'Escape quotes to form valid PO values
    msgstrValue = Replace(msgstrValue, """", "\""")
    'Build PO contents
    PO_Contents = PO_Contents & "msgid """ & msgidValue & """" & vbCrLf & "msgstr """ & msgstrValue & """" & vbCrLf & vbCrLf
Next

'Set PO file name and path
Set fso = CreateObject("Scripting.FileSystemObject")
FileName = fso.GetBaseName(ActiveWorkbook.Name)
PO_FilePath = Application.GetSaveAsFilename(FileName, "(*.po),*.po", , "Save As...")
If PO_FilePath = vbNullString Then
    Exit Sub
End If

'Write PO_Contents into file
Set UTFStream = CreateObject("adodb.stream")
UTFStream.Type = 2 'Put the text into the buffer as adTypeText
UTFStream.Mode = adModeReadWrite
UTFStream.Charset = "UTF-8"
UTFStream.Open
UTFStream.WriteText PO_Contents
UTFStream.Position = 3 'Skip BOM

Set BinaryStream = CreateObject("adodb.stream")
BinaryStream.Type = 1 'Set the stream type as adTypeBinary
BinaryStream.Mode = adModeReadWrite
BinaryStream.Open

UTFStream.CopyTo BinaryStream 'Strip BOM (first 3 bytes)
UTFStream.flush
UTFStream.Close

BinaryStream.SaveToFile PO_FilePath, 2
BinaryStream.flush
BinaryStream.Close

SelectionCanceled:
Set SourceColumn = Nothing

End Sub
Sub ExportAsTMX()

Dim TMX_Contents, Locale, TextValue As String
Dim RowNo, ColumnNo, LastRow, LastColumn As Integer
Dim TMX_FilePath As String
'Dim intFileNum As Integer
Dim UTFStream, BinaryStream As Object


'Info and confirmation
If MsgBox("This worksheet must have language columns only, with source in A column." & vbNewLine _
        & "Locale names (en, de, ...) must be in row 1." & vbNewLine _
        & "IDs are not supported." & vbNewLine & vbNewLine _
        & "Continue?", _
        vbYesNo, "Export As TMX") = vbNo Then
    Exit Sub
End If

'Set TMX file path
TMX_FilePath = Application.GetSaveAsFilename(, "(*.tmx),*.tmx", , "Save As...")
If TMX_FilePath = vbNullString Then
    Exit Sub
End If

'Find data range limits

LastRow = ActiveSheet.Cells.Find(What:="*", After:=[A1], _
    SearchOrder:=xlByRows, _
    SearchDirection:=xlPrevious).Row

LastColumn = ActiveSheet.Cells.Find(What:="*", After:=[A1], _
    SearchOrder:=xlByColumns, _
    SearchDirection:=xlPrevious).column

'Run through data in range, wrap into TMX markup and concatenate
TMX_Contents = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf & _
    "<tmx version=""1.4"">" & vbCrLf & _
    "  <header datatype=""PlainText"" srclang=""" & Range("A1").Value & """/>" & vbCrLf & _
    "  <body>" & vbCrLf
Locale = ""
TextValue = ""
For RowNo = 2 To LastRow
    TMX_Contents = TMX_Contents & "    <tu>" & vbCrLf
    For ColumnNo = 1 To LastColumn
        Locale = ActiveSheet.Cells(1, ColumnNo).Value
        TextValue = ActiveSheet.Cells(RowNo, ColumnNo).Value
        TMX_Contents = TMX_Contents & _
            "      <tuv xml:lang=""" & Locale & """>" & vbCrLf & _
            "        <seg>" & TextValue & "</seg>" & vbCrLf & _
            "      </tuv>" & vbCrLf
    Next
    TMX_Contents = TMX_Contents & _
        "    </tu>" & vbCrLf
Next
TMX_Contents = TMX_Contents & _
    "  </body>" & vbCrLf & "</TMX>" & vbCrLf

'Write TMX_Contents into file

' This is the solution to write utf-8 encoded text file, also without BOM.
' http://stackoverflow.com/questions/2524703/save-text-file-utf-8-encoded-with-vba
' http://stackoverflow.com/questions/4143524/can-i-export-excel-data-with-utf-8-without-bom

Set UTFStream = CreateObject("adodb.stream")
UTFStream.Type = 2 'Put the text into the buffer as adTypeText
UTFStream.Mode = adModeReadWrite
UTFStream.Charset = "UTF-8"
UTFStream.Open
UTFStream.WriteText TMX_Contents
UTFStream.Position = 3 'Skip BOM

Set BinaryStream = CreateObject("adodb.stream")
BinaryStream.Type = 1 'Set the stream type as adTypeBinary
BinaryStream.Mode = adModeReadWrite
BinaryStream.Open

UTFStream.CopyTo BinaryStream 'Strip BOM (first 3 bytes)
UTFStream.flush
UTFStream.Close

BinaryStream.SaveToFile TMX_FilePath, 2
BinaryStream.flush
BinaryStream.Close

'' Doesn't work for non-ASCII characters, so this solition is discarded
'' Get next file number
'intFileNum = FreeFile
'' Open the file, write output, then close file
'Open TMX_FilePath For Output As #intFileNum
'Print #intFileNum, TMX_Contents
'Close #intFileNum

End Sub

