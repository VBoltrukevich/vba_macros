Attribute VB_Name = "ExportAsTMX"
Sub ExportAsTMX()

' This macro exports a TMX file from a single Excel worksheet with this structure:
' Locale names (en, de, ...) are taken from row 1.
' Source text entries are taken from column A.
' Translations in each target language go in its respective column.
' IDs are not supported, so delete the column for Context in Excel files exported from Sisulizer.

Dim TMX_Contents, Locale, TextValue As String
Dim RowNo, ColumnNo, LastRow, LastColumn As Integer
Dim TMX_FilePath As String
'Dim intFileNum As Integer
Dim UTFStream, BinaryStream As Object


' Info and confirmation
If MsgBox("This worksheet must have language columns only, with source in column A." & vbNewLine _
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

' Find data range limits

LastRow = ActiveSheet.Cells.Find(What:="*", After:=[A1], _
    SearchOrder:=xlByRows, _
    SearchDirection:=xlPrevious).Row

LastColumn = ActiveSheet.Cells.Find(What:="*", After:=[A1], _
    SearchOrder:=xlByColumns, _
    SearchDirection:=xlPrevious).column

' Run through data in range, wrap into TMX markup and concatenate
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

' Write TMX_Contents into file

' This is the solution to write utf-8 encoded text file, also without BOM.
' http://stackoverflow.com/questions/2524703/save-text-file-utf-8-encoded-with-vba
' http://stackoverflow.com/questions/4143524/can-i-export-excel-data-with-utf-8-without-bom

Set UTFStream = CreateObject("adodb.stream")
UTFStream.Type = 2 ' Put the text into the buffer as adTypeText
UTFStream.Mode = adModeReadWrite
UTFStream.Charset = "UTF-8"
UTFStream.Open
UTFStream.WriteText TMX_Contents
UTFStream.Position = 3 'skip BOM

Set BinaryStream = CreateObject("adodb.stream")
BinaryStream.Type = 1 ' set the stream type as adTypeBinary
BinaryStream.Mode = adModeReadWrite
BinaryStream.Open

UTFStream.CopyTo BinaryStream 'Strips BOM (first 3 bytes)
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
