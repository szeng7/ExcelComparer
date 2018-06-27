<html>
<head>
</head>
<b>Excel File Comparison</b>
<Body>
<br>
<Form action='' method='post'>
Excel File 1: 
<input type='file' name='File1'>
Excel File 2: 
<input type='file' name='File2'>
<input type='submit' value='Compare'>
</Form>
<%
const adopenforwardonly = 0
const adopenstatic = 3
const adlockreadonly = 1
const adlockpessimistic = 2
const adcmdtext = &H0001
const adcmdtable = &H0002

function IsBlank(Value)
'returns True if Empty or NULL or Zero
If IsEmpty(Value) or IsNull(Value) Then
 IsBlank = True
ElseIf VarType(Value) = vbString Then
 If Value = "" Then
  IsBlank = True
 End If
ElseIf IsObject(Value) Then
 If Value Is Nothing Then
  IsBlank = True
 End If
ElseIf IsNumeric(Value) Then
 If Value = 0 Then
  wscript.echo " Zero value found"
  IsBlank = True
 End If
Else
 IsBlank = False
End If
End Function

'get number of columns/fields in file'
function numFields(sheetName, file)
    Dim CS1, RS1, SQ
    CS1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(file) & ";Persist Security Info=False;Extended Properties=""Excel 8.0;IMEX=1"""
    SQ = "SELECT * FROM [" & sheetName & "]"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQ, CS1, adopenforwardonly, adlockreadonly, adcmdtext
    numFields = RS1.Fields.Count
    End Function

'get number of rows in file'
function numRows(sheetName, file)
    Dim CS1, RS1, SQ, rows
    rows = 0
    CS1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(file) & ";Persist Security Info=False;Extended Properties=""Excel 8.0;IMEX=1"""
    SQ = "SELECT * FROM [" & sheetName & "]"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQ, CS1, adopenforwardonly, adlockreadonly, adcmdtext 
    Do While Not RS1.EOF
        rows = rows + 1
        RS1.MoveNext
        Loop
    numRows = rows
    End Function

'get values of all cells in a file as a string delimited by *'
function getValues(sheetName, file, maxRows, maxFields)
    Dim differences, value
    Dim CS1, RS1, SQ, columns, rows
    differences = ""
    CS1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(file) & ";Persist Security Info=False;Extended Properties=""Excel 8.0;IMEX=1"""
    SQ = "SELECT * FROM [" & sheetName & "]"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQ, CS1, adopenforwardonly, adlockreadonly, adcmdtext
    rows = 0
    Do While Not RS1.EOF
        rows = rows + 1
        columns = 0
        For Each F in RS1.Fields
            columns = columns + 1
            If IsBlank(F) = True Then
                value = "{Empty}"
            Else
                value = RS1(F.Name)
            End If
            differences = differences & value & "*" 
            Next
        Do While columns < maxFields 
            differences = differences & "{Empty}*"
            columns = columns + 1
            Loop
        RS1.MoveNext
        Loop
    Do While rows < maxRows
        For I = 0 to maxFields
            differences = differences & "{Empty}*"
        Next
        rows = rows + 1
    Loop
    getValues = differences
    End Function

Sub Pr(S)
    Response.Write S
    End Sub


Dim File1Sheets, File2Sheets
File1Sheets = ""
File2Sheets = ""
If Request.Form <> "" Then 

    Dim oConn1,sConn1,oConn2,sConn2
    Set oConn1 = Server.CreateObject("ADODB.Connection")
    Set oConn2 = Server.CreateObject("ADODB.Connection")
    sConn1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(Request.Form("File1")) & ";Persist Security Info=False; Extended Properties=""Excel 8.0;IMEX=1"""
    sConn2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(Request.Form("File2")) & ";Persist Security Info=False; Extended Properties=""Excel 8.0;IMEX=1"""
    oConn1.Open sConn1
    oConn2.Open sConn2
    Dim oRS1, oRS2
    Set oRS1 = oConn1.OpenSchema(20)
    Set oRS2 = oConn2.OpenSchema(20)
    Do While Not oRS1.EOF
        sSheetName = oRS1.Fields("table_name").Value
        File1Sheets = File1Sheets & sSheetName & ":"
        oRS1.MoveNext()
    Loop
    File1Sheets = StrReverse(File1Sheets)
    File1Sheets = StrReverse(Replace(File1Sheets,":","",1,1))
    Do While Not oRS2.EOF
        sSheetName = oRS2.Fields("table_name").Value
        File2Sheets = File2Sheets & sSheetName & ":"
        oRS2.MoveNext()
    Loop
    File2Sheets = StrReverse(File2Sheets)
    File2Sheets = StrReverse(Replace(File2Sheets,":","",1,1))
    Pr "<center><table border='1' cellspacing='0'>"
    Pr "<tr><td><b>" & Request.Form("File1") & " Sheets</b></td>"
    Pr "<td><b>" & Request.Form("File2") & " Sheets</b></td>"
    Pr "<td><b>Differences</b></td></tr>"

    For Each sheet in Split(File1Sheets,":")
        If Instr(File2Sheets, sheet) Then
            Dim values1, values2, valuesplit, valuesplit2
            Dim filename1, filename2, maxRows, maxFields, maxAttempt
            maxRows = numRows(sheet,Request.Form("File1")) 'number of rows in file1'
            maxAttempt = numRows(sheet, Request.Form("File2")) 'number of rows in file2'
            If maxAttempt > maxRows Then
                maxRows = maxAttempt
                End If
            maxFields = numFields(sheet, Request.Form("File1"))
            maxAttempt = numFields(sheet, Request.Form("File2"))
            If maxAttempt > maxFields Then
                maxFields = maxAttempt
                End If
            values1 = getValues(sheet, Request.Form("File1"), maxRows, maxFields)
            values2 = getValues(sheet, Request.Form("File2"), maxRows, maxFields)
            valuesplit = Split(values1, "*")
            valuesplit2 = Split(values2, "*")
            filename1 = Request.Form("file1")
            filename2 = Request.Form("file2")
            Dim I, J, finaldiff, cellValue, cellValue2
            finaldiff = ""
            For I=0 to maxRows-1
                For J=0 to maxFields - 1
                    cellValue = valuesplit(maxFields*I + J)
                    cellValue2 = valuesplit2(maxFields*I + J)
                    If StrComp(cellValue, cellValue2) <> 0 Then
                        finaldiff = finaldiff & "(Row " & I+2 & ", Column " & J+1 & ")\ " & cellValue & " vs " & cellValue2 & "\"
                    End If
                    Next
                Next

            Pr "<tr>"
            Pr "<td>" & sheet & "</td>"
            Pr "<td>" & sheet & "</td>"
            If Len(finaldiff) > 0 Then
                Pr "<td><Form action='sheetCompare.asp' method='post'>"
                Pr "<input type='hidden' name='finaldiff' value='"&finaldiff&"'>"
                Pr "<input type='hidden' name='sheet' value='"&sheet&"'>"
                Pr "<input type='hidden' name='file1' value='"&filename1&"'>"
                Pr "<input type='hidden' name='file2' value='"&filename2&"'>"
                Pr "<input type='submit' value='View Differences'>"
                Pr "</Form></td></tr>"
            Else
                Pr "<td></td></tr>"
                End If
            finaldiff=""
        Else 
            Pr "<tr>"
            Pr "<td>" & sheet & "</td>"
            Pr "<td></td>"
            Pr "<td></td>"
            Pr "</tr>"
            End If
        Next
    For Each sheet in Split(File2Sheets,":")
        If Instr(File1Sheets, sheet) Then
            dim p
        Else
            Pr "<tr>"
            Pr "<td></td>"
            Pr "<td>" & sheet & "</td>"
            Pr "<td></td>"
            Pr "</tr>"
            End If
        Next
    Pr "</center>"
    End If
%>
