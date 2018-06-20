<html>
<head>
</head>
<b>Excel File Comparison</b>
<Body>
<br>
<Form action='' method='post'>
Excel File 1: 
<input name='File1'>
Excel File 2: 
<input name='File2'>
<input type='submit' value='Compare'>
</Form>
<%
const adopenforwardonly = 0
const adopenstatic = 3
const adlockreadonly = 1
const adlockpessimistic = 2
const adcmdtext = &H0001
const adcmdtable = &H0002


function cellComparison(sheetName, file1, file2)
    Dim differences
    Dim CS1, RS1, SQ, CS2, RS2
    differences = ""
    CS1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(file1) & ";Persist Security Info=False;Extended Properties=""Excel 8.0;IMEX=1"""
    CS2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(file2) & ";Persist Security Info=False;Extended Properties=""Excel 8.0;IMEX=1"""
    SQ = "SELECT * FROM [" & sheetName & "]"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    Set RS2 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQ, CS1, adopenforwardonly, adlockreadonly, adcmdtext
    RS2.Open SQ, CS2, adopenforwardonly, adlockreadonly, adcmdtext
    dim lineNum, differences1, differences2
    lineNum=1
    Do While Not RS1.EOF 
        Dim fieldNum
        fieldNum=0 'column number'
        lineNum=lineNum+1
        For Each F in RS1.Fields 
            fieldNum=fieldNum+1
            differences1 = differences1 & RS1(F.Name) & "*" 
            Next
        RS1.MoveNext
        Loop

    'lineNum=1
    'Do While Not RS2.EOF
        'fieldNum=0
        'lineNum=lineNum+1
        'For Each F in RS2.Fields 
            'fieldNum=fieldNum+1
            'difference2 = differences2 & "(" & lineNum & "," & fieldNum & "): " & RS2(F.Name) & "*" 
            'Next
        'RS2.MoveNext
        'Loop

    'Dim cell1, cell2
    'cell1 = Split(differences1, "*")
    'cell2 = Split(differences2, "*")
        
    cellComparison = differences1
    End Function

function getValues(sheetName, file1, file2)
    Dim differences
    Dim CS1, RS1, SQ, CS2, RS2
    differences = ""
    CS1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(file1) & ";Persist Security Info=False;Extended Properties=""Excel 8.0;IMEX=1"""
    CS2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(file2) & ";Persist Security Info=False;Extended Properties=""Excel 8.0;IMEX=1"""
    SQ = "SELECT * FROM [" & sheetName & "]"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    Set RS2 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQ, CS1, adopenforwardonly, adlockreadonly, adcmdtext
    RS2.Open SQ, CS2, adopenforwardonly, adlockreadonly, adcmdtext
    dim lineNum, differences1, differences2
    lineNum=1
    Do While Not RS1.EOF 
        Dim fieldNum
        fieldNum=0 'column number'
        lineNum=lineNum+1
        For Each F in RS1.Fields 
            fieldNum=fieldNum+1
            differences1 = differences1 & lineNum & "*" & fieldNum & "*" 
            Next
        RS1.MoveNext
        Loop


    getValues = differences1
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
            Dim diff, parts, index
            diff = cellComparison(sheet, Request.Form("File1"), Request.Form("File2"))
            index = getValues(sheet, Request.Form("File1"), Request.Form("File2"))
            Pr "<tr>"
            Pr "<td>" & sheet & "</td>"
            Pr "<td>" & sheet & "</td>"
            If Len(diff) > 0 Then
                Pr "<td><Form action='sheetCompare.asp' method='post'>"
                Pr "<input type='hidden' name='diff' value='"&diff&"'>"
                Pr "<input type='hidden' name='index' value='"&index&"'>"
                Pr "<input type='hidden' name='sheet' value='"&sheet&"'>"
                Pr "<input type='hidden' name='file1' value='"&Request.Form("file1")&"'>"
                Pr "<input type='hidden' name='file2' value='"&Request.Form("file2")&"'>"
                Pr "<input type='submit' value='View Differences'>"
                Pr "</Form></td></tr>"
            Else
                Pr "<td></td></tr>"
                End If
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
            dim x
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
