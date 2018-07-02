<html>
<Body>

<style>

table {
  overflow: hidden;
}

body,td {
	font-family: tahoma,verdana,arial,helvetica;
}
td { 
	color: black; 
	text-decoration: none; 
	background-color: #eeeeee;
}

tr:hover {
  text-shadow: 0.5px 0.5px grey;
}
</style>

<%
Sub Pr(S)
    Response.Write S
    End Sub
Dim I, differencesplit, lastRow, file1contents, file2contents
Pr Request.Form("sheet") & " Comparison<br>"
Pr "<center><table border='1' cellspacing='1' style='background-color: black' ><tr><td style='background-color: #FFFFCC' ><b>Column</b></td>"&"<td style='background-color: #CCFFFF'><b>Row</b></td>"&"<td style='background-color: #74A4BC'><b>"&Request.Form("file1")&" Contents</b></td><td style='background-color: #74A4BC'><b>"&Request.Form("file2")&" Contents</b></td>"&"<td style='background-color: #FFFFCC'><b>Column</b></td>"&"<td style='background-color: #CCFFFF'><b>Row</b></td></tr>"
finalsplit = Split(Request.Form("finaldiff"), "\|\")

I = 0
Do While I < Ubound(finalsplit)
    Pr "<tr><td style='width:50px; background-color: #FFFFCC' >" &finalsplit(I)& "</td>"
    I = I + 1
    Pr "<td style='width:50px; background-color: #CCFFFF'>"&finalsplit(I)&" </td>"
    I = I + 1
    Pr "<td style='width:250px;'>"&finalsplit(I)&" </td>"
    I = I + 1
    Pr "<td style='width:250px;'>"&finalsplit(I)&" </td>"
    I = I + 1
    Pr "<td style='width:50px; background-color: #FFFFCC'>"&finalsplit(I)&" </td>"
    I = I + 1
    Pr "<td style='width:50px; background-color: #CCFFFF'>"&finalsplit(I)&" </td> </tr>"
    I = I + 1
Loop
Pr "</table></center>"
%>