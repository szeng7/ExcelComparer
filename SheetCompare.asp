<html>
<Body>
<%
Sub Pr(S)
    Response.Write S
    End Sub
Dim I, differencesplit, lastRow, file1contents, file2contents
Pr Request.Form("sheet") & " Comparison<br>"
Pr "<center><table border='1' cellspacing='0'><tr><td><b>Column</b></td>"&"<td><b>Row</b></td>"&"<td><b>"&Request.Form("file1")&" Contents</b></td><td><b>"&Request.Form("file2")&" Contents</b></td>"&"<td><b>Column</b></td>"&"<td><b>Row</b></td></tr>"
finalsplit = Split(Request.Form("finaldiff"), "\|\")

I = 0
Do While I < Ubound(finalsplit)
    Pr "<tr><td style='width:50px'>" &finalsplit(I)& "</td>"
    I = I + 1
    Pr "<td style='width:50px'>"&finalsplit(I)&" </td>"
    I = I + 1
    Pr "<td style='width:250px'>"&finalsplit(I)&" </td>"
    I = I + 1
    Pr "<td style='width:250px'>"&finalsplit(I)&" </td>"
    I = I + 1
    Pr "<td style='width:50px'>"&finalsplit(I)&" </td>"
    I = I + 1
    Pr "<td style='width:50px'>"&finalsplit(I)&" </td> </tr>"
    I = I + 1
Loop
Pr "</table></center>"
%>