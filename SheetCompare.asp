<html>
<Body>
<%
Sub Pr(S)
    Response.Write S
    End Sub
Dim I, differencesplit, lastRow
Pr Request.Form("sheet") & " Comparison<br>"
Pr "<center><table border='1' cellspacing='0'><tr><td><b>Location</b></td>"&"<td><b>"&Request.Form("file1")&" Contents</b></td><td><b>"&Request.Form("file2")&" Contents</b></td></tr>"
differencesplit = Split(Request.Form("finaldiff"), "\")
I = 0
Do While I < Ubound(differencesplit)
    If I/3 = Int(I/3) Then
        Pr "<tr>"
        lastRow = I
        End If
    Pr "<td style='width:400px'>"&differencesplit(I)&"</td>"
    If I = lastRow + 2 Then
        Pr "</tr>"
        End If
    I = I + 1
    Loop
Pr "</table></center>"
%>