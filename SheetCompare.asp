<html>
<Body>
<%

Sub Pr(S)
    Response.Write S
    End Sub

Dim I, differencesplit

Pr Request.Form("sheet") & " Comparison<br>"
Pr "<center><table border='1' cellspacing='0'><tr><td>Differences ("&Request.Form("file1")&" vs "&Request.Form("file2")&") </td></tr>"
differencesplit = Split(Request.Form("finaldiff"), "\")
I = 0
Do While I < Ubound(differencesplit)
    Pr "<tr> <td style='width:400px'>"&differencesplit(I)&"</td></tr>"
    I = I + 1
    Loop
Pr "</table></center>"
%>
