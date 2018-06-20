<html>
<Body>
<%

Sub Pr(S)
    Response.Write S
    End Sub

Pr Request.Form("sheet") & " Comparison<br>"
Pr "<center><table border='1' cellspacing='0'><tr><td>Cell Number</td></tr>"
diffsplit = Split(Request.Form("diff"), "*")
Dim I, position
I = 0
Do While I < Ubound(diffsplit)
    position = diffsplit(I)
    I = I + 1
    Pr "<tr>"
    Pr "<td>"&position&"</td></tr>"
    Loop
Pr "</table></center>"
%>
