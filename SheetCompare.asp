<html>
<Body>
<%

Sub Pr(S)
    Response.Write S
    End Sub

Pr Request.Form("sheet") & " Comparison<br>"
Pr "<center><table border='1' cellspacing='0'><tr><td></td></tr>"
diffsplit = Split(Request.Form("diff"), "*")
indexsplit = Split(Request.Form("index"), "*") 'row, column, row, column'
Dim I, value, prev, curr, newRow
I = 0
J = 0
newRow = 1
Do While I < Ubound(diffsplit)
    value = diffsplit(I)
    If I = 0 Then
        prev = indexsplit(J)
        curr = indexsplit(J)
        I = I + 1
        J = J + 2
    Else
        prev = curr
        J = J + 2
        curr = indexsplit(J)
        I = I + 1
    End If

    If newRow = 1 Then
        Pr "<tr>"
        newRow = 0
    End If

    If prev = curr Then
        Pr "<td style='width:30px'>"&value&"</td>"
    Else
        Pr "<td style='width:30px'>"&value&"</td></tr>"
        newRow = 1
    End If

    Loop
Pr "</table></center>"
%>
