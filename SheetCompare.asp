<html>
<Body>
<%

Sub Pr(S)
    Response.Write S
    End Sub

Pr Request.Form("sheet") & " Comparison<br>"
Pr "<center><table border='1' cellspacing='0'><tr><td>Differences</td></tr>"
diffsplit = Split(Request.Form("diff1"), "*")
indexsplit = Split(Request.Form("index1"), "*") 'row, column, row, column'
diffsplit2 = Split(Request.Form("diff2"), "*")
indexsplit2 = Split(Request.Form("index2"), "*") 'row, column, row, column'
Dim I, value, prev, curr, X, Y, value2
I = 0
J = 0
Do While I < Ubound(diffsplit)
    If indexsplit(J) = indexsplit2(J) Then
        If indexsplit(J+1) = indexsplit2(J+1) Then
            value = diffsplit(I)
            value2 = diffsplit2(I)
            If I = 0 Then
                prev = indexsplit(J)
                curr = indexsplit(J)
                X = indexsplit(J)
                Y = indexsplit(J)
                I = I + 1
                J = J + 2
            Else
                prev = curr
                J = J + 2
                curr = indexsplit(J)
                X = indexsplit(J)
                Y = indexsplit(J+1)
                I = I + 1
            End If

            Pr "<tr> <td style='width:200px'>("&X&","&Y&"): "&value&" vs "&value2&"</td></tr>"
        End If
    End If
    Loop
Pr "</table></center>"
%>
