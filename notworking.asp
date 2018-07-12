<html>
<Body>
<%
Sub Pr(S)
    Response.Write S
    End Sub
Dim I, differencesplit, lastRow, file1contents, file2contents
Pr Request.Form("sheet") & " Comparison<br>"
Pr "<center><table border='1' cellspacing='0'><tr><td><b>Column</b></td>"&"<td><b>Row</b></td>"&"<td><b>"&Request.Form("file1")&" Contents</b></td><td><b>"&Request.Form("file2")&" Contents</b></td>"&"<td><b>Column</b></td>"&"<td><b>Row</b></td></tr>"
file1contents = Split(Request.Form("finaldiff1"), "\|\")
file2contents = Split(Request.Form("finaldiff2"), "\|\")
Dim A, B, origA, origB, stopper, finalstring, insertedLength, origLength, check
origA = 0
origB = 0
A = 0
B = 0
Do While A < Ubound(file1contents)
	stopper = 0
	Do While B < Ubound(file2contents)
		If ((file2contents(B+1) - file1contents(A+1)) < 5 And stopper = 0) Then 'rows are at most 5 apart'
			If (Strcomp(file1contents(A), file2contents(B)) = 0) Then 'columns are the same'
				insertedLength = file2contents(B+1) - file1contents(A+1)
				'we have an insertion'
				check = 0
				origA = A
				origB = B
				origLength = insertedLength
				Do While insertedLength > 0
					If (Strcomp(file1contents(A+2), file2contents(B+2)) <> 0) Then
						insertedLength = insertedLength - 1
						A = A + 3
						B = B + 3
					Else
						check = 1
					End If
				Loop
				B = origB
				If (check = 0) Then
					Do While origLength > 0
						finalstring = finalstring & "inserted\" & "inserted\" & "inserted\" & file2contents(B+2) & "\" & file2contents(B+1) & "\" & file2contents(B) & "\"
						B = B + 3
						origLength = origLength - 1

					Loop


					finalstring = finalstring & file1contents(origA) & "\" & file1contents(origA+1) & "\" & file1contents(origA+2) & "\" & file2contents(origB+2) & "\" & file2contents(origB+1) & "\" & file2contents(origB) & "\"
				End If

			End If
		Else
			stopper = 1
		End If
		B = B + 3
		Loop
	'concatenate'
	If (stopper = 1) Then
		finalstring = finalstring & file1contents(origA) & "\" & file1contents(origA+1) & "\" & file1contents(origA+2) & "\" & file2contents(origA+2) & "\" & file2contents(origA+1) & "\" & file2contents(origA) & "\"
	End If
	origA = origA + 3
	A = origA
	Loop

finalsplit = Split(finalstring, "\")

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