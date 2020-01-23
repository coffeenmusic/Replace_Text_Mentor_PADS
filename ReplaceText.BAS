' Description
' Find and replace text in Mentor PADs Layout

Public Function ReplaceText (find, repl)
	cnt = 0
	For Each t In ActiveDocument.Texts
		If InStr(t.Text, "CONTROLLER") Then
			newStr = Replace(t.Text, find, repl)
			t.Text = newStr
			cnt = cnt + 1
		End If
	Next t
	MsgBox("Successfully replaced " + Format(cnt,"0") + " strings.")
End Function

Sub Main
	Begin Dialog UserDialog 360,112,"Center Component From Pins"
		CancelButton 200,77,110,21
		Text 30,17,40,14,"Find",.Text1
		TextBox 110,14,230,21,.txtFind
		Text 30,45,60,14,"Replace",.lblReplace
		TextBox 110,42,230,21,.txtReplace
		OKButton 40,77,120,21,.btnOK
	End Dialog
	Dim dlg As UserDialog
	DialogOK = Dialog(dlg)
	If DialogOK Then
		find = dlg.txtFind
		repl = dlg.txtReplace
		If (find <> "") And (repl <> "") Then
			ReplaceText (find, repl)
		Else
			MsgBox("Find and Replace text boxes must not be empty.")
		End If
	End If	
End Sub
