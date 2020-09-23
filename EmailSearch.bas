Attribute VB_Name = "EmailSearch"
Public fso
Public extarray() As String
Dim strings As String

Public Sub SearchForEmail(Fi As String)
On Error GoTo Errh:
If FrmOptions.Check3.Value = 1 And FileLen(Fi) > FrmOptions.Text2 * 1024 Then Exit Sub
strings = " ~,~" & Chr(34) & "~=~:~(~)~<~>~'~;"
Dim temp As String: Dim temp1() As String: Dim temp2() As String
Dim temp3() As String
Set thefi = fso.opentextfile(Fi)
For a = 1 To 1000000
temp = thefi.ReadLine
If InStr(temp, "@") = False Or InStr(temp, ".") = False Then GoTo nxt:
emailfromline temp, 0
nxt:
Next
Errh:
If Err.Number <> 62 Then
frmerror.Visible = True
frmerror.Label3.Caption = 10
frmerror.Timer1.Enabled = True
frmerror.Label1.Caption = "Error: " & Err.Number & " occured." & vbNewLine & Err.Description
epp = 1
While epp = 1
DoEvents
Wend
End If
End Sub

Private Sub emailfromline(ttemp As String, poss As Integer)
Dim temp4() As String
Dim strni As String
spliter = Split(strings, "~")(poss)
temp4() = Split(ttemp, spliter)
For Each strn In temp4()
If InStr(strn, "@") > 0 And InStr(strn, ".") > 0 And poss >= 10 Then addtoem strn: Exit Sub
strni = strn
If poss < 10 Then emailfromline strni, poss + 1
Next
End Sub

Public Sub addtoem(theemai)
Dim theemail As String
If Len(theemai) > 39 Then Exit Sub
            If FrmOptions.Check4.Value = 1 Then
For ax = 0 To 255
If InStr(" 1234567890@_-qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM.", Chr(ax)) = 0 Then
If InStr(theemai, Chr(ax)) Then Exit Sub
End If
Next
             Else
  For ax = 181 To 255
  If InStr(theemai, Chr(ax)) Then Exit Sub
  Next
    For ax = 0 To 27
  If InStr(theemai, Chr(ax)) Then Exit Sub
  Next
             End If
theemail = theemai
'No Duplicate Check As that severly slows things down.
If InStr(theemail, "@") > 1 And InStr(theemail, ".") > InStr(theemail, "@") Then ' dot must be after @
If Right$(theemail, 1) = "." Then theemail = Left$(theemail, Len(theemail) - 1)
Form1.List2.AddItem Trim(theemail)
End If
End Sub
