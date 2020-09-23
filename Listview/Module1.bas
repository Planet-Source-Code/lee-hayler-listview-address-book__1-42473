Attribute VB_Name = "Module1"
Dim Email() As String
Dim Icount%
Dim S1 As Variant
Dim S2 As Variant
Public Function Extract(InptStr As String, _
firstDelimiter As String, _
secondDelimiter As String, _
ByVal Keychar As String) As String
S1 = Split(InptStr, firstDelimiter)
For Each f In S1
Icount = Icount + 1
ReDim Preserve Email(Icount)
Email(Icount) = f
S2 = Split(Email(Icount), secondDelimiter)
For Each t In S2
Email(Icount) = t
If InStr(1, Email(Icount), Keychar) > 0 Then
Extract = Extract & t
End If
Next
Next
Icount = 0
End Function


