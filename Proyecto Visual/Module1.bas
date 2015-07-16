Attribute VB_Name = "Module1"
'Fuentes Modulo del  Proyecto
Global usuario As String
Global NivelUsuario As Integer
Function QBlancos(Cadena As String) As String
Dim Cantidad As Integer
Dim i As Integer
Dim J As Integer
Dim Aux As String
Dim Cad As String
Dim CadB As String

i = 1
Cadena = LTrim(Cadena)
Cantidad = Len(Cadena)

While (i <= Cantidad)
Cad = Mid(Cadena, i, 1)

If Cad = " " Then
   Aux = Aux + " "
   J = i
   Do
     J = J + 1
     CadB = Mid(Cadena, J, 1)
   Loop While (CadB = " ")
   i = J
Else
   Aux = Aux + Cad
   i = i + 1
End If
Wend
QBlancos = Aux
End Function

