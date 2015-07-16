Attribute VB_Name = "Module1"
Public Base_Datos  As New ADODB.Connection
Public rs As ADODB.Recordset
Public TablaExiste As New ADODB.Recordset
Public SOpt As Variant
Public Sql As String
Public titulo As String
Public I As Integer
Public CodigoCata As String
Public Escogio As Boolean



Public Sub OpenTabla(Tabla As Recordset, Sql As String)
Set Tabla = New ADODB.Recordset
Tabla.Open Sql, Base_Datos
End Sub

Public Function Apost(Texto As String) As String
Apost = "'" & Texto & "'"
End Function
Public Function Numeral(Texto As String) As String
Numeral = "#" & Texto & "#"
End Function
Public Function Existe(Tabla As String, CampoClave As String, Valor As String) As Boolean
Sql = "SELECT * FROM " & Tabla & " WHERE " & CampoClave & "=" & Apost(Valor)

Set TablaExiste = New ADODB.Recordset
TablaExiste.Open Sql, Base_Datos
If TablaExiste.EOF Then
   Existe = False
Else
   Existe = True
End If
End Function

Public Sub ValidaEntero(KeyAscii As Integer)
If KeyAscii = 8 Then
   Exit Sub
End If

If (Chr(KeyAscii) < "0") Or (Chr(KeyAscii) > "9") Then
   KeyAscii = 0
End If
End Sub

Public Sub ValidaLetras(KeyAscii As Integer)
If (KeyAscii = 8) Or (KeyAscii = 32) Then
   Exit Sub
End If

If (UCase(Chr(KeyAscii)) < "A") Or (UCase(Chr(KeyAscii)) > "Z") Then
   KeyAscii = 0
End If
End Sub

Public Function DateVal(Fecha As String) As String
DateVal = "DateValue(" & Apost(Fecha) & ")"
End Function






