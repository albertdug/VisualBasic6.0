Attribute VB_Name = "Module1"
Public Conexion As New ADODB.Connection
Public GLOBALDATE As Date

Sub Conectar()
Conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Base de datos1.mdb;Persist Security Info=False"
End Sub

Sub LlenarCombo(c As ComboBox, tabla, campo)
Dim t As New ADODB.Recordset
c.Clear
t.Open tabla, Conexion
While Not t.EOF
c.AddItem t(campo)
t.MoveNext
Wend
End Sub
