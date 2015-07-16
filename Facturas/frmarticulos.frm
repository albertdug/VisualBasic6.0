VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmarticulos 
   Caption         =   "Actualización de Articulos"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CR 
      Left            =   6600
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Salir"
      Height          =   615
      Left            =   8400
      Picture         =   "frmarticulos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   6720
      Picture         =   "frmarticulos.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Ultimo"
      Height          =   495
      Left            =   5040
      TabIndex        =   17
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Siguiente"
      Height          =   495
      Left            =   3480
      TabIndex        =   16
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Anterior"
      Height          =   495
      Left            =   1920
      TabIndex        =   15
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Primero"
      Height          =   495
      Left            =   240
      Picture         =   "frmarticulos.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   8400
      Picture         =   "frmarticulos.frx":0418
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   6720
      Picture         =   "frmarticulos.frx":0562
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   5040
      Picture         =   "frmarticulos.frx":06AC
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   3480
      Picture         =   "frmarticulos.frx":07F6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   1920
      Picture         =   "frmarticulos.frx":0940
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Incluir"
      Height          =   495
      Left            =   240
      Picture         =   "frmarticulos.frx":0A8A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Precio del  Articulo :"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Existencia del  Articulo :"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Descripción del  Articulo :"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo del  Articulo :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmarticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cn As New ADODB.Connection
Dim Proceso As String
Function Existe() As Boolean
 Dim TB As New ADODB.Recordset
 Dim TSQL As String
 
 
 TSQL = "SELECT * FROM Articulos WHERE Codigo='" + Text1.Text + "'"
 TB.Open TSQL, Cn
 If Not TB.EOF Then
    ExisteNum = True
 Else
    ExisteNum = False
 End If
 TB.Close
End Function
Sub Mostrar(cadcodigo)
Dim TSQL As String
Dim TB As New ADODB.Recordset
Dim cadsexo As String
Dim cadestadoc As String

TSQL = "SELECT * FROM Articulos Where Codigo='" + cadcodigo + "'"
TB.Open TSQL, Cn
If Not TB.EOF Then
   Text1.Text = TB("Codigo")
   Text2.Text = TB("Descripcion")
   Text3.Text = CStr(TB("Existencia"))
   Text4.Text = CStr(TB("Preciou"))

End If
TB.Close
End Sub
Sub Limpiar()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub
Sub Activar(ByVal Proceso As String)
If Proceso = "I" Then
   Text1.Enabled = True
End If
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
End Sub
'Este procedimiento desactiva los controles
'principales del formulario es decir no responden
'a los eventos generados por el usuario
Sub Desactivar()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
End Sub
'Procedimiento que desactiva o activa controles
'segun el botón que se presiona
'I-- Incluir
'M-- Modificar
'B-- Buscar
'C-- Cancelar
Sub Botones(ByVal Proceso As String)
Select Case Proceso
Case "I"
     Activar (Proceso)
     Command1.Enabled = False
     Command2.Enabled = False
     Command3.Enabled = False
     Command4.Enabled = False
     Command5.Enabled = True
     Command6.Enabled = True
 Case "M"
     Activar (Proceso)
     Command1.Enabled = False
     Command2.Enabled = False
     Command3.Enabled = False
     Command4.Enabled = False
     Command5.Enabled = True
     Command6.Enabled = True
Case "B"
     Desactivar
     Command1.Enabled = False
     Command2.Enabled = True
     Command3.Enabled = False
     Command4.Enabled = True
     Command5.Enabled = False
     Command6.Enabled = True
Case Else
     Desactivar
     Command1.Enabled = True
     Command2.Enabled = False
     Command3.Enabled = True
     Command4.Enabled = False
     Command5.Enabled = False
     Command6.Enabled = False
End Select
End Sub
Private Sub Command1_Click()
Proceso = "I"
Botones (Proceso)
Limpiar
Text1.SetFocus
End Sub
Private Sub Command10_Click()
Dim TB As New ADODB.Recordset
Dim TSQL As String
TSQL = "SELECT * FROM  Articulos Order by Val(Codigo) desc"
TB.Open TSQL, Cn
If Not TB.EOF Then
   TB.MoveFirst
   Text1.Text = TB("Codigo")
   Mostrar (Text1.Text)
End If
TB.Close
End Sub

Private Sub Command11_Click()
Dim Res As Byte
If Text1.Text <> "" Then
   CR.ReportFileName = App.Path & "\RepArticulo.rpt"
   CR.SelectionFormula = "({Articulos.Codigo}='" + Text1.Text + "')"
   CR.PrintReport
Else
  Res = MsgBox("Error Debe Escoger el Cliente que desea imprimir", 0 + 64 + 0, "Información")
End If

End Sub

Private Sub Command12_Click()

Unload Me
End Sub

Private Sub Command2_Click()
Proceso = "M"
Botones (Proceso)
End Sub

Private Sub Command3_Click()
Dim TB As New ADODB.Recordset
 Dim TSQL As String
 
 
 TSQL = "SELECT * FROM Articulos"
 TB.Open TSQL, Cn
 If Not TB.EOF Then
   frmcatalogoArticulos.Show vbModal, Me
   Proceso = "B"
   Botones (Proceso)
 Else
    Res = MsgBox("No hay Registros", 64, "Información")
 End If
 
End Sub

Private Sub Command4_Click()
If MsgBox("¿Está Seguro de Eliminar Este Registro?", vbQuestion + vbYesNo, "Pregunta") = vbYes Then
   TSQL = "DELETE  FROM Articulos WHERE Codigo='" + Text1.Text + "'"
   Cn.Execute TSQL
   MsgBox "Registro Eliminado", vbInformation, "Mensaje"
End If
Proceso = "C"
Botones (Proceso)
Limpiar
End Sub

Private Sub Command5_Click()
Dim TSQL As String
Dim Res As Byte
If (Text1.Text <> "") And (Text2.Text <> "") And (Text3.Text <> "") And (Text4.Text <> "") Then
    If Proceso = "I" Then
      If Existe = True Then
         Res = MsgBox("Codigo ya Existe,Verifique", 0 + 64 + 0, "Actualización")
         Text1.Text = ""
         Text1.SetFocus
      Else
         TSQL = "INSERT INTO Articulos (" & _
         "Codigo,Descripcion,Existencia,Preciou) " & _
                   "VALUES(" & _
                   "'" + Text1.Text + "'," & _
                   "'" + Text2.Text + "'," & _
                   "'" + Text3.Text + "'," & _
                   "'" + Text4.Text + "')"
                  Cn.Execute TSQL
                   MsgBox "Registro Incluido", vbInformation, "Mensaje"
                   
      End If
   End If
   
   If Proceso = "M" Then
      TSQL = "UPDATE Articulos SET " & _
      "Descripcion='" + Text2.Text + "'," & _
      "Existencia='" + Text3.Text + "'," & _
      "Preciou='" + Text4.Text + "' WHERE Codigo='" + Text1.Text + "'"
      Cn.Execute TSQL
       MsgBox "Registro Actualizado", vbInformation, "Mensaje"
   End If
   
   
   Proceso = "C"
   Botones (Proceso)
   Limpiar
Else
  Res = MsgBox("Hay campos en blanco", 0 + 64 + 0, "Información")
End If
End Sub

Private Sub Command6_Click()
Proceso = "C"
Botones (Proceso)
Limpiar
End Sub

Private Sub Command7_Click()
Dim TB As New ADODB.Recordset
Dim TSQL As String
TSQL = "SELECT * FROM  Articulos Order by Val(Codigo)"
TB.Open TSQL, Cn
If Not TB.EOF Then
   TB.MoveFirst
   Text1.Text = TB("Codigo")
   Mostrar (Text1.Text)
End If
TB.Close
End Sub

Private Sub Command8_Click()
Dim Res As Byte
Dim TB2 As New ADODB.Recordset
Dim TSQL As String
Dim cadcedula As String
Dim Enc As Boolean

If Text1.Text <> "" Then
  numcodigo = CInt(Text1.Text)
  TSQL = "SELECT * FROM Articulos order by Val(Codigo) desc"
  TB2.Open TSQL, Cn
  TB2.MoveFirst
  Enc = False
  
  While Not TB2.EOF() And Not (Enc)
    numnumero = CInt(TB2("Codigo"))
    If numcodigo > numnumero Then
     Enc = True
     Mostrar (TB2("Codigo"))
   Else
     TB2.MoveNext
  End If
  Wend
  
  
     
  
End If
TB2.Close
End Sub

Private Sub Command9_Click()
Dim Res As Byte
Dim TB2 As New ADODB.Recordset
Dim TSQL As String
Dim cadcedula As String
Dim Enc As Boolean

If Text1.Text <> "" Then
  numcodigo = CInt(Text1.Text)
  TSQL = "SELECT * FROM Articulos order by Val(Codigo)"
  TB2.Open TSQL, Cn
  TB2.MoveFirst
  Enc = False
  
  While Not TB2.EOF() And Not (Enc)
    numnumero = CInt(TB2("Codigo"))
    If numcodigo < numnumero Then
     Enc = True
     Mostrar (TB2("Codigo"))
   Else
     TB2.MoveNext
  End If
  Wend
  
  
     
  
End If
TB2.Close

End Sub

Private Sub Form_Load()
Cn.Open "DSN=sisfact; UID=Admin; PWD=123;"
Proceso = "C"
Botones (Proceso)
Limpiar
End Sub
Private Sub Text1_LostFocus()
Dim TB As New ADODB.Recordset
Dim TSQL As String
  If Proceso = "I" Then
     Text1.Text = Trim(Text1.Text)
     If Text1.Text <> "" Then
        TSQL = "SELECT * FROM Articulos WHERE Codigo='" + Text1.Text + "'"
        TB.Open TSQL, Cn
        If Not TB.EOF Then
           MsgBox "Ya Existe Un Articulo con ese codigo", vbInformation, "Mensaje"
           Limpiar
           Text1.SetFocus
        End If 'Ya existe el registro
        TB.Close
     Else
       Limpiar
       Text1.SetFocus
     End If
  End If ' IMEC="I"

End Sub

Private Sub Text2_Change()
Dim Cadena As String
Cadena = QBlancos(Text2.Text)
Text2.Text = Cadena
Text2.SelStart = Len(Text2.Text)
End Sub


