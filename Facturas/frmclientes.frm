VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmclientes 
   Caption         =   "Actualización de Clientes"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   24
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   22
      Top             =   840
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sexo del Cliente "
      Height          =   1215
      Left            =   120
      TabIndex        =   19
      Top             =   1440
      Width           =   1695
      Begin VB.OptionButton OptionM 
         Caption         =   "Masculino"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton OptionF 
         Caption         =   "Femenino"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Estado Civil del Cliente"
      Height          =   1815
      Left            =   2040
      TabIndex        =   14
      Top             =   1440
      Width           =   2055
      Begin VB.OptionButton OptionS 
         Caption         =   "Soltero(a)"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton OptionC 
         Caption         =   "Casado(a)"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptionV 
         Caption         =   "Viudo(a)"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton OptionD 
         Caption         =   "Divorciado(a)"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1575
      End
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6120
      TabIndex        =   13
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   975
      Left            =   6240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   2160
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Incluir"
      Height          =   495
      Left            =   120
      Picture         =   "frmclientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   1680
      Picture         =   "frmclientes.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   3240
      Picture         =   "frmclientes.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   4920
      Picture         =   "frmclientes.frx":03DE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   6600
      Picture         =   "frmclientes.frx":0528
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   8400
      Picture         =   "frmclientes.frx":0672
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Primero"
      Height          =   495
      Left            =   120
      Picture         =   "frmclientes.frx":07BC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Anterior"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Siguiente"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Ultimo"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   6600
      Picture         =   "frmclientes.frx":0940
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Salir"
      Height          =   615
      Left            =   8400
      Picture         =   "frmclientes.frx":0A8A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   1455
   End
   Begin Crystal.CrystalReport CR 
      Left            =   6480
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Datos de Clientes"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker DTFecha 
      Height          =   375
      Left            =   9480
      TabIndex        =   23
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   21561345
      CurrentDate     =   36892
   End
   Begin VB.Label Label1 
      Caption         =   "Cedula del Cliente :"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha de Nacimiento:"
      Height          =   255
      Left            =   7320
      TabIndex        =   28
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre del Cliente :"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Telefono del Cliente :"
      Height          =   255
      Left            =   4440
      TabIndex        =   26
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Dirección del Cliente :"
      Height          =   255
      Left            =   4440
      TabIndex        =   25
      Top             =   2160
      Width           =   1575
   End
End
Attribute VB_Name = "frmclientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cn As New ADODB.Connection
Dim Sex As String
Dim Estado As String
Dim Proceso As String
Function Existe() As Boolean
 Dim TB As New ADODB.Recordset
 Dim TSQL As String
 
 
 TSQL = "SELECT * FROM Clientes WHERE Cedula='" + Text1.Text + "'"
 TB.Open TSQL, Cn
 If Not TB.EOF Then
    ExisteNum = True
 Else
    ExisteNum = False
 End If
 TB.Close
End Function
Sub Mostrar(cadcedula)
Dim TSQL As String
Dim TB As New ADODB.Recordset
Dim cadsexo As String
Dim cadestadoc As String

TSQL = "SELECT * FROM Clientes Where Cedula='" + cadcedula + "'"
TB.Open TSQL, Cn
If Not TB.EOF Then
   Text1.Text = TB("Cedula")
   Text2.Text = TB("Nombre")
   DTFecha.Value = TB("Fecnac")
   cadsexo = TB("Sexo")
   
   If cadsexo = "M" Then
      OptionM.Value = True
   Else
      OptionF.Value = True
   End If
   
   cadestadoc = TB("Estadoc")
   
   Select Case cadestadoc
          Case "Soltero"
               OptionS.Value = True
               OptionC.Value = False
               OptionV.Value = False
               OptionD.Value = False
          Case "Casado"
               OptionS.Value = False
               OptionC.Value = True
               OptionV.Value = False
               OptionD.Value = False
          Case "Viudo"
               OptionS.Value = False
               OptionC.Value = False
               OptionV.Value = True
               OptionD.Value = False
         Case Else
               OptionS.Value = False
               OptionC.Value = False
               OptionV.Value = False
               OptionD.Value = True
      End Select
      
      Text3.Text = TB("Telefono")
      Text4.Text = TB("Direccion")

End If
TB.Close
End Sub
Sub Limpiar()
Text1.Text = ""
Text2.Text = ""
DTFecha.Value = Date
OptionM.Value = True
OptionS.Value = True
Text3.Text = ""
Text4.Text = ""
End Sub
Sub Activar(ByVal Proceso As String)
If Proceso = "I" Then
   Text1.Enabled = True
End If
Text2.Enabled = True
DTFecha.Enabled = True
OptionM.Enabled = True
OptionF.Enabled = True
OptionS.Enabled = True
OptionC.Enabled = True
OptionV.Enabled = True
OptionD.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
End Sub
'Este procedimiento desactiva los controles
'principales del formulario es decir no responden
'a los eventos generados por el usuario
Sub Desactivar()
Text1.Enabled = False
Text2.Enabled = False
DTFecha.Enabled = False
OptionM.Enabled = False
OptionF.Enabled = False
OptionS.Enabled = False
OptionC.Enabled = False
OptionV.Enabled = False
OptionD.Enabled = False
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

Sub Llenarsexo()
If OptionM.Value = True Then
   Sex = "M"
Else
    Sex = "F"
End If
End Sub
Sub llenarestado()
If OptionS.Value = True Then
   Estado = "Soltero"
End If

If OptionC.Value = True Then
   Estado = "Casado"
End If

If OptionV.Value = True Then
   Estado = "Viudo"
End If

If OptionD.Value = True Then
   Estado = "Divorciado"
End If

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
TSQL = "SELECT * FROM  Clientes Order by Val(Cedula) desc"
TB.Open TSQL, Cn
If Not TB.EOF Then
   TB.MoveFirst
   Text1.Text = TB("Cedula")
   Mostrar (Text1.Text)
End If
TB.Close
End Sub

Private Sub Command11_Click()
Dim Res As Byte
If Text1.Text <> "" Then
   CR.ReportFileName = App.Path & "\RepCliente.rpt"
   CR.SelectionFormula = "({Clientes.Cedula}='" + Text1.Text + "')"
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
 
 
 TSQL = "SELECT * FROM Clientes"
 TB.Open TSQL, Cn
 If Not TB.EOF Then
   frmcatalogoClientes.Show vbModal, Me
   Proceso = "B"
   Botones (Proceso)
 Else
    Res = MsgBox("No hay Registros", 64, "Información")
 End If
 
End Sub

Private Sub Command4_Click()
If MsgBox("¿Está Seguro de Eliminar Este Registro?", vbQuestion + vbYesNo, "Pregunta") = vbYes Then
   TSQL = "DELETE  FROM Clientes WHERE Cedula='" + Text1.Text + "'"
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
         Res = MsgBox("Cedula ya Existe,Verifique", 0 + 64 + 0, "Actualización")
         Text1.Text = ""
         Text1.SetFocus
      Else
         Llenarsexo
         llenarestado
         Dim cadfecha As String
         cadfecha = CStr(DTFecha.Value)
         TSQL = "INSERT INTO Clientes (" & _
         "Cedula,Nombre,Fecnac,Sexo,Estadoc,Telefono,Direccion) " & _
                   "VALUES(" & _
                   "'" + Text1.Text + "'," & _
                   "'" + Text2.Text + "'," & _
                   "'" + cadfecha + "'," & _
                   "'" + Sex + "'," & _
                   "'" + Estado + "'," & _
                    "'" + Text3.Text + "'," & _
                   "'" + Text4.Text + "')"
                  Cn.Execute TSQL
                   MsgBox "Registro Incluido", vbInformation, "Mensaje"
                   
      End If
   End If
   
   If Proceso = "M" Then
      Llenarsexo
      llenarestado
      cadfecha = CStr(DTFecha.Value)
      TSQL = "UPDATE Clientes SET " & _
      "Nombre='" + Text2.Text + "'," & _
      "Fecnac='" + cadfecha + "'," & _
      "Sexo='" + Sex + "'," & _
      "Estadoc='" + Estado + "'," & _
      "Telefono='" + Text3.Text + "'," & _
      "Direccion='" + Text4.Text + "' WHERE Cedula='" + Text1.Text + "'"
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
TSQL = "SELECT * FROM  Clientes Order by Val(Cedula)"
TB.Open TSQL, Cn
If Not TB.EOF Then
   TB.MoveFirst
   Text1.Text = TB("Cedula")
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
  numcedula = CLng(Text1.Text)
  TSQL = "SELECT * FROM Clientes order by Val(Cedula) desc"
  TB2.Open TSQL, Cn
  TB2.MoveFirst
  Enc = False
  
  While Not TB2.EOF() And Not (Enc)
    numnumero = CLng(TB2("Cedula"))
    If numcedula > numnumero Then
     Enc = True
     Mostrar (CStr(numnumero))
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
  numcedula = CLng(Text1.Text)
  TSQL = "SELECT * FROM Clientes order by Val(Cedula)"
  TB2.Open TSQL, Cn
  TB2.MoveFirst
  Enc = False
  
  While Not TB2.EOF() And Not (Enc)
    numnumero = CLng(TB2("Cedula"))
    If numcedula < numnumero Then
     Enc = True
     Mostrar (CStr(numnumero))
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
        TSQL = "SELECT * FROM Clientes WHERE Cedula='" + Text1.Text + "'"
        TB.Open TSQL, Cn
        If Not TB.EOF Then
           MsgBox "Ya Existe Un Cliente con esa cedula", vbInformation, "Mensaje"
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
