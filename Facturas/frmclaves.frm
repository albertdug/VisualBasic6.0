VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmclaves 
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   11295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   495
      Left            =   360
      Picture         =   "frmclaves.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Incluir"
      Height          =   495
      Left            =   240
      Picture         =   "frmclaves.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   1920
      Picture         =   "frmclaves.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   3480
      Picture         =   "frmclaves.frx":03DE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   5040
      Picture         =   "frmclaves.frx":0528
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   6720
      Picture         =   "frmclaves.frx":0672
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   8400
      Picture         =   "frmclaves.frx":07BC
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nivel  del Usuario"
      Height          =   1335
      Left            =   5640
      TabIndex        =   6
      Top             =   480
      Width           =   2415
      Begin VB.OptionButton Option2 
         Caption         =   "Nivel 2"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nivel  1"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin Crystal.CrystalReport CR 
      Left            =   5400
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre del  Usuario :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Usuario :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Clave de Usuario :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmclaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cn As New ADODB.Connection
Dim Proceso As String
Dim cnivel As String
Sub llenarnivel()
If Option1.Value = True Then
   cnivel = "1"
Else
   cnivel = "2"
End If
End Sub
Function Existe() As Boolean
 Dim TB As New ADODB.Recordset
 Dim TSQL As String
 
 
 TSQL = "SELECT * FROM Usuarios WHERE Clave='" + Text1.Text + "'"
 TB.Open TSQL, Cn
 If Not TB.EOF Then
    ExisteNum = True
 Else
    ExisteNum = False
 End If
 TB.Close
End Function
Sub Mostrar(cadclave)
Dim TSQL As String
Dim TB As New ADODB.Recordset
Dim cadnivel

TSQL = "SELECT * FROM Usuarios Where Clave='" + cadclave + "'"
TB.Open TSQL, Cn
If Not TB.EOF Then
   Text1.Text = TB("Clave")
   Text2.Text = TB("Usuario")
   Text3.Text = TB("Nombre")
   cadnivel = TB("Nivel")
   
   If cadnivel = "1" Then
      Option1.Value = True
   Else
      Option2.Value = True
   End If

End If
TB.Close
End Sub
Sub Limpiar()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Option1.Value = True
Option2.Value = False
End Sub
Sub Activar(ByVal Proceso As String)
If Proceso = "I" Then
   Text1.Enabled = True
End If
Text2.Enabled = True
Text3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
End Sub
'Este procedimiento desactiva los controles
'principales del formulario es decir no responden
'a los eventos generados por el usuario
Sub Desactivar()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Option1.Enabled = True
Option2.Enabled = True
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


Private Sub Command2_Click()
Proceso = "M"
Botones (Proceso)
End Sub

Private Sub Command3_Click()
Dim TB As New ADODB.Recordset
 Dim TSQL As String
 
 
 TSQL = "SELECT * FROM Usuarios"
 TB.Open TSQL, Cn
 If Not TB.EOF Then
   frmcatalogoClaves.Show vbModal, Me
   Proceso = "B"
   Botones (Proceso)
 Else
    Res = MsgBox("No hay Registros", 64, "Información")
 End If
 
End Sub

Private Sub Command4_Click()
If MsgBox("¿Está Seguro de Eliminar Este Registro?", vbQuestion + vbYesNo, "Pregunta") = vbYes Then
   TSQL = "DELETE  FROM Usuarios WHERE Clave='" + Text1.Text + "'"
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
If (Text1.Text <> "") And (Text2.Text <> "") And (Text3.Text <> "") Then
    If Proceso = "I" Then
      If Existe = True Then
         Res = MsgBox("Clave ya Existe,Verifique", 0 + 64 + 0, "Actualización")
         Text1.Text = ""
         Text1.SetFocus
      Else
         llenarnivel
         TSQL = "INSERT INTO Usuarios (" & _
         "Clave,Usuario,Nombre,Nivel) " & _
                   "VALUES(" & _
                   "'" + Text1.Text + "'," & _
                   "'" + Text2.Text + "'," & _
                   "'" + Text3.Text + "'," & _
                   "'" + cnivel + "')"
                  Cn.Execute TSQL
                   MsgBox "Registro Incluido", vbInformation, "Mensaje"
                   
      End If
   End If
   
   If Proceso = "M" Then
      llenarnivel
      TSQL = "UPDATE Usuarios SET " & _
      "Usuario='" + Text2.Text + "'," & _
      "Nombre='" + Text3.Text + "'," & _
      "Nivel='" + cnivel + "' WHERE Clave='" + Text1.Text + "'"
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
Unload Me
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
        TSQL = "SELECT * FROM Usuarios WHERE Clave='" + Text1.Text + "'"
        TB.Open TSQL, Cn
        If Not TB.EOF Then
           MsgBox "Ya Existe Un Usuario con ese clave", vbInformation, "Mensaje"
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




