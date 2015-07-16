VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmfacturas 
   Caption         =   "Form1"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   13695
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   13695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   11520
      TabIndex        =   43
      Top             =   7320
      Width           =   1575
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   11400
      TabIndex        =   41
      Top             =   6840
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   11400
      TabIndex        =   39
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Salir"
      Height          =   495
      Left            =   10320
      Picture         =   "frmfacturas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   8880
      Picture         =   "frmfacturas.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   7200
      Picture         =   "frmfacturas.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   5880
      Picture         =   "frmfacturas.frx":03DE
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   4440
      Picture         =   "frmfacturas.frx":0528
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   3000
      Picture         =   "frmfacturas.frx":0672
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   1680
      Picture         =   "frmfacturas.frx":07BC
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Emitir"
      Height          =   495
      Left            =   240
      Picture         =   "frmfacturas.frx":0906
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Eliminar Articulos"
      Height          =   495
      Left            =   6840
      TabIndex        =   29
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Agregar Articulos"
      Height          =   495
      Left            =   6840
      TabIndex        =   28
      Top             =   720
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   2655
      Left            =   120
      TabIndex        =   27
      Top             =   3720
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      FormatString    =   $"frmfacturas.frx":0A50
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   11040
      TabIndex        =   25
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   8880
      TabIndex        =   23
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   6720
      TabIndex        =   21
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   2160
      TabIndex        =   19
      Top             =   3240
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   ".."
      Height          =   375
      Left            =   1440
      TabIndex        =   18
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Forma de Pago"
      Height          =   735
      Left            =   10680
      TabIndex        =   13
      Top             =   720
      Width           =   2655
      Begin VB.OptionButton OptionCre 
         Caption         =   "Credito"
         Height          =   255
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptionCon 
         Caption         =   "Contado"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   2280
      Width           =   3855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   ".."
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   600
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTFecha 
      Height          =   255
      Left            =   11640
      TabIndex        =   2
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Format          =   21561345
      CurrentDate     =   36892
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label14 
      Caption         =   "Total :"
      Height          =   375
      Left            =   10200
      TabIndex        =   42
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Impuesto :"
      Height          =   255
      Left            =   10200
      TabIndex        =   40
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Subtotal"
      Height          =   255
      Left            =   10200
      TabIndex        =   38
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Importe"
      Height          =   375
      Left            =   11280
      TabIndex        =   26
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "Precio del Articulo"
      Height          =   255
      Left            =   9120
      TabIndex        =   24
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   6720
      TabIndex        =   22
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Descripción del  Articulo :"
      Height          =   255
      Left            =   2160
      TabIndex        =   20
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label Label7 
      Caption         =   "Codigo del  Articulo "
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Dirección del  Cliente:"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "telefono del Cliente :"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Nombre del  Cliente :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Cedula del  Cliente :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha de la Factura :"
      Height          =   255
      Left            =   9960
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de la Factura :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmfacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Proceso As String
Sub Limpiar()
Text1.Text = ""
DTFecha.Value = Date
OptionCon.Value = True
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
FG.Rows = 1
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
End Sub
Sub RestarExistencia(CodArt As String, Cantidad As Long)
Dim TSQL As String
Dim TB As New ADODB.Recordset
Dim CExistencia As Long
Dim Resto As Long
TSQL = "SELECT * FROM Articulos Where Codigo='" + CodArt + "'"
TB.Open TSQL, Cn
If Not TB.EOF Then
   CExistencia = TB("Existencia")
   Resto = CExistencia - Cantidad
   TSQL = "UPDATE Articulos SET " & _
        "Existencia='" + Resto + "' WHERE Codigo='" + CodArt + "'"
      Cn.Execute TSQL
End If
TB.Close
End Sub
Sub SumarExistencia(CodArt As String, Cantidad As Long)
Dim TSQL As String
Dim TB As New ADODB.Recordset
Dim CExistencia As Long
Dim Resto As Long
TSQL = "SELECT * FROM Articulos Where Codigo='" + CodArt + "'"
TB.Open TSQL, Cn
If Not TB.EOF Then
   CExistencia = TB("Existencia")
   Resto = CExistencia + Cantidad
   TSQL = "UPDATE Articulos SET " & _
        "Existencia='" + Resto + "' WHERE Codigo='" + CodArt + "'"
      Cn.Execute TSQL
End If
TB.Close
End Sub


