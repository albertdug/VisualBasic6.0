VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form fmrmateriasxestudiante 
   Caption         =   "Materias_Estudiante"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1680
      Top             =   4680
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\trabajo visual\bd.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\trabajo visual\bd.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   6735
      Begin VB.CommandButton btneliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4200
         Picture         =   "materiasxestudiante.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton btnguardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3120
         Picture         =   "materiasxestudiante.frx":054A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton btncancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         Picture         =   "materiasxestudiante.frx":0E14
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Elimina el registro Actual"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton btnincluir 
         Caption         =   "&Incluir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Picture         =   "materiasxestudiante.frx":105E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Agrega un nuevo registros"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton btnmodificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         Picture         =   "materiasxestudiante.frx":14DB
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Modifica el Registro Actual"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton btnsalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5520
         Picture         =   "materiasxestudiante.frx":1974
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "salir"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "MeteriasxEstudiantes"
      Height          =   2655
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      Begin VB.CommandButton btnconsultarm 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         Picture         =   "materiasxestudiante.frx":1EF5
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "consultar clientes"
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton btnconsultare 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         Picture         =   "materiasxestudiante.frx":2319
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "consultar clientes"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtnota 
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtcodmat 
         Height          =   405
         Left            =   2760
         TabIndex        =   2
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtcodest 
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nota:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codigo de la materia:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   2100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo del Estudiante:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2250
      End
   End
End
Attribute VB_Name = "fmrmateriasxestudiante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnconsultare_Click()
fmrCataestudiante.CmdAceptar.Enabled = True
fmrCataestudiante.Show 1

If Escogio Then
  txtcodest.Text = CodigoCata
  btnmodificar.Enabled = True
  btnguardar.Enabled = False
  btneliminar.Enabled = True
'  TxtNombre.SetFocus
'FrameDatos.Enabled = True
  
End If
End Sub

Private Sub btnconsultarm_Click()
fmrCatmaterias.CmdAceptar.Enabled = True
fmrCatmaterias.Show 1

   If Escogio Then
     txtcodmat.Text = CodigoCata
     btnmodificar.Enabled = True
     btnguardar.Enabled = False
     btneliminar.Enabled = True
   ' TxtNombre.SetFocus
   ' FrameDatos.Enabled = True
  End If
End Sub

Private Sub btnsalir_Click()
   If MsgBox(" Esta seguro que deseas salir ", vbYesNo + vbQuestion) = vbYes Then
       Unload Me
   End If
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection '( base de datos)
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\bd.mdb;Persist Security Info=False"
SOpt = dbSQLpasstrhough
End Sub

