VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fmrCataestudiante 
   Caption         =   "catalogo de estudiante"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   1935
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   10
      FixedCols       =   0
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   3480
      Width           =   8295
      Begin VB.CommandButton CmdAceptar 
         BackColor       =   &H80000016&
         Caption         =   "Aceptar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5640
         Picture         =   "FormCataestudiante.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
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
         Left            =   4560
         Picture         =   "FormCataestudiante.frx":0413
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1092
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
         Left            =   1200
         Picture         =   "FormCataestudiante.frx":095D
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Modifica el Registro Actual"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton btnnuevo 
         Caption         =   "&Incluir"
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
         Left            =   120
         Picture         =   "FormCataestudiante.frx":0DF6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Agrega un nuevo registros"
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
         Left            =   2400
         Picture         =   "FormCataestudiante.frx":1273
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Elimina el registro Actual"
         Top             =   240
         Width           =   1095
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
         Left            =   3480
         Picture         =   "FormCataestudiante.frx":14BD
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdSalir 
         BackColor       =   &H80000016&
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7200
         Picture         =   "FormCataestudiante.frx":1D87
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   6015
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5040
      Top             =   1680
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
End
Attribute VB_Name = "fmrCataestudiante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAceptar_Click()
CodigoCata = Grid.TextMatrix(Grid.RowSel, 0)
Escogio = True
Unload Me
End Sub

Private Sub CmdSalir_Click()
If MsgBox(" Esta seguro que deseas salir ", vbYesNo + vbQuestion) = vbYes Then
       Unload Me
       End If
End Sub

Private Sub Form_Load()
   titulo$ = " Codigo del Estudiante |^Nombre                                   "
    Grid.FormatString = titulo$
    Set rs = New ADODB.Recordset
    Sql = "select Ci, nombre from estudiantes where estatus = 'A'"
    rs.Open Sql, db, adOpenStatic
    If Not rs.EOF Then
        I = 1
        While Not rs.EOF
            Grid.TextMatrix(I, 0) = rs!Ci
            Grid.TextMatrix(I, 1) = rs!nombre
            Grid.Rows = Grid.Rows + 1
            I = I + 1
            rs.MoveNext
        Wend
        End If
End Sub
