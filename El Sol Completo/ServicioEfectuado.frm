VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ServicioEfectuado 
   BackColor       =   &H80000005&
   Caption         =   "Transporte EL SOL // Servicios Efectuados"
   ClientHeight    =   8730
   ClientLeft      =   2250
   ClientTop       =   480
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   Picture         =   "ServicioEfectuado.frx":0000
   ScaleHeight     =   8730
   ScaleWidth      =   11085
   Begin VB.CommandButton CMDsal 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      Picture         =   "ServicioEfectuado.frx":55AE
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton CMDgua 
      Caption         =   "&Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      Picture         =   "ServicioEfectuado.frx":5B38
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton CMDmod 
      Caption         =   "&Modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      Picture         =   "ServicioEfectuado.frx":60C2
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton CMDeli 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Picture         =   "ServicioEfectuado.frx":664C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8160
      Width           =   1815
   End
   Begin VB.TextBox TXTnro 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   1
      Top             =   2640
      Width           =   855
   End
   Begin VB.ComboBox CMBcli 
      Height          =   315
      Left            =   3360
      TabIndex        =   2
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox TXTnom 
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
      Height          =   375
      Left            =   5160
      MaxLength       =   30
      TabIndex        =   25
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox TXTori 
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
      Height          =   375
      Left            =   5880
      MaxLength       =   8
      TabIndex        =   24
      Top             =   4320
      Width           =   735
   End
   Begin VB.ComboBox CMBori 
      Height          =   315
      Left            =   3720
      TabIndex        =   4
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox TXTdes 
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
      Height          =   375
      Left            =   5880
      MaxLength       =   8
      TabIndex        =   6
      Top             =   4800
      Width           =   735
   End
   Begin VB.ComboBox CMBdes 
      Height          =   315
      Left            =   3720
      TabIndex        =   5
      Top             =   4800
      Width           =   2055
   End
   Begin VB.TextBox TXTpro 
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
      Height          =   375
      Left            =   5880
      MaxLength       =   8
      TabIndex        =   23
      Top             =   5520
      Width           =   735
   End
   Begin VB.ComboBox CMBpro 
      Height          =   315
      Left            =   3720
      TabIndex        =   8
      Top             =   5520
      Width           =   2055
   End
   Begin VB.TextBox TXTckg 
      Height          =   375
      Left            =   3720
      MaxLength       =   6
      TabIndex        =   9
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton CMDbus 
      Height          =   495
      Left            =   5160
      Picture         =   "ServicioEfectuado.frx":6BD6
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton CMDlim 
      Caption         =   "&Limpiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      Picture         =   "ServicioEfectuado.frx":7160
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2520
      Width           =   1815
   End
   Begin VB.ComboBox CMBpla 
      Height          =   315
      Left            =   3480
      TabIndex        =   22
      Text            =   "16"
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton CMDnue 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      Picture         =   "ServicioEfectuado.frx":76EA
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton CMDpro 
      Caption         =   "&Procesar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      Picture         =   "ServicioEfectuado.frx":7C74
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton CMDprc 
      Caption         =   "Proce&sar Camion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      Picture         =   "ServicioEfectuado.frx":81FE
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7440
      Width           =   1815
   End
   Begin VB.TextBox TXTprecio 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   6840
      TabIndex        =   20
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox TXTpso 
      Height          =   375
      Left            =   5760
      TabIndex        =   19
      Top             =   6840
      Width           =   855
   End
   Begin VB.TextBox TXTBxC 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   7
      Top             =   4680
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTser 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   3720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   54460417
      CurrentDate     =   39947
   End
   Begin MSComCtl2.DTPicker DTsol 
      Height          =   375
      Left            =   7440
      TabIndex        =   17
      Top             =   3720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   54460417
      CurrentDate     =   39947
   End
   Begin MSComCtl2.DTPicker DTglo 
      Height          =   375
      Left            =   7920
      TabIndex        =   27
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   54460417
      CurrentDate     =   39947
   End
   Begin VB.Label LBLnro 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Nro. de Solicitud"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   44
      Top             =   2760
      Width           =   1740
   End
   Begin VB.Label LBLcli 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Codigo del Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1320
      TabIndex        =   43
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label LBLori 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Ciudad de Origen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1320
      TabIndex        =   42
      Top             =   4320
      Width           =   1830
   End
   Begin VB.Label LBLdes 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Ciudad de Destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1320
      TabIndex        =   41
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label LBLcod 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Nombre del Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1320
      TabIndex        =   40
      Top             =   5520
      Width           =   2220
   End
   Begin VB.Label LBLkg 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Peso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1320
      TabIndex        =   39
      Top             =   6120
      Width           =   555
   End
   Begin VB.Label LBLkg 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Kg."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   5040
      TabIndex        =   38
      Top             =   6240
      Width           =   300
   End
   Begin VB.Label LBLsol 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fecha Solicitud"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   5640
      TabIndex        =   37
      Top             =   3840
      Width           =   1620
   End
   Begin VB.Label LBLser 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fecha Servicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   840
      TabIndex        =   36
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label LBLpla 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Placa del Camion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1320
      TabIndex        =   35
      Top             =   6960
      Width           =   1845
   End
   Begin VB.Label LBLpre 
      BackColor       =   &H80000005&
      Caption         =   "Bs.F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   34
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label LBLcarga 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   33
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Precio Tentativo de Servicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6840
      TabIndex        =   32
      Top             =   5640
      Width           =   2445
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Precio por Kg."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7200
      TabIndex        =   31
      Top             =   4320
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   1200
      Top             =   4200
      Width           =   8415
   End
   Begin VB.Shape Shape2 
      Height          =   1215
      Left            =   1200
      Top             =   5400
      Width           =   8415
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000005&
      Caption         =   "Bs.F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   30
      Top             =   4800
      Width           =   375
   End
   Begin VB.Shape Shape3 
      Height          =   1335
      Left            =   1200
      Top             =   6720
      Width           =   8415
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Peso de Carga Acumulado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2640
      TabIndex        =   29
      Top             =   7440
      Width           =   2250
   End
   Begin VB.Label LBLkg 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Kg."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   6360
      TabIndex        =   28
      Top             =   7440
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Efectuados"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8040
      TabIndex        =   18
      Top             =   1080
      Width           =   2280
   End
   Begin VB.Label LBLservefect 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Servicios "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8280
      TabIndex        =   0
      Top             =   480
      Width           =   1950
   End
End
Attribute VB_Name = "ServicioEfectuado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim acum As Integer
Private Sub CMBcli_Click()
Dim t As New ADODB.Recordset
        t.Open "select NomCli from Clientes where CodCli = '" & CMBcli.Text & "' ", Conexion
        txtnom.Text = t!nomcli
        DTser.SetFocus
End Sub
Private Sub CMBdes_Click()
Dim AUX As New ADODB.Recordset
Dim t As New ADODB.Recordset
       If CMBori.Text = "" Then
       MsgBox "Debe seleccionar una ciudad de origen", vbInformation
       CMBdes.Text = ""
       CMBori.SetFocus
       Else
        t.Open "select CodCiudad from Ciudades where Ciudades = '" & CMBdes.Text & "'", Conexion
        
        If Not t.EOF And CMBori.Text <> CMBdes.Text Then
          TXTdes.Text = t!CodCiudad
          CMBpro.SetFocus
        Else
          MsgBox (" No se permiten Viajes hacia la Misma Ciudad de Origen, Cambie la Condición"), vbCritical, "Datos Incorrectos"
          CMBdes.Text = ""
          CMBdes.SetFocus
        End If
        BuscarTab
       End If
End Sub
Private Sub CMBori_Click()
Dim t As New ADODB.Recordset
        t.Open "select CodCiudad from Ciudades where Ciudades = '" & CMBori.Text & "'", Conexion
        TXTori.Text = t!CodCiudad
        CMBdes.SetFocus
End Sub

Private Sub CMBpla_Click()
Dim t As New ADODB.Recordset
t.Open "select CargaMaximaKg from Camiones where Placa = '" & CMBpla.Text & "'", Conexion
  TXTpso.Text = t!CargaMaximaKg

End Sub

Private Sub CMDprc_Click()
 Dim BC As New ADODB.Recordset
  Dim DP As Boolean
  Dim Fecha As Date
  If Val(LBLcarga.Caption) > Val(TXTckg.Text) Then
  MsgBox "Ya tiene camiones suficientes para ese peso", vbInformation
  CMDprc.Enabled = False
  CMDgua.Enabled = True
  
  Else
     DP = False
   BC.Open "SELECT * FROM CamionesTemp where '" & CMBpla.Text & "' = placaB and FECHAP=#" & DTser.Value & "# ", Conexion
   
   If (BC.EOF = True) Then
        Conexion.Execute "insert into CamionesTemp (Numsol,PlacaB,Fechap,Peso,EstatusCP) values ('" & TXTnro.Text & "','" & CMBpla.Text & "','" & DTser.Value & "','" & TXTpso.Text & "', 'A')"
        MsgBox ("el camion fue apartado" & CMBpla)
        acum = acum + Val(TXTpso.Text)
        LBLcarga.Caption = acum
        DP = True
    
        Else
        MsgBox ("Camion ocupado en esa fecha...")
     End If
     CMDgua.Enabled = True
 End If
 
End Sub

Private Sub CMDpro_Click()
 Dim c As New ADODB.Recordset
  Dim BC As New ADODB.Recordset
  Dim tcamdis As New ADODB.Recordset
  Dim DP As Boolean
  Dim FechaS As Date
  c.Open "Select *from camiones", Conexion
     DP = False
  While (c.EOF = False) And (DP = False)
      If (c(4) - (c(4) * 0.1) >= Val(TXTckg.Text)) Then
      BC.Open "SELECT * FROM CamionesTemp where '" & c(0) & "' = placaB and Fechap=#" & DTser.Value & "#", Conexion
     If (BC.EOF = True) Then
         Conexion.Execute "insert into CamionesTemp(NumSol,placaB,Fechap,Peso,EstatusCP) values ('" & TXTnro.Text & "','" & c(0) & "','" & DTser.Value & "'," & TXTckg.Text & ", 'A')"
         MsgBox ("El camion fue apartado " & c(0))
         MsgBox "Asegurese de llenar el combo de placa con la placa asignada"
        CMDprc.Enabled = False
        DP = True
        CMDgua.Enabled = True
        Else
        MsgBox ("Tenemos un Camion para ese peso pero ya esta ocupado para esa fecha")
     End If
         BC.Close
    End If
    c.MoveNext
  Wend
  If DP Then

Else
    MsgBox "NO HAY CAMIONES DISPONIBLES QUE SOPORTEN ESE PESO", vbExclamation
    CMDpro.Enabled = False
    CMDsal.Enabled = True
    CMDprc.Enabled = True
    CMDgua.Enabled = False
   tcamdis.Open "select * from Camiones", Conexion
   CMBpla.Clear
   While tcamdis.EOF = False
     CMBpla.AddItem (tcamdis(0))
     tcamdis.MoveNext
   Wend
End If
End Sub

Private Sub CMBpro_Click()
Dim t As New ADODB.Recordset
Dim AUX As New ADODB.Recordset
        t.Open "select CodPro from TiposProducto where NomPro = '" & CMBpro.Text & "'", Conexion
        TXTpro.Text = t!CodPro
        TXTckg.SetFocus
AUX.Open "Select * from Tabulador Where (CodCiudadInicio = '" & TXTori.Text & "' and CodCiudadDestino = '" & TXTdes.Text & "')or(CodCiudadInicio = '" & TXTdes.Text & "' and CodCiudadDestino = '" & TXTori.Text & "')", Conexion
TXTBxC.Text = AUX(3)
AUX.Close
End Sub
Private Sub CMDbus_Click()
Dim ELNU As Integer
Dim t As New ADODB.Recordset
If TXTnro.Text = "" Then
  MsgBox " Debe Escribir el Numero de Servicio", vbInformation, "Omision de Datos"
  TXTnro.SetFocus
Else
    ELNU = Val(TXTnro.Text)
   t.Open "Select * from ServiciosEfectuados where NumSolic =" & ELNU & "", Conexion
 If t.EOF Then
  MsgBox " El Servicio No esta Registrado, Desea Incluirlo ahora?", vbExclamation, "Inclusion de Datos"
  A = TXTnro.Text
  CMDlim_Click
  TXTnro.Text = A
  CMDgua.Enabled = True
  TXTnro.SetFocus
  Else
  If t!EstatusSef = "E" Then
   If MsgBox(" El Servicio esta Inactivo, Desea Reactivarlo?", vbQuestion + vbYesNo, "Reactivacion de Datos") = vbYes Then
   activar
   MsgBox " Servicio Reactivado Exitosamente ", vbInformation, "Reactivacion de Datos"
   mostrar t
   End If
  ElseIf t!EstatusSef = "R" Then
   If MsgBox(" El Servicio ya se realizo, Desea Ver lo Datos?", vbQuestion + vbYesNo, "Reactivacion de Datos") = vbYes Then
   mostrar t
   End If
  Else
  CMDlim_Click
  TXTnro.SetFocus
  End If
 ' Else
  mostrar t
  End If
  End If
  End Sub

Private Sub CMDeli_Click()
If TXTnro.Text = "" Then
 MsgBox " Debe escribir el Numero de Servicio ", vbInformation, "Omision de Datos"
 TXTnro.SetFocus
 ElseIf CMBcli.Text = "" Then
  MsgBox " Debe Seleccionar el Codigo del Cliente ", vbInformation, "Omision de Datos"
  CMBcli.SetFocus
 ElseIf DTser.Value = "" Then
  MsgBox " Debe Escribir una Fecha deservicio ", vbExclamation, "Omision de Datos"
  DTser.SetFocus
 ElseIf CMBori.Text = "" Then
  MsgBox " Debe Seleccionar la Ciudad de Origen ", vbInformation, "Omision de Datos"
  CMBori.SetFocus
 ElseIf CMBdes.Text = "" Then
  MsgBox " Debe Seleccionar la Ciudad de Destino ", vbInformation, "Omision de Datos"
  CMBdes.SetFocus
 ElseIf CMBpro.Text = "" Then
  MsgBox " Debe Seleccionar un codigo de producto ", vbInformation, "Omision de Datos"
  CMBpro.SetFocus
 ElseIf TXTckg.Text = "" Then
  MsgBox " Debe escribir un Peso ", vbInformation, "Omision de Datos"
  TXTckg.SetFocus
 ElseIf DTsol.Value = "" Then
  MsgBox " Debe escribir una Fecha de solicitud", vbInformation, "Omision de Datos"
  DTsol.SetFocus
 ElseIf CMBpla.Text = "" Then
  MsgBox " Debe Seleccionar una Placa ", vbInformation, "Omision de Datos"
  CMBpla.SetFocus
 Else
    If (DTser.Value) < (DTglo.Value) Then
        MsgBox " No puedes Eliminar un Servicio que ya se efectuo ", vbCritical
    Else
        If MsgBox(" Desea Eliminar el Servicio?", vbQuestion + vbYesNo, "Eliminacion de Datos") = vbYes Then
        Conexion.Execute " update ServiciosEfectuados set EstatusSef = 'E' where NumSolic = " & TXTnro.Text & ""
        MsgBox " Servicio Eliminado Satisfactoriamente", vbInformation, "Eliminacion de Datos"
        CMDlim_Click
        TXTnro.SetFocus
  End If
End If
 End If
End Sub

Private Sub CMDgua_Click()
Dim t As New ADODB.Recordset
If TXTnro.Text = "" Then
 MsgBox " Debe escribir el Numero de Servicio ", vbInformation, "Omision de Datos"
 TXTnro.SetFocus
 ElseIf CMBcli.Text = "" Then
  MsgBox " Debe Seleccionar el Codigo del Cliente ", vbInformation, "Omision de Datos"
  CMBcli.SetFocus
 ElseIf DTser.Value = "" Then
  MsgBox " Debe Escribir una Fecha deservicio ", vbExclamation, "Omision de Datos"
  DTser.SetFocus
 ElseIf CMBori.Text = "" Then
  MsgBox " Debe Seleccionar la Ciudad de Origen ", vbInformation, "Omision de Datos"
  CMBori.SetFocus
 ElseIf CMBdes.Text = "" Then
  MsgBox " Debe Seleccionar la Ciudad de Destino ", vbInformation, "Omision de Datos"
  CMBdes.SetFocus
 ElseIf CMBpro.Text = "" Then
  MsgBox " Debe Seleccionar un codigo de producto ", vbInformation, "Omision de Datos"
  CMBpro.SetFocus
 ElseIf TXTckg.Text = "" Then
  MsgBox " Debe escribir un Peso ", vbInformation, "Omision de Datos"
  TXTckg.SetFocus
 ElseIf DTsol.Value = "" Then
  MsgBox " Debe escribir una Fecha de solicitud", vbInformation, "Omision de Datos"
  DTsol.SetFocus
 ElseIf CMBpla.Text = "" Then
  MsgBox " Debe Seleccionar una Placa ", vbInformation, "Omision de Datos"
  CMBpla.SetFocus
Else
If (DTser.Value) < (DTsol.Value) Then
MsgBox " La Fecha de servicio debe ser mayor que la de la solicitud", vbInformation, "Datos Incorrectos"
DTser.SetFocus
ElseIf DTsol.Value <> (DTglo.Value) Then
MsgBox "No puede Guardar ya que la fecha de solicitud no coincide con la de hoy", vbInformation
DTsol.SetFocus
Else
 Conexion.Execute "insert into ServiciosEfectuados (NumSolic,CodCliente,FecServicio,CodCiudOrig,CodCiudDest,CodProd,CantKg,FechaSolicitud,PlacaS,EstatusSef) values (" & TXTnro.Text & ", '" & CMBcli.Text & "', '" & DTser.Value & "', '" & TXTori.Text & "', '" & TXTdes.Text & "', '" & TXTpro.Text & "', '" & TXTckg.Text & "', '" & DTsol.Value & "', '" & CMBpla.Text & "', 'A')"
  MsgBox " Servicio Incluido Satisfactoriamente ", vbExclamation, "Inclusion de Datos"
  CMDlim_Click
  TXTnro.SetFocus
End If
End If

End Sub

Private Sub CMDlim_Click()
TXTnro.Text = ""
txtnom.Text = ""
DTser.Value = GLOBALDATE
TXTori.Text = ""
TXTdes.Text = ""
TXTpro.Text = ""
TXTckg.Text = ""
TXTprecio.Text = ""
DTsol.Value = GLOBALDATE
TXTpso.Text = ""
CMBpla.Text = ""
CMBcli.Text = ""
CMBori.Text = ""
CMBdes.Text = ""
CMBpro.Text = ""
TXTBxC.Text = ""
LBLcarga.Caption = ""
CMDgua.Enabled = False
CMDmod.Enabled = False
CMDeli.Enabled = False
CMDprc.Enabled = True
CMDpro.Enabled = True
End Sub
Private Sub CMDmod_Click()
If TXTnro.Text = "" Then
 MsgBox " Debe Generar el Numero de Servicio ", vbInformation, "Omision de Datos"
 TXTnro.SetFocus
 ElseIf CMBcli.Text = "" Then
  MsgBox " Debe Seleccionar el Codigo del Cliente ", vbInformation, "Omision de Datos"
  CMBcli.SetFocus
 ElseIf DTser.Value = "" Then
  MsgBox " Debe Seleccionar una Fecha deservicio ", vbExclamation, "Omision de Datos"
  DTser.SetFocus
 ElseIf CMBori.Text = "" Then
  MsgBox " Debe Seleccionar la Ciudad de Origen ", vbInformation, "Omision de Datos"
  CMBori.SetFocus
 ElseIf CMBdes.Text = "" Then
  MsgBox " Debe Seleccionar la Ciudad de Destino ", vbInformation, "Omision de Datos"
  CMBdes.SetFocus
 ElseIf CMBpro.Text = "" Then
  MsgBox " Debe Seleccionar un codigo de producto ", vbInformation, "Omision de Datos"
  CMBpro.SetFocus
 ElseIf TXTckg.Text = "" Then
  MsgBox " Debe escribir un Peso ", vbInformation, "Omision de Datos"
  TXTckg.SetFocus
 ElseIf DTsol.Value = "" Then
  MsgBox " Debe Seleccionar una Fecha de solicitud", vbInformation, "Omision de Datos"
  DTsol.SetFocus
 ElseIf CMBpla.Text = "" Then
  MsgBox " Debe Seleccionar una Placa ", vbInformation, "Omision de Datos"
  CMBpla.SetFocus
Else
 If (DTser.Value) < (DTglo.Value) Then
        MsgBox " No puedes Modificar un Servicio que ya se efectuo ", vbCritical
    Else
Conexion.Execute " update ServiciosEfectuados set CodCliente = '" & CMBcli.Text & "', FecServicio = '" & DTser.Value & "', CodCiudOrig = '" & TXTori.Text & "', CodCiudDest = '" & TXTdes.Text & "', CodProd = '" & TXTpro.Text & "', CantKg = '" & TXTckg.Text & "', FechaSolicitud = '" & DTsol.Value & "', PlacaS = '" & CMBpla.Text & "'  where NumSolic = " & TXTnro.Text & ""
  MsgBox " Repuesto Modificado Satisfactoriamente ", vbInformation, "Modificacion de Datos"
  CMDlim_Click
  TXTnro.SetFocus
End If
End If
End Sub

Private Sub CMDnue_Click()
Dim t As New ADODB.Recordset
    t.Open " select max(NumSolic) from ServiciosEfectuados", Conexion
If Not IsNumeric(t(0)) Then
    TXTnro.Text = "001"
Else
    TXTnro.Text = Format(Str(t(0)) + 1, "000")
End If
    CMBcli.SetFocus
End Sub
Private Sub CMDsal_Click()
If MsgBox(" Desea realmente salir de 'Servicios Efectuados'?", vbQuestion + vbYesNo, "Salida de Pantalla") = vbYes Then
Unload Me
End If
End Sub

Private Sub Form_Load()
DTglo.Value = GLOBALDATE
DTser.Value = GLOBALDATE
DTsol.Value = GLOBALDATE
LlenarCombo CMBcli, "select * from Clientes", "CodCli"
LlenarCombo CMBori, "select * from Ciudades", "Ciudades"
LlenarCombo CMBdes, "select * from Ciudades", "Ciudades"
LlenarCombo CMBpro, "select * from TiposProducto", "NomPro"
LlenarCombo CMBpla, "select * from Camiones", "Placa"
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
CMDlim_Click
End Sub
Sub activar()
Conexion.Execute " update ServiciosEfectuados set EstatusSef = 'A' where NumSolic = " & TXTnro.Text & ""
End Sub
Private Sub mostrar(t As ADODB.Recordset)
Dim AUX As New ADODB.Recordset
Dim valor As New ADODB.Recordset
TXTnro = t!NumSolic
CMBcli.Text = t!CodCliente
DTser.Value = t!FecServicio
TXTori.Text = t!CodCiudOrig
TXTdes.Text = t!CodCiudDest
TXTpro.Text = t!CodProd
TXTckg.Text = t!CantKg
DTsol.Value = t!FechaSolicitud
TXTpso.Visible = False
CMBpla.Text = t!placas
CMDgua.Enabled = False
CMDmod.Enabled = True
CMDeli.Enabled = True
valor.Open " SELECT * FROM CIUDADES Where CodCiudad ='" & TXTori.Text & "'", Conexion
CMBori.Text = valor(1)
valor.Close
valor.Open " SELECT * FROM CIUDADES Where CodCiudad ='" & TXTdes.Text & "'", Conexion
CMBdes.Text = valor!Ciudades
valor.Close
valor.Open " select * from Clientes where CodCli = '" & CMBcli.Text & "'", Conexion
txtnom.Text = valor(1)
valor.Close
valor.Open " Select * from TiposProducto where CodPro = '" & TXTpro.Text & "'", Conexion
CMBpro.Text = valor(1)
valor.Close
AUX.Open "Select * from Tabulador Where CodCiudadInicio = '" & TXTori.Text & "' and CodCiudadDestino = '" & TXTdes.Text & "'", Conexion
TXTBxC.Text = AUX(3)
AUX.Close
End Sub
Private Sub TXTBxC_Change()

If TXTckg.Text <> "" Then
TXTprecio.Text = Val(TXTckg.Text) * Val(TXTBxC.Text)
End If
End Sub
Private Sub TXTckg_GotFocus()
TXTprecio.Text = Val(TXTckg.Text) * Val(TXTBxC.Text)
End Sub
Private Sub TXTckg_KeyPress(KeyAscii As Integer)
Dim t As New ADODB.Recordset
If KeyAscii = 13 Then
If IsNumeric(TXTckg.Text) = False Then
    MsgBox "Este Campo es Numerico, Cambie la Condición", vbExclamation, "Dato Incorrecto"
    TXTckg.Text = ""
    TXTckg.SetFocus
    ElseIf Val(TXTckg.Text) < 0 Then
MsgBox " El Peso no puede ser Negativo, Cambie la Condición", vbExclamation, "Dato Incorrecto"
TXTckg.Text = ""
TXTckg.SetFocus
Else
t.Open "Select * from Tabulador Where (CodCiudadInicio = '" & TXTori.Text & "' and CodCiudadDestino = '" & TXTdes.Text & "')or(CodCiudadInicio = '" & TXTdes.Text & "' and CodCiudadDestino = '" & TXTori.Text & "')", Conexion
TXTprecio.Text = Val(TXTckg.Text) * t(3)
t.Close
t.Open "select CodCiudadInicio,CodCiudadDestino,TiempoEstimDias,BsPorKgTransp,EstatusTab  from Tabulador where CodCiudadInicio  = '" & TXTori.Text & "' and CodCiudadDestino = '" & TXTdes.Text & "'", Conexion
If t.EOF Then
t.Close
t.Open "Select CodCiudadInicio,CodCiudadDestino,TiempoEstimDias,BsPorKgTransp,EstatusTab from Tabulador where CodCiudadInicio = '" & TXTdes.Text & "' and CodCiudadDestino = '" & TXTori.Text & "'", Conexion
If t.EOF Then
MsgBox "Imposible Calcular el Precio de la Transaccion, ya que no Creo este Tabulador", vbExclamation, "No Puede Proceder"
   CMBori.Text = ""
   CMBdes.Text = ""
   TXTdes.Text = ""
   TXTori.Text = ""
   CMBori.SetFocus
End If
End If
End If
End If
End Sub

Sub BuscarTab()
Dim t As New ADODB.Recordset
If CMBori.Text = "" Then
  MsgBox " Debe Elegir la Ciudad de Origen", vbInformation, "Omision de Datos"
  CMBori.SetFocus
  ElseIf CMBdes.Text = "" Then
    MsgBox " Debe Elegir la Ciudad de Destino distinta", vbInformation, "Omision de Datos"
    CMBdes.SetFocus
Else
t.Open "select CodCiudadInicio,CodCiudadDestino,TiempoEstimDias,BsPorKgTransp,EstatusTab  from Tabulador where CodCiudadInicio  = '" & TXTori.Text & "' and CodCiudadDestino = '" & TXTdes.Text & "'", Conexion
If t.EOF Then
t.Close
t.Open "Select CodCiudadInicio,CodCiudadDestino,TiempoEstimDias,BsPorKgTransp,EstatusTab from Tabulador where CodCiudadInicio = '" & TXTdes.Text & "' and CodCiudadDestino = '" & TXTori.Text & "'", Conexion
 If t.EOF Then
  If MsgBox(" Este Tabulador No Existe, Desea Crearlo Ahora?", vbQuestion + vbYesNo, "Inclusion de Datos") = vbYes Then
    Tabulador.Show
    Tabulador.CMBori.Text = CMBori.Text
    Tabulador.CMBdes.Text = CMBdes.Text
    Tabulador.TXTdes.Text = TXTdes.Text
    Tabulador.TXTori.Text = TXTori.Text
    Tabulador.Buscar
    Tabulador.TXTdia.SetFocus
Else
MsgBox "Al No crear el Tabulador no puede Registrar este Servicio, Revise Bien su Información y Vuelva a Intentarlo", vbExclamation, "No Puede Proceder"
    CMBori.Text = ""
    CMBdes.Text = ""
    TXTdes.Text = ""
    TXTori.Text = ""
    CMBori.SetFocus
    End If
End If
 Else
    If t!EstatusTab <> "A" Then
        If MsgBox(" Este Tabulador esta actualmente Inactivo, Desea Reactivarlo?", vbQuestion + vbYesNo, "Reactivacion de Datos") = vbYes Then
            activar
            MsgBox " Registro Reactivado Satisfactoriamente ", vbInformation, "Reactivación de Datos"
    Conexion.Execute "update Tabulador set EstatusTab = 'A' where (CodCiudadInicio = '" & TXTori.Text & "' and CodCiudadDestino = '" & TXTdes.Text & "')or(CodCiudadInicio = '" & TXTdes.Text & "' and CodCiudadDestino = '" & TXTori.Text & "')and EstatusTab = 'E'"
    Tabulador.CMBori.Text = CMBori.Text
    Tabulador.CMBdes.Text = CMBdes.Text
    Tabulador.TXTdes.Text = TXTdes.Text
    Tabulador.TXTori.Text = TXTori.Text
        Else
            CMBori.Text = ""
    CMBdes.Text = ""
    TXTori.Text = ""
    TXTdes.Text = ""
            CMBori.SetFocus
        End If
    Else
        Tabulador.CMBori.Text = CMBori.Text
    Tabulador.CMBdes.Text = CMBdes.Text
    Tabulador.TXTdes.Text = TXTdes.Text
    Tabulador.TXTori.Text = TXTori.Text
    End If
 End If
End If
End Sub
Private Sub TXTser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsDate(DTser.Value) = False Then
    MsgBox "Este campo es de tipo fecha, Cambie la Condicion", vbExclamation, "Dato Incorrecto"
    DTser.Value = ""
    DTser.SetFocus
Else
 DTsol.SetFocus
End If
End If
End Sub


Private Sub TXTnro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(TXTnro.Text) = False Then
    MsgBox "Este Campo es Numerico, Cambie la Condición", vbExclamation, "Dato Incorrecto"
    TXTnro.Text = ""
    TXTnro.SetFocus
    ElseIf Val(TXTnro.Text) < 0 Then
MsgBox " EL Codigo del Servicio no puede ser Negativo, Cambie la Condición", vbExclamation, "Dato Incorrecto"
TXTnro.Text = ""
TXTnro.SetFocus
End If
End If
End Sub

