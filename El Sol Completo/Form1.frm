VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Principal 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Transporte EL SOL // Bienvenido"
   ClientHeight    =   8595
   ClientLeft      =   2235
   ClientTop       =   885
   ClientWidth     =   10530
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":0000
   Picture         =   "Form1.frx":030A
   ScaleHeight     =   8595
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Height          =   2415
      Left            =   8040
      Picture         =   "Form1.frx":F919
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   3840
      Picture         =   "Form1.frx":106CD
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   2160
      Picture         =   "Form1.frx":10D34
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   840
      Picture         =   "Form1.frx":11446
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   7560
      Picture         =   "Form1.frx":11ABE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2400
      Top             =   7200
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
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
      Enabled         =   0
      Connect         =   $"Form1.frx":1230D
      OLEDBString     =   $"Form1.frx":123B0
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
   Begin VB.Image Image4 
      Height          =   390
      Left            =   -120
      Picture         =   "Form1.frx":12453
      Top             =   7920
      Width           =   12030
   End
   Begin VB.Menu MNclientes 
      Caption         =   "Clientes"
      Begin VB.Menu OPclientes 
         Caption         =   "Clientes"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu MNservicios 
      Caption         =   "Servicios"
      Begin VB.Menu OPserviciosefec 
         Caption         =   "Servicios Efectuados"
         Shortcut        =   ^E
      End
      Begin VB.Menu TpPro 
         Caption         =   "Tipos Producto"
         Shortcut        =   ^P
      End
      Begin VB.Menu OPtabulador 
         Caption         =   "Tabulador"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu MNcamiones 
      Caption         =   "Camiones"
      Begin VB.Menu OPcamiones 
         Caption         =   "Camiones"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu MNmantenimiento 
      Caption         =   "Mantenimiento"
      Begin VB.Menu OPmantenimientoxcamion 
         Caption         =   "Mantenimiento por Camión"
         Shortcut        =   ^M
      End
      Begin VB.Menu OPmantenimientoreq 
         Caption         =   "Mantenimientos Requeridos"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu MNrepuestos 
      Caption         =   "Repuestos"
      Begin VB.Menu OPrepuestos 
         Caption         =   "Repuestos"
         Shortcut        =   {F3}
      End
      Begin VB.Menu OPrepuestosxcamion 
         Caption         =   "Repuestos por Camión"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu MNfacturacion 
      Caption         =   "Facturación"
      Begin VB.Menu OPfacturas 
         Caption         =   "Facturas"
         Shortcut        =   {F4}
      End
      Begin VB.Menu OPfacturascobradas 
         Caption         =   "Facturas Cobradas"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu MNreportes 
      Caption         =   "Reportes"
      Begin VB.Menu OPreporteMorosos 
         Caption         =   "Reporte Clientes Morosos"
         Shortcut        =   {F5}
      End
      Begin VB.Menu OPreportePend 
         Caption         =   "Reporte Servicios Pendientes"
         Shortcut        =   {F6}
      End
      Begin VB.Menu OpVenFac 
         Caption         =   "Reporte Vencimiento De Facturas"
         Shortcut        =   {F7}
      End
      Begin VB.Menu OpMan 
         Caption         =   "Reporte Mantenimiento Por Camion"
         Shortcut        =   {F8}
      End
      Begin VB.Menu OpRep 
         Caption         =   "Reporte Repuesto Por Camion"
         Shortcut        =   {F9}
      End
      Begin VB.Menu Opfactc 
         Caption         =   "Reporte Facturas Cobradas"
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu MNayuda 
      Caption         =   "Ayuda"
   End
   Begin VB.Menu MNacerca 
      Caption         =   "Acerca de..."
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If MsgBox("Desea Realmente Salir de 'Transporte EL SOL Todo Terreno!'?", vbQuestion + vbYesNo, "Salida de Programa") = vbYes Then
Unload Me
End If
End Sub

Private Sub Command2_Click()
Mision.Show
End Sub

Private Sub Command3_Click()
Objetivos.Show
End Sub

Private Sub Command4_Click()
Vision.Show
End Sub

Private Sub Command5_Click()
Acercade.Show
End Sub

Private Sub Command6_Click()
'Inicial2.Show

End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Private Sub MNacerca_Click()
Acercade.Show
End Sub

Private Sub MNayuda_Click()
Ayuda.Show
End Sub

Private Sub OPcamiones_Click()
Camiones.Show
End Sub

Private Sub OPclientes_Click()
Clientes.Show
End Sub

Private Sub OPmantenimiento_Click()

End Sub

Private Sub Opfactc_Click()
RepFacCobRadas.Show
End Sub

Private Sub OPfacturas_Click()
Factura.Show
End Sub

Private Sub OPfacturascobradas_Click()
FacturasCobradas.Show
End Sub

Private Sub OpMan_Click()
ReporteMantReq.Show
End Sub

Private Sub OPmantenimientoreq_Click()
Mantreq.Show
End Sub

Private Sub OPmantenimientoxcamion_Click()
Mantporcamion.Show
End Sub

Private Sub OPreportediario_Click()
ReporteDiario.Show
End Sub

Private Sub OpRep_Click()
ReporteRepCaMion.Show
End Sub

Private Sub OPreporteMorosos_Click()
ReporteCliente.Show
End Sub

Private Sub OPreportePend_Click()
ReporteServReq.Show
End Sub

Private Sub OPrepuestos_Click()
Repuestos.Show
End Sub

Private Sub OPrepuestosxcamion_Click()
Repuestoporcam.Show
End Sub

Private Sub OPservicioscontratados_Click()
ServiciosContratos.Show
End Sub

Private Sub OPserviciosefec_Click()
ServicioEfectuado.Show
End Sub

Private Sub OPtabulador_Click()
Tabulador.Show
End Sub

Private Sub OpVenFac_Click()
ReporteFacturas.Show
End Sub

Private Sub TpPro_Click()
Productos.Show
End Sub
