VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Factura 
   BackColor       =   &H80000005&
   Caption         =   "    "
   ClientHeight    =   10665
   ClientLeft      =   2385
   ClientTop       =   990
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Factura.frx":0000
   ScaleHeight     =   10665
   ScaleWidth      =   11940
   Begin VB.CommandButton CMDbus 
      Height          =   495
      Left            =   6360
      Picture         =   "Factura.frx":55AE
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox TXTfac 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   23
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox TXTcod 
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
      Left            =   3360
      TabIndex        =   22
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox TXTori 
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
      Left            =   5880
      MaxLength       =   8
      TabIndex        =   21
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox TXTdes 
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
      Left            =   5880
      MaxLength       =   8
      TabIndex        =   20
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox TXTpro 
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
      Left            =   5880
      MaxLength       =   8
      TabIndex        =   19
      Top             =   7200
      Width           =   855
   End
   Begin VB.TextBox TXTckg 
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
      Left            =   7680
      MaxLength       =   3
      TabIndex        =   18
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox TXTnom 
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
      Left            =   3360
      MaxLength       =   30
      TabIndex        =   17
      Top             =   3960
      Width           =   2895
   End
   Begin VB.TextBox txtpla 
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
      Left            =   3600
      TabIndex        =   16
      Top             =   7680
      Width           =   2055
   End
   Begin VB.TextBox TXTtot 
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
      Height          =   615
      Left            =   5640
      TabIndex        =   15
      Top             =   8400
      Width           =   1695
   End
   Begin VB.TextBox TXTsol 
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
      Left            =   4080
      TabIndex        =   14
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox TXTser 
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
      Left            =   1080
      TabIndex        =   13
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox TXTfev 
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
      Left            =   6960
      TabIndex        =   12
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox TXTndes 
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
      Left            =   3600
      TabIndex        =   11
      Top             =   6720
      Width           =   2055
   End
   Begin VB.TextBox TXTnpro 
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
      Left            =   3600
      TabIndex        =   10
      Top             =   7200
      Width           =   2055
   End
   Begin VB.TextBox TXTnori 
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
      Left            =   3600
      TabIndex        =   9
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox txtfecha 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox TXTsub 
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
      Left            =   2280
      MaxLength       =   8
      TabIndex        =   7
      Top             =   8400
      Width           =   1095
   End
   Begin VB.TextBox TXTiva 
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
      Left            =   2280
      MaxLength       =   8
      TabIndex        =   6
      Top             =   8880
      Width           =   1095
   End
   Begin VB.TextBox TXTpre 
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton CMDimp 
      Caption         =   "&Imprimir"
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
      Left            =   2880
      Picture         =   "Factura.frx":5B38
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9720
      Width           =   1815
   End
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
      Left            =   6720
      Picture         =   "Factura.frx":60C2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9720
      Width           =   1815
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
      Left            =   4800
      Picture         =   "Factura.frx":664C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CommandButton CMDgua 
      Caption         =   "&Cobrar"
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
      Left            =   960
      Picture         =   "Factura.frx":6BD6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9720
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   9720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LBLkg 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "BsF/Kg"
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
      Index           =   4
      Left            =   6840
      TabIndex        =   44
      Top             =   7080
      Width           =   765
   End
   Begin VB.Label LBLkg 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   3
      Left            =   7440
      TabIndex        =   42
      Top             =   8760
      Width           =   390
   End
   Begin VB.Shape Shape2 
      Height          =   2175
      Left            =   960
      Top             =   6000
      Width           =   8415
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   960
      Top             =   3360
      Width           =   8295
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Datos del Servicio"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   41
      Top             =   5640
      Width           =   2220
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Nombre"
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
      Left            =   1200
      TabIndex        =   40
      Top             =   3960
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Datos del Cliente"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   39
      Top             =   3000
      Width           =   2100
   End
   Begin VB.Label LBLnro 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Nro. de Factura"
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
      Left            =   240
      TabIndex        =   38
      Top             =   2520
      Width           =   1620
   End
   Begin VB.Label Label1 
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
      TabIndex        =   37
      Top             =   7800
      Width           =   1845
   End
   Begin VB.Label LBLven 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fecha Vencimiento"
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
      Left            =   6960
      TabIndex        =   36
      Top             =   4680
      Width           =   1995
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
      TabIndex        =   35
      Top             =   6240
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
      TabIndex        =   34
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label LBLcod 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Codigo del Producto"
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
      TabIndex        =   33
      Top             =   7200
      Width           =   2145
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
      Left            =   6960
      TabIndex        =   32
      Top             =   6480
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
      Left            =   9000
      TabIndex        =   31
      Top             =   6480
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
      Left            =   4080
      TabIndex        =   30
      Top             =   4680
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
      Left            =   1080
      TabIndex        =   29
      Top             =   4680
      Width           =   1575
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
      Left            =   1200
      TabIndex        =   28
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label LBLkg 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   4680
      TabIndex        =   27
      Top             =   8640
      Width           =   705
   End
   Begin VB.Label LBLsol 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fecha "
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
      Left            =   6840
      TabIndex        =   26
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "I.V.A."
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
      Left            =   1080
      TabIndex        =   25
      Top             =   9000
      Width           =   555
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Sub Total"
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
      Left            =   1080
      TabIndex        =   24
      Top             =   8400
      Width           =   1020
   End
   Begin VB.Label LBLfac 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Facturas"
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
      Left            =   8520
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "Factura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cont As Integer
Private Sub Command1_Click()
Imprimir.ShowPrinter
End Sub

Private Sub CMDbus_Click()
Dim t As New ADODB.Recordset
If TXTCod.Text = "" Then
    MsgBox "Debe ingresar El Codigo del Cliente", vbInformation, "Omision de Datos"
Else
    t.Open "select * from Facturas where CodClientef = '" & TXTCod.Text & "' and fecvencim = (select min(fecvencim) from facturas where codclientef = '" & TXTCod.Text & "') ", Conexion
    If t.EOF Then
        MsgBox " El Cliente No Esta Registrado O No Tiene Facturas Pendientes", vbInformation, "Inclusion de Datos"
        CMDlim_Click
    Else
    Dim tf As New ADODB.Recordset
    tf.Open "select numfact from Facturas where codclientef = '" & TXTCod.Text & "'", Conexion
     cont = 0
While Not tf.EOF
    cont = cont + 1
    tf.MoveNext
 Wend
    MsgBox "El Cliente debe: " & cont & " Facturas "
    mostrar t
    TXTfac.Enabled = False
    TXTsol.Enabled = False
    TXTser.Enabled = False
    txtpla.Enabled = False
    txtnom.Enabled = False
    TXTfev.Enabled = False
    TXTnori.Enabled = False
    TXTori.Enabled = False
    TXTndes.Enabled = False
    TXTdes.Enabled = False
    TXTnpro.Enabled = False
    TXTpro.Enabled = False
    TXTckg.Enabled = False
    TXTpre.Enabled = False
    CMDgua.Enabled = True
    CMDlim.Enabled = True
    End If
    End If
    TXTCod.SetFocus

End Sub

Private Sub CMDimp_Click()
CommonDialog1.ShowPrinter
End Sub

Private Sub Form_Initialize()
txtfecha.Text = GLOBALDATE
TXTCod.Enabled = True

End Sub


Private Sub CMDlim_Click()
    TXTfac.Text = ""
    TXTsol.Text = ""
    TXTser.Text = ""
    txtpla.Text = ""
    TXTCod.Text = ""
    txtnom.Text = ""
    TXTfev.Text = ""
    TXTnori.Text = ""
    TXTori.Text = ""
    TXTndes.Text = ""
    TXTdes.Text = ""
    TXTnpro.Text = ""
    TXTpro.Text = ""
    TXTckg.Text = ""
    TXTtot.Text = ""
End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Private Sub CMDgua_Click()
Dim t As New ADODB.Recordset
If cont >= 2 Then
MsgBox ("Se Va A Cobrar La Factura Mas Antigua ")
     If MsgBox("Desea Cancelar Esta Factura?", vbQuestion + vbYesNo, "Omision de Datos") = vbYes Then
            Conexion.Execute "update Facturas set Estatusfac = 'E' where numfact = '" & TXTfac.Text & "' "
            Conexion.Execute "insert into FactCobradas (Numfaco,CodCli,FecVencim,FecCobro,Placa,estatusFaco) values('" & TXTfac.Text & "', '" & TXTCod.Text & "', '" & TXTfev.Text & "','" & GLOBALDATE & "', '" & txtpla.Text & "','C')"
            MsgBox " Factura Cancelada ", vbInformation, "Atencion"
            CMDlim_Click
            
            Conexion.Execute "delete * from Facturas where Estatusfac = 'E'"
            End If
    Else
    If MsgBox("Desea Cancelar Esta Factura?", vbQuestion + vbYesNo, "Omision de Datos") = vbYes Then
            Conexion.Execute "update Facturas set Estatusfac = 'E' where numfact = '" & TXTfac.Text & "' "
            Conexion.Execute "insert into FactCobradas (Numfaco,CodCli,FecVencim,FecCobro,Placa,estatusFaco) values('" & TXTfac.Text & "', '" & TXTCod.Text & "', '" & TXTfev.Text & "','" & GLOBALDATE & "', '" & txtpla.Text & "','C')"
            MsgBox " Factura Cancelada ", vbInformation, "Atencion"
            CMDlim_Click
            
            Conexion.Execute "delete * from Facturas where Estatusfac = 'E'"
            End If
        
     CMDlim_Click
     End If

End Sub

Private Sub CMDsal_Click()
If MsgBox("Desea Realmente Salir de 'Facturas'?", vbQuet + vbYesNo, "Salida de Pantalla") = vbYes Then
Unload Me
End If
End Sub

Sub mostrar(t As ADODB.Recordset)

Dim AUX As New ADODB.Recordset
TXTfac.Text = t!NumFact
TXTfev.Text = t!FecVencim
TXTori.Text = t!CodCiudOrig
TXTdes.Text = t!CodCiudDest
TXTser.Text = t!FechaServicio
TXTsol.Text = t!FechaSolicitud
TXTpro.Text = t!CodProd
TXTckg.Text = t!CantKg
txtpla.Text = t!placa
TXTiva.Text = 12
AUX.Open "select * from Clientes where CodCli = '" & TXTCod.Text & "'", Conexion
txtnom.Text = AUX(1)
AUX.Close

AUX.Open "select * from ciudades where CodCiudad = '" & TXTori.Text & "'", Conexion
TXTnori.Text = AUX(1)
AUX.Close

AUX.Open "select * from ciudades where CodCiudad = '" & TXTdes.Text & "'", Conexion
TXTndes.Text = AUX(1)
AUX.Close

AUX.Open "select * from TiposProducto where Codpro = '" & TXTpro.Text & "'", Conexion
TXTnpro.Text = AUX(1)
AUX.Close

AUX.Open "Select * from Tabulador Where CodCiudadInicio = '" & TXTori.Text & "' and CodCiudadDestino = '" & TXTdes.Text & "'", Conexion
TXTpre.Text = AUX(3)
AUX.Close
End Sub

Private Sub TXTCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim t As New ADODB.Recordset
If TXTCod.Text = "" Then
    MsgBox "Debe ingresar El Codigo del Cliente", vbInformation, "Omision de Datos"
Else
    t.Open "select * from Facturas where CodClientef = '" & TXTCod.Text & "' and fecvencim = (select min(fecvencim) from facturas where codclientef = '" & TXTCod.Text & "') ", Conexion
    If t.EOF Then
        MsgBox " El Cliente No Esta Registrado O No Tiene Facturas Pendientes", vbInformation, "Inclusion de Datos"
        CMDlim_Click
    Else
    Dim tf As New ADODB.Recordset
    tf.Open "select numfact from Facturas where codclientef = '" & TXTCod.Text & "'", Conexion
     cont = 0
While Not tf.EOF
    cont = cont + 1
    tf.MoveNext
 Wend
    MsgBox "El Cliente debe: " & cont & " Facturas "
    mostrar t
    TXTfac.Enabled = False
    TXTsol.Enabled = False
    TXTser.Enabled = False
    txtpla.Enabled = False
    txtnom.Enabled = False
    TXTfev.Enabled = False
    TXTnori.Enabled = False
    TXTori.Enabled = False
    TXTndes.Enabled = False
    TXTdes.Enabled = False
    TXTnpro.Enabled = False
    TXTpro.Enabled = False
    TXTckg.Enabled = False
    TXTpre.Enabled = False
    CMDgua.Enabled = True
    CMDlim.Enabled = True
    End If
    End If
    End If
    TXTCod.SetFocus

End Sub

Private Sub TXTpre_Change()
If TXTckg.Text <> "" Then
     TXTsub.Text = Val(TXTckg.Text) * Val(TXTpre.Text)
     TXTiva.Text = Val(TXTsub.Text) * 0.12
     TXTtot.Text = Val(TXTsub.Text) + Val(TXTiva.Text)
  End If
End Sub

Private Sub TXTtot_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub



