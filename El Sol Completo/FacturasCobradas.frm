VERSION 5.00
Begin VB.Form FacturasCobradas 
   BackColor       =   &H80000005&
   Caption         =   "Transportes EL SOL // Facturas Cobradas"
   ClientHeight    =   8775
   ClientLeft      =   2385
   ClientTop       =   1125
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FacturasCobradas.frx":0000
   ScaleHeight     =   8775
   ScaleWidth      =   10905
   Begin VB.ComboBox Cmbfc 
      Height          =   315
      Left            =   2400
      TabIndex        =   16
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton CMDRepor 
      Caption         =   "&Reporte"
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
      Left            =   3600
      Picture         =   "FacturasCobradas.frx":55AE
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox TXTcli 
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
      Left            =   3720
      MaxLength       =   8
      TabIndex        =   13
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox TXTplac 
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
      Left            =   3720
      MaxLength       =   8
      TabIndex        =   12
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton CMDsal 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   1320
      Picture         =   "FacturasCobradas.frx":5B38
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox TXTcob 
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
      Left            =   3720
      MaxLength       =   8
      TabIndex        =   8
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox TXTven 
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
      Left            =   3720
      MaxLength       =   8
      TabIndex        =   6
      Top             =   4440
      Width           =   2055
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
      Left            =   3720
      MaxLength       =   30
      TabIndex        =   4
      Top             =   3840
      Width           =   3015
   End
   Begin VB.CommandButton CMDbus 
      Height          =   495
      Left            =   3600
      Picture         =   "FacturasCobradas.frx":60C2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton CMDlim 
      Caption         =   "&Limpiar"
      Height          =   495
      Left            =   4200
      Picture         =   "FacturasCobradas.frx":664C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Nombre del Cliente"
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
      TabIndex        =   14
      Top             =   3960
      Width           =   2010
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
      TabIndex        =   10
      Top             =   5760
      Width           =   1845
   End
   Begin VB.Label LBLcob 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fecha Cobro"
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
      Left            =   1320
      TabIndex        =   9
      Top             =   5160
      Width           =   1350
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
      Left            =   1320
      TabIndex        =   7
      Top             =   4560
      Width           =   1995
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
      TabIndex        =   5
      Top             =   3360
      Width           =   1935
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
      Left            =   600
      TabIndex        =   1
      Top             =   2640
      Width           =   1620
   End
   Begin VB.Label LBLfactcobrada 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Facturas Cobradas"
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
      Left            =   6720
      TabIndex        =   0
      Top             =   720
      Width           =   3930
   End
End
Attribute VB_Name = "FacturasCobradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDbus_Click()
Buscar
End Sub

Private Sub CMDlim_Click()
Limpiar
End Sub

Private Sub CMDRepor_Click()
RepFacCobRadas.Show
End Sub

Private Sub CMDsal_Click()
If MsgBox("Desea Realmente Salir de 'Facturas Cobradas'?", vbQuestion + vbYesNo, "Salida de Pantalla") = vbYes Then
Unload Me
End If
End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
LlenarCombo Cmbfc, "select * from FactCobradas", "numfaco"
End Sub

Sub Limpiar()
Cmbfc.Text = ""
TXTcli.Text = ""
txtnom.Text = ""
TXTcob.Text = ""
TXTplac.Text = ""
TXTven.Text = ""
Cmbfc.SetFocus
End Sub

Sub Buscar()
Dim t As New ADODB.Recordset
If Cmbfc.Text = "" Then
    MsgBox " Debe Ingresar el Numero de la Factura", vbInformation, " Omision De Datos"
    Cmbfc.SetFocus
Else
    t.Open "select * from FactCobradas where NumFaco = '" & Cmbfc.Text & "'", Conexion
    If t.EOF Then
        MsgBox "La Factura Solicitada No Existe, Revise la Información", vbExclamation, "Dato Inexistente"
        Limpiar
        Cmbfc.SetFocus
    Else
        mostrar t
        
        TXTcli.Enabled = False
        TXTven.Enabled = False
        TXTcob.Enabled = False
        TXTplac.Enabled = False
    End If
End If
End Sub

Sub mostrar(t As ADODB.Recordset)
Dim tp As New ADODB.Recordset
    TXTcli.Text = t!CodCli
    TXTven.Text = t!FecVencim
    TXTcob.Text = t!Feccobro
    TXTplac.Text = t!placa
    tp.Open "Select * from Clientes where CodCli = '" & TXTcli.Text & "' and EstatusCli = 'A'", Conexion
    txtnom.Text = tp!nomcli
    tp.Close
    't.Open "Select * from Clientes where CodCli = '" & TXTcli.Text & "' and EstatusCli = 'A'", Conexion
    'TXTtcli = t!tipo
End Sub

Private Sub TXTfac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsNumeric(TXTfac) = False Then
        MsgBox "Este Campo es Numerico, Cambie la Condición", vbExclamation, "Dato Incorrecto"
        TXTfac.Text = ""
        TXTfac.SetFocus
        ElseIf Val(TXTfac.Text) < 0 Then
            MsgBox "Este Numero de Factura es Incorrecto, Cambie la Condición", vbExclamation, "Dato Incorrecto"
            TXTfac.Text = ""
            TXTfac.SetFocus
    End If
Buscar
End If
End Sub
