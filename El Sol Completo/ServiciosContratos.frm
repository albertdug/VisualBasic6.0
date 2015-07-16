VERSION 5.00
Begin VB.Form ServiciosContratos 
   BackColor       =   &H80000005&
   Caption         =   "Transporte EL SOL // Servicios Contratos"
   ClientHeight    =   8685
   ClientLeft      =   2130
   ClientTop       =   870
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   Picture         =   "ServiciosContratos.frx":0000
   ScaleHeight     =   8685
   ScaleWidth      =   10800
   Begin VB.CommandButton CMDlim 
      Caption         =   "&Limpiar"
      Height          =   375
      Left            =   4920
      TabIndex        =   27
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton CMDbus 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   3480
      TabIndex        =   26
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton CMDsal 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   7200
      TabIndex        =   25
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton CMDeli 
      Caption         =   "&Eliminar"
      Height          =   495
      Left            =   5280
      TabIndex        =   24
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton CMDmod 
      Caption         =   "&Modificar"
      Height          =   495
      Left            =   3360
      TabIndex        =   23
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton CMDgua 
      Caption         =   "&Guardar"
      Height          =   495
      Left            =   1440
      TabIndex        =   22
      Top             =   7680
      Width           =   1815
   End
   Begin VB.TextBox TXTser 
      Height          =   375
      Left            =   3840
      MaxLength       =   8
      TabIndex        =   21
      Top             =   6960
      Width           =   2055
   End
   Begin VB.TextBox TXTsol 
      Height          =   375
      Left            =   3840
      MaxLength       =   8
      TabIndex        =   19
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox TXTckg 
      Height          =   375
      Left            =   3840
      MaxLength       =   3
      TabIndex        =   15
      Top             =   5520
      Width           =   1335
   End
   Begin VB.ComboBox CMBpro 
      Height          =   315
      Left            =   3840
      TabIndex        =   14
      Top             =   4920
      Width           =   2055
   End
   Begin VB.TextBox TXTpro 
      Height          =   375
      Left            =   6000
      MaxLength       =   8
      TabIndex        =   13
      Top             =   4920
      Width           =   855
   End
   Begin VB.ComboBox CMBdes 
      Height          =   315
      Left            =   3840
      TabIndex        =   11
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox TXTdes 
      Height          =   375
      Left            =   6000
      MaxLength       =   8
      TabIndex        =   10
      Top             =   4320
      Width           =   855
   End
   Begin VB.ComboBox CMBori 
      Height          =   315
      Left            =   3840
      TabIndex        =   8
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox TXTori 
      Height          =   375
      Left            =   6000
      MaxLength       =   8
      TabIndex        =   7
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox TXTnom 
      Height          =   375
      Left            =   6240
      MaxLength       =   30
      TabIndex        =   5
      Top             =   3120
      Width           =   2895
   End
   Begin VB.ComboBox CMBcli 
      Height          =   315
      Left            =   3840
      TabIndex        =   4
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox TXTnro 
      Height          =   375
      Left            =   2400
      MaxLength       =   3
      TabIndex        =   1
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Contratados"
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
      Left            =   7800
      TabIndex        =   28
      Top             =   960
      Width           =   2565
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
      Left            =   1440
      TabIndex        =   20
      Top             =   6960
      Width           =   1575
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
      Left            =   1440
      TabIndex        =   18
      Top             =   6240
      Width           =   1620
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
      Left            =   5280
      TabIndex        =   17
      Top             =   5640
      Width           =   300
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
      Left            =   1440
      TabIndex        =   16
      Top             =   5520
      Width           =   555
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
      Left            =   1440
      TabIndex        =   12
      Top             =   4920
      Width           =   2145
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
      Left            =   1440
      TabIndex        =   9
      Top             =   4320
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
      Left            =   1440
      TabIndex        =   6
      Top             =   3720
      Width           =   1830
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
      Left            =   1440
      TabIndex        =   3
      Top             =   3120
      Width           =   1935
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
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   1740
   End
   Begin VB.Label LBLservicioscon 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Servicios"
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
      Top             =   480
      Width           =   1830
   End
End
Attribute VB_Name = "ServiciosContratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDbus_Click()
Dim t As New ADODB.Recordset
If TXTnro.Text = "" Then
  MsgBox " Debe Escribir el Numero de Solicitud", vbInformation
  TXTnro.SetFocus
Else
   TXTnro.ForeColor = &HFF&
  t.Open "Select * from ServiciosContratados where NumSolic = '" & TXTnro.Text & "'", Conexion
 If t.EOF Then
  MsgBox " La Solicitud No esta Registrada, Desea Incluirla ahora?", vbExclamation
  a = TXTnro.Text
  CMDlim_Click
  TXTnro.Text = a
  CMDinc.Enabled = True
  TXTnro.SetFocus
  Else
  If t!EstatusSec <> "A" Then
  If MsgBox(" La Solicitud esta Inactiva, Desea Reactivarla?", vbQuestion + vbYesNo) = vbYes Then
  Activar
  MsgBox " Solicitud Reactivada Exitosamente ", vbInformation
  mostrar t
  Else
  CMDlim_Click
  TXTnro.SetFocus
  End If
  Else
  mostrar t
  End If
  End If
  End If
End Sub

Private Sub CMDgua_Click()
If TXTnro.Text = "" Then
 MsgBox " Debe escribir el Numero de Solicitud ", vbInformation
 TXTnro.SetFocus
 ElseIf CMBcli.Text = "" Then
  MsgBox " Debe Seleccionar el Codigo del Cliente ", vbInformation
  CMBcli.SetFocus
 ElseIf CMBori.Text = "" Then
  MsgBox " Debe Seleccionar la Ciudad de Origen ", vbInformation
  CMBori.SetFocus
 ElseIf CMBdes.Text = "" Then
  MsgBox " Debe Seleccionar la Ciudad de Destino ", vbInformation
  TXTexi.SetFocus
 ElseIf TXTmin.Text = "" Then
  MsgBox " Debe escribir un Stock minimo para el Repuesto ", vbInformation
  TXTmin.SetFocus
 ElseIf TXTmax.Text = "" Then
  MsgBox " Debe escribir un Stock maximo para el Repuesto ", vbInformation
  TXTmax.SetFocus
 ElseIf TXTuti.Text = "" Then
  MsgBox " Debe escribir el tiempo de vida util ", vbInformation
  TXTuti.SetFocus
Else
 Conexion.Execute "insert into Repuesto (CodRep,NomRep,Costo,Existencia,StMin,StMax,VidaUtil,EstatusRep) values ('" & TXTrep.Text & "', '" & TXTnre.Text & "', '" & TXTcos.Text & "', '" & TXTexi.Text & "', '" & TXTmin.Text & "', '" & TXTmax.Text & "', '" & TXTuti.Text & "', 'A')"
  MsgBox " Repuesto Incluido Satisfactoriamente ", vbExclamation
  CMDlim_Click
  TXTrep.SetFocus
End If
End Sub

Private Sub CMDlim_Click()
TXTnro.Text = ""
TXTnom.Text = ""
TXTori.Text = ""
TXTdes.Text = ""
TXTpro.Text = ""
TXTckg.Text = ""
TXTsol.Text = ""
TXTser.Text = ""
CMBcli.Text = ""
CMBori.Text = ""
CMBdes.Text = ""
CMBpro.Text = ""
CMDinc.Enabled = False
CMDmod.Enabled = False
CMDeli.Enabled = False
End Sub

Private Sub CMDsal_Click()
If MsgBox(" Desea realmente salir de 'Servicios Contratados'?", vbQuestion + vbYesNo) = vbYes Then
Unload Me
End If
End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Sub Activar()

End Sub
