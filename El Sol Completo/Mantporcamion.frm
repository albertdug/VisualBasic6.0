VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Mantporcamion 
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   1860
   ClientTop       =   990
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Mantporcamion.frx":0000
   ScaleHeight     =   8730
   ScaleWidth      =   10800
   Begin MSComCtl2.DTPicker DTfecmant 
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Top             =   4200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   23330817
      CurrentDate     =   39958
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
      Left            =   3240
      Picture         =   "Mantporcamion.frx":55AE
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5520
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
      Left            =   5880
      Picture         =   "Mantporcamion.frx":5B38
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2880
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
      Left            =   5160
      Picture         =   "Mantporcamion.frx":60C2
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton CMDinc 
      Caption         =   "&Incluir"
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
      Picture         =   "Mantporcamion.frx":664C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox TXThor 
      Height          =   375
      Left            =   3720
      MaxLength       =   8
      TabIndex        =   8
      Top             =   4800
      Width           =   2055
   End
   Begin VB.ComboBox CMBman 
      Height          =   315
      Left            =   3720
      TabIndex        =   5
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox TXTman 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      MaxLength       =   8
      TabIndex        =   3
      Top             =   3480
      Width           =   855
   End
   Begin VB.ComboBox CMBpla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Por Camión "
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
      Left            =   7560
      TabIndex        =   13
      Top             =   1080
      Width           =   2625
   End
   Begin VB.Label LBLmil 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "H/mil"
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
      Left            =   5880
      TabIndex        =   9
      Top             =   4800
      Width           =   465
   End
   Begin VB.Label LBLhor 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Hora"
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
      TabIndex        =   7
      Top             =   4800
      Width           =   525
   End
   Begin VB.Label LBLfec 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fecha"
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
      TabIndex        =   6
      Top             =   4200
      Width           =   660
   End
   Begin VB.Label LBLcol 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Tipo Mantenimiento"
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
      TabIndex        =   4
      Top             =   3600
      Width           =   2055
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
      TabIndex        =   1
      Top             =   3000
      Width           =   1845
   End
   Begin VB.Label LBLmantporcamion 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mantenimientos "
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
      Left            =   7080
      TabIndex        =   0
      Top             =   480
      Width           =   3480
   End
End
Attribute VB_Name = "Mantporcamion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New ADODB.Recordset

Private Sub CMBpla_Click()
Dim t As New ADODB.Recordset
t.Open "select Placa from Camiones where Placa = '" & CMBpla.Text & "'", Conexion
CMDinc.Enabled = True
End Sub

Private Sub CMBman_Click()
Dim t As New ADODB.Recordset
t.Open "select CodMant from Mantenimiento where Mantenimiento = '" & CMBman.Text & "'", Conexion
If Not t.EOF Then
    TXTman.Text = t!CodMant
    CMDinc.Enabled = True
End If
End Sub

Private Sub CMDinc_Click()
Dim m As New ADODB.Recordset
If CMBpla.Text = "" Then
  MsgBox " Debe Elegir la Placa del Camión", vbInformation
  CMBpla.SetFocus
  ElseIf CMBman.Text = "" Then
    MsgBox " Debe Elegir El Mantenimiento que Corresponde", vbInformation
    CMBman.SetFocus
  ElseIf TXTman.Text = "" Then
    MsgBox " Debe Elegir El Mantenimiento que Corresponde", vbInformation
    CMBman.SetFocus
  ElseIf DTfecmant.Value = "" Then
    MsgBox " Debe Incluir la Fecha de Realización de Mantenimiento", vbInformation
    DTfecmant.SetFocus
  ElseIf TXThor.Text = "" Then
    MsgBox " Debe Incluir la Hora de Realización de Mantenimiento", vbInformation
    TXThor.SetFocus
Else
    m.Open " select placa, codmant, frecuanciadias, fecha from mantreqporcam, mantenporcam where placa = placam and codmant = codmantm", Conexion
     If DTfecmant.Value < (m!Fecha + m!frecuanciadias) Then
    MsgBox " No Puede Efectuar Este Mantenimiento Al Camion Indicado. La Frecuencia de realizacion del mismo, aun no ha expirado", vbExclamation, "No puede proceder"
Else
Conexion.Execute "insert into MantenPorCam (Placam, CodMantm, Fecha, Hora, EstatusMxc) values ('" & Trim(UCase(CMBpla.Text)) & "', '" & Trim(UCase(TXTman.Text)) & "', '" & DTfecmant.Value & "', '" & Trim(TXThor.Text) & "', 'A')"
    MsgBox " Mantenimiento Por Camión Incluido Satisfactoriamente ", vbInformation
    Limpiar
    CMBpla.SetFocus
End If
End If
End Sub

Private Sub CMDlim_Click()
Limpiar
End Sub

Private Sub CMDRepor_Click()
ReporteMantReq.Show
End Sub

Private Sub CMDsal_Click()
If MsgBox(" Desea realmente salir de 'Mantenimiento por Camión'?", vbQuestion + vbYesNo) = vbYes Then
Unload Me
End If
End Sub

Private Sub Form_Load()
LlenarCMB CMBpla, "select * from Camiones", "Placa"
LlenarCMB CMBman, "select * from Mantenimiento", "Mantenimiento"
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
DTfecmant.Value = GLOBALDATE
Limpiar
End Sub

Sub Limpiar()
CMBman.Text = ""
CMBpla.Text = ""
TXTman.Text = ""
TXThor.Text = ""
CMDinc.Enabled = False
End Sub

Sub LlenarCMB(c As ComboBox, tabla, campo)
Dim t As New ADODB.Recordset
c.Clear
t.Open tabla, Conexion
While Not t.EOF
c.AddItem t(campo)
t.MoveNext
Wend
End Sub

Private Sub TXTfre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Val(TXTfre.Text) < 0 Then
 MsgBox " La Frecuencia de Mantenimiento no puede ser Negativo, Cambie la Condición", vbInformation
 TXTfre.Text = ""
Else
CMDinc_Click
End If
End If
End Sub

Private Sub TXThor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsDate(TXThor.Text) = False Then
MsgBox " Este Campo es de Tipo Hora, Cambie la Condición", vbExclamation, "Datos Incorrectos"
TXThor.Text = ""
TXThor.SetFocus
End If
End If
End Sub
