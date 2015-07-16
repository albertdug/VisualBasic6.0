VERSION 5.00
Begin VB.Form Mantreq 
   Caption         =   "Transporte EL SOL // Mantenimientos Requeridos"
   ClientHeight    =   8535
   ClientLeft      =   2130
   ClientTop       =   990
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Mantreq.frx":0000
   ScaleHeight     =   8535
   ScaleWidth      =   10845
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
      Left            =   6480
      Picture         =   "Mantreq.frx":55AE
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton CMDbus 
      Height          =   495
      Left            =   5880
      Picture         =   "Mantreq.frx":5B38
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2640
      Width           =   495
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
      Left            =   6840
      Picture         =   "Mantreq.frx":60C2
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
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
      Left            =   4920
      Picture         =   "Mantreq.frx":664C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
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
      Left            =   3000
      Picture         =   "Mantreq.frx":6BD6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
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
      Left            =   1080
      Picture         =   "Mantreq.frx":7160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox TXTfre 
      Height          =   375
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ComboBox CMBman 
      Height          =   315
      Left            =   3600
      TabIndex        =   5
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox TXTman 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      MaxLength       =   8
      TabIndex        =   3
      Top             =   3360
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
      Left            =   3600
      TabIndex        =   2
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Requeridos"
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
      TabIndex        =   15
      Top             =   840
      Width           =   2340
   End
   Begin VB.Label LBLdia 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Días"
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
      Left            =   4920
      TabIndex        =   8
      Top             =   4080
      Width           =   420
   End
   Begin VB.Label LBLfre 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Frecuencia"
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
      Left            =   1200
      TabIndex        =   6
      Top             =   3960
      Width           =   1170
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
      Left            =   1200
      TabIndex        =   4
      Top             =   3360
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
      Left            =   1200
      TabIndex        =   1
      Top             =   2760
      Width           =   1845
   End
   Begin VB.Label LBLmantenimiento 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mantenimientos"
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
      Left            =   6960
      TabIndex        =   0
      Top             =   360
      Width           =   3360
   End
End
Attribute VB_Name = "Mantreq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New ADODB.Recordset

Private Sub CMDbus_Click()
Buscar
End Sub

Private Sub CMBpla_Click()
Dim t As New ADODB.Recordset
t.Open "select Placa from Camiones where Placa = '" & CMBpla.Text & "'", Conexion
End Sub

Private Sub CMBman_Click()
Dim t As New ADODB.Recordset
t.Open "select CodMant from Mantenimiento where Mantenimiento = '" & CMBman.Text & "'", Conexion
If Not t.EOF Then
    TXTman.Text = t!CodMant
    Buscar
End If
End Sub

Private Sub CMDeli_Click()
If MsgBox(" Desea Realmente Eliminar este Registro?", vbQuestion + vbYesNo, "Eliminacion de Datos") = vbYes Then
Eliminar
MsgBox "Mantenimiento Requerido Eliminado Satisfactoriamente", vbInformation, "Eliminacion de Datos"
Limpiar
End If
End Sub

Private Sub CMDinc_Click()
If CMBpla.Text = "" Then
  MsgBox " Debe Elegir la Placa del Camión", vbInformation, "Omision de Datos"
  CMBpla.SetFocus
  ElseIf CMBman.Text = "" Then
    MsgBox " Debe Elegir El Mantenimiento que Corresponde", vbInformation, "Omision de Datos"
    CMBman.SetFocus
  ElseIf TXTman.Text = "" Then
    MsgBox " Debe Elegir El Mantenimiento que Corresponde", vbInformation, "Omision de Datos"
    CMBman.SetFocus
  ElseIf TXTfre.Text = "" Then
    MsgBox " Debe Incluir la Frecuencia de Realización de Mantenimiento", vbInformation, "Omision de Datos"
    TXTfre.SetFocus
Else
    Conexion.Execute "insert into MantReqPorCam (Placa, CodMant, FrecuanciaDias, EstatusMrc) values ('" & CMBpla.Text & "', '" & TXTman.Text & "', '" & TXTfre.Text & "', 'A')"
    MsgBox " Mantenimiento Requerido Incluido Satisfactoriamente ", vbInformation, "Inclusion de Datos"
    Limpiar
    CMBpla.SetFocus
End If
End Sub

Private Sub CMDlim_Click()
Limpiar
End Sub

Private Sub CMDmod_Click()
If CMBpla.Text = "" Then
  MsgBox " Debe Elegir la Placa del Camión", vbInformation, "Omision de Datos"
  CMBpla.SetFocus
  ElseIf CMBman.Text = "" Then
    MsgBox " Debe Elegir El Mantenimiento que Corresponde", vbInformation, "Omision de Datos"
    CMBman.SetFocus
  ElseIf TXTman.Text = "" Then
    MsgBox " Debe Elegir El Mantenimiento que Corresponde", vbInformation, "Omision de Datos"
    CMBman.SetFocus
  ElseIf TXTfre.Text = "" Then
    MsgBox " Debe Incluir la Frecuencia de Realización de Mantenimiento", vbInformation, "Omision de Datos"
    TXTfre.SetFocus
Else
    Conexion.Execute " update MantReqPorCam set FrecuanciaDias = '" & TXTfre.Text & "'"
    MsgBox " Mantenimiento Requerido Modificado Satisfactoriamente ", vbInformation, "Modificacion de Datos"
    Limpiar
    CMBpla.SetFocus
End If
End Sub

Private Sub CMDsal_Click()
If MsgBox(" Desea realmente salir de 'Mantenimiento por Camión'?", vbQuestion + vbYesNo, "Salida de Pantalla") = vbYes Then
Unload Me
End If
End Sub

Private Sub Form_Load()
LlenarCMB CMBpla, "select * from Camiones", "Placa"
LlenarCMB CMBman, "select * from Mantenimiento", "Mantenimiento"
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
Limpiar
End Sub

Sub Limpiar()
CMBman.Text = ""
CMBpla.Text = ""
TXTman.Text = ""
TXTfre.Text = ""
CMDinc.Enabled = False
CMDmod.Enabled = False
CMDeli.Enabled = False
End Sub

Sub Eliminar()
Conexion.Execute "update MantReqPorCam set EstatusMrc = 'I' where Placa = '" & CMBpla.Text & "'"
End Sub

Sub Buscar()
Dim t As New ADODB.Recordset
If CMBpla.Text = "" Then
  MsgBox " Debe Seleccionar la Placa del Camión", vbInformation, "Omision de Datos"
  CMBpla.SetFocus
  ElseIf TXTman.Text = "" Then
    MsgBox " Debe Seleccionar el Mantenimiento", vbInformation, "Omision de Datos"
    CMBman.SetFocus
Else
t.Open "select * from MantReqPorCam where Placa  = '" & CMBpla.Text & "' and CodMant = '" & TXTman.Text & "'", Conexion
If t.EOF Then
 If MsgBox(" Este Mantenimiento Requerido No Existe, Desea Crearlo Ahora?", vbQuestion + vbYesNo, "Inclusion de Datos") = vbYes Then
    CMDinc.Enabled = True
    TXTfre.SetFocus
    Else
        Limpiar
        CMBpla.SetFocus
    End If
 Else
    If t!EstatusMrc <> "A" Then
        If MsgBox(" Este Mantenimiento Requerido esta actualmente Inactivo, Desea Reactivarlo?", vbQuestion + vbYesNo, "Reactivacion de Datos") = vbYes Then
            activar
            MsgBox " Mantenimiento Requerido Reactivado Satisfactoriamente ", vbInformation, "Reactivacion de Datos"
            mostrar t
        Else
            Limpiar
            CMBpla.SetFocus
        End If
    Else
        mostrar t
    End If
 End If
End If

End Sub

Sub mostrar(t As ADODB.Recordset)
TXTfre.Text = t!FrecuanciaDias
CMDinc.Enabled = False
CMDmod.Enabled = True
CMDeli.Enabled = True
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
If IsNumeric(TXTfre.Text) = False Then
    MsgBox "Este Campo es de Tipo Numerico, Cambie la Condicion", vbInformation, "Datos Incorrectos"
    TXTfre.Text = ""
    TXTfre.SetFocus
End If
If Val(TXTfre.Text) < 0 Then
 MsgBox " La Frecuencia de Mantenimiento no puede ser Negativo, Cambie la Condición", vbInformation, "Datos Incorrectos"
 TXTfre.Text = ""
End If
End If
End Sub

Sub activar()
Conexion.Execute "update MantReqPorCam set EstatusMrc = 'A' where Placa = '" & CMBpla.Text & "' and CodMant = '" & TXTman.Text & "'"
End Sub
