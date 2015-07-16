VERSION 5.00
Begin VB.Form Tabulador 
   Caption         =   "Transporte EL SOL // Tabulador"
   ClientHeight    =   8115
   ClientLeft      =   2505
   ClientTop       =   735
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   Picture         =   "Tabulador.frx":0000
   ScaleHeight     =   8115
   ScaleWidth      =   10725
   Begin VB.CommandButton CMDlim 
      Caption         =   "&Limpiar"
      Height          =   495
      Left            =   7440
      Picture         =   "Tabulador.frx":55AE
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton CMDbus 
      Height          =   495
      Left            =   6840
      Picture         =   "Tabulador.frx":5B38
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton CMDsal 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   6960
      Picture         =   "Tabulador.frx":60C2
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton CMDeli 
      Caption         =   "&Eliminar"
      Height          =   495
      Left            =   5040
      Picture         =   "Tabulador.frx":664C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton CMDmod 
      Caption         =   "&Modificar"
      Height          =   495
      Left            =   3120
      Picture         =   "Tabulador.frx":6BD6
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton CMDinc 
      Caption         =   "&Incluir"
      Height          =   495
      Left            =   1200
      Picture         =   "Tabulador.frx":7160
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5400
      Width           =   1815
   End
   Begin VB.TextBox TXTbsk 
      Height          =   375
      Left            =   3600
      MaxLength       =   6
      TabIndex        =   11
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox TXTdia 
      Height          =   375
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ComboBox CMBdes 
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
      TabIndex        =   6
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox TXTdes 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      MaxLength       =   8
      TabIndex        =   5
      Top             =   3360
      Width           =   855
   End
   Begin VB.ComboBox CMBori 
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
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox TXTori 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      MaxLength       =   8
      TabIndex        =   2
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label LBLbsk 
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
      Index           =   1
      Left            =   4920
      TabIndex        =   12
      Top             =   4800
      Width           =   390
   End
   Begin VB.Label LBLcap 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Costo por Kilogramo"
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
      Left            =   1200
      TabIndex        =   10
      Top             =   4680
      Width           =   2130
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
      TabIndex        =   9
      Top             =   4080
      Width           =   420
   End
   Begin VB.Label LBLcap 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Tiempo de Duracion"
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
      TabIndex        =   7
      Top             =   3960
      Width           =   2130
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
      Left            =   1200
      TabIndex        =   4
      Top             =   3360
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
      Left            =   1200
      TabIndex        =   1
      Top             =   2760
      Width           =   1830
   End
   Begin VB.Label LBLtabulador 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Tabulador"
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
      Top             =   720
      Width           =   2190
   End
End
Attribute VB_Name = "Tabulador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New ADODB.Recordset

Private Sub Limpiar()
TXTori.Text = ""
TXTdes.Text = ""
CMBori.Text = ""
CMBdes.Text = ""
TXTbsk.Text = ""
TXTdia.Text = ""
CMDinc.Enabled = False
CMDmod.Enabled = False
CMDeli.Enabled = False
TXTdia.Enabled = True
End Sub

Private Sub CMBdes_Click()
Dim t As New ADODB.Recordset
t.Open "select CodCiudad from Ciudades where Ciudades = '" & CMBdes.Text & "'", Conexion
If Not t.EOF And CMBdes.Text <> CMBori.Text Then
    TXTdes.Text = t!CodCiudad
    buscar
Else
    MsgBox (" No se permiten Viajes hacia la Misma Ciudad de Origen, Cambie la Condición"), vbInformation, "Datos Incorrectos"
    CMBdes.Text = ""
    CMBdes.SetFocus
End If
If CMBori.Text = "" Then
MsgBox "Debe Elegir Ciudad de Origen", vbInformation, "Omision de Datos"
CMBori.SetFocus
End If
End Sub

Private Sub CMBori_Click()
Dim t As New ADODB.Recordset
t.Open "select CodCiudad from Ciudades where Ciudades = '" & CMBori.Text & "'", Conexion
If Not t.EOF And CMBori.Text <> CMBdes.Text Then
    TXTori.Text = t!CodCiudad
Else
    MsgBox (" No se permiten Viajes hacia la Misma Ciudad de Destino, Cambie la Condición"), vbInformation, "Datos Incorrectos"
    CMBori.Text = ""
    CMBori.SetFocus
End If
End Sub

Private Sub CMDbus_Click()
buscar
End Sub

Private Sub CMDeli_Click()
If MsgBox(" Desea Realmente Eliminar este Tabulador?", vbQuestion + vbYesNo, "Eliminacion de Datos") = vbYes Then
Eliminar
MsgBox "Tabulador Eliminado Satisfactoriamente", vbInformation, "Eliminacion de Datos"
Limpiar
End If
End Sub

Private Sub CMDinc_Click()
If CMBori.Text = "" Then
  MsgBox " Debe Elegir la Ciudad de Origen", vbInformation, "Omision de Datos"
  CMBori.SetFocus
  ElseIf TXTori.Text = "" Then
    MsgBox " Debe Elegir la Ciudad de Origen", vbInformation, "Omision de Datos"
    CMBori.SetFocus
  ElseIf CMBdes.Text = "" Then
    MsgBox " Debe Elegir la Ciudad de Destino", vbInformation, "Omision de Datos"
    CMBdes.SetFocus
  ElseIf TXTdes.Text = "" Then
    MsgBox " Debe Elegir la Ciudad de Origen", vbInformation, "Omision de Datos"
    CMBdes.SetFocus
  ElseIf TXTdia.Text = "" Then
    MsgBox " Debe Incluir la Cantidad de Dias ", vbInformation, "Omision de Datos"
    Textcodco.SetFocus
  ElseIf TXTbsk.Text = "" Then
    MsgBox " Debe Incluir el Precio del Kg. En este Tabulador ", vbInformation, "Omision de Datos"
    TXTbsk.SetFocus
Else
    Conexion.Execute "insert into Tabulador (CodCiudadInicio, CodCiudadDestino, TiempoEstimDias, BsPorKgTransp, EstatusTab) values ('" & TXTori.Text & "', '" & TXTdes.Text & "', '" & TXTdia.Text & "', '" & TXTbsk.Text & "', 'A')"
    MsgBox " Tabulador Incluido Satisfactoriamente ", vbInformation, "Inclusion de Datos"
    Limpiar
    CMBori.SetFocus
End If
End Sub

Private Sub CMDlim_Click()
Limpiar
End Sub

Private Sub CMDmod_Click()
If CMBori.Text = "" Then
  MsgBox " Debe Elegir la Ciudad de Origen", vbInformation, "Omision de Datos"
  CMBori.SetFocus
  ElseIf TXTori.Text = "" Then
    MsgBox " Debe Elegir la Ciudad de Origen", vbInformation, "Omision de Datos"
    CMBori.SetFocus
  ElseIf CMBdes.Text = "" Then
    MsgBox " Debe Elegir la Ciudad de Destino", vbInformation, "Omision de Datos"
    CMBdes.SetFocus
  ElseIf TXTdes.Text = "" Then
    MsgBox " Debe Elegir la Ciudad de Origen", vbInformation, "Omision de Datos"
    CMBdes.SetFocus
  ElseIf TXTdia.Text = "" Then
    MsgBox " Debe Incluir la Cantidad de Dias ", vbInformation, "Omision de Datos"
    Textcodco.SetFocus
  ElseIf TXTbsk.Text = "" Then
    MsgBox " Debe Incluir el Precio del Kg. En este Tabulador ", vbInformation, "Omision de Datos"
    TXTbsk.SetFocus
Else
    Conexion.Execute " update Tabulador set TiempoEstimDias = '" & TXTdia.Text & "', BsPorKgTransp = '" & TXTbsk.Text & "' where CodCiudadInicio = '" & TXTori.Text & "' and CodCiudadDestino = '" & TXTdes.Text & "'"
    Conexion.Execute " update Tabulador set TiempoEstimDias = '" & TXTdia.Text & "', BsPorKgTransp = '" & TXTbsk.Text & "' where CodCiudadInicio = '" & TXTdes.Text & "' and CodCiudadDestino = '" & TXTori.Text & "'"
    MsgBox " Tabulador Modificado Satisfactoriamente ", vbInformation, "Modificacion de Datos"
    Limpiar
    CMBori.SetFocus
End If
End Sub

Private Sub CMDsal_Click()
If MsgBox(" Desea realmente salir de 'Tabulador'?", vbQuestion + vbYesNo, "Salida de Pantalla") = vbYes Then
Unload Me
End If
End Sub

Private Sub Form_Load()
LlenarCMB CMBori, "select * from Ciudades", "Ciudades"
LlenarCMB CMBdes, "select * from Ciudades", "Ciudades"
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
Limpiar
End Sub
Sub buscar()
Dim t As New ADODB.Recordset
If CMBori.Text = "" Then
  MsgBox " Debe Elegir la Ciudad de Origen", vbInformation, "Omision de Datos"
  CMBori.SetFocus
  ElseIf CMBdes.Text = "" Then
    MsgBox " Debe Elegir la Ciudad de Destino", vbInformation, "Omision de Datos"
    CMBdes.SetFocus
Else
t.Open "select CodCiudadInicio,CodCiudadDestino,TiempoEstimDias,BsPorKgTransp,EstatusTab  from Tabulador where CodCiudadInicio  = '" & TXTori.Text & "' and CodCiudadDestino = '" & TXTdes.Text & "'", Conexion
If t.EOF Then
t.Close
t.Open "Select CodCiudadInicio,CodCiudadDestino,TiempoEstimDias,BsPorKgTransp,EstatusTab from Tabulador where CodCiudadInicio = '" & TXTdes.Text & "' and CodCiudadDestino = '" & TXTori.Text & "'", Conexion
 If t.EOF Then
  If MsgBox(" Este Tabulador No Existe, Desea Crearlo Ahora?", vbQuestion + vbYesNo, "Inclusion de Datos") = vbYes Then
    CMDinc.Enabled = True
    TXTdia.SetFocus
    Else
    Limpiar
    CMBori.SetFocus
    End If
  Else
   If t!EstatusTab = "A" Then
   mostrar t
   TXTdia.Enabled = False
   Else
   If MsgBox(" Este Tabulador esta actualmente Inactivo, Desea Reactivarlo?", vbQuestion + vbYesNo, "Reactivacion de Datos") = vbYes Then
            activar
            MsgBox " Registro Reactivado Satisfactoriamente ", vbInformation, "Reactivación de Datos"
            mostrar t
            TXTdia.Enabled = False
    End If
    End If
End If
 Else
    If t!EstatusTab <> "A" Then
        If MsgBox(" Este Tabulador esta actualmente Inactivo, Desea Reactivarlo?", vbQuestion + vbYesNo, "Reactivacion de Datos") = vbYes Then
            activar
            MsgBox " Registro Reactivado Satisfactoriamente ", vbInformation, "Reactivación de Datos"
            mostrar t
            TXTdia.Enabled = False
        Else
            Limpiar
            CMBori.SetFocus
        End If
    Else
        mostrar t
        TXTdia.Enabled = False
    End If
 End If
End If
End Sub

Sub activar()
Conexion.Execute "update Tabulador set EstatusTab = 'A' where CodCiudadInicio = '" & TXTori.Text & "' and CodCiudadDestino = '" & TXTdes.Text & "'"
End Sub

Sub Eliminar()
Conexion.Execute "update Tabulador set EstatusTab = 'I' where CodCiudadInicio = '" & TXTori.Text & "' and CodCiudadDestino = '" & TXTdes.Text & "'"
Conexion.Execute "update Tabulador set EstatusTab = 'I' where CodCiudadInicio = '" & TXTdes.Text & "' and CodCiudadDestino = '" & TXTori.Text & "'"
End Sub


Sub mostrar(t As ADODB.Recordset)

TXTdia.Text = t!TiempoEstimDias
TXTbsk.Text = t!BsPorKgTransp
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

Private Sub TXTbsk_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(TXTbsk.Text) = False Then
    MsgBox "Este Campo es de Tipo Numerico, Cambie la Condición", vbInformation, "Datos Incorrectos"
    TXTbsk.Text = ""
    TXTbsk.SetFocus
End If
If Val(TXTbsk.Text) < 0 Then
 MsgBox " La Cantidad en Kilogramos no puede ser Negativo, Cambie la Condición", vbInformation, "Datos Incorrectos"
 TXTbsk.Text = ""
End If
End If
End Sub

Private Sub TXTdia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(TXTdia.Text) = False Then
    MsgBox "Este Campo es de Tipo Numerico, Cambie la Condición", vbInformation, "Datos Incorrectos"
    TXTdia.Text = ""
    TXTdia.SetFocus
End If
If Val(TXTdia.Text) < 0 Then
 MsgBox " La Duración en Días no puede ser Negativo, Cambie la Condición", vbInformation, "Datos Incorrectos"
 TXTdia.Text = ""
Else
TXTbsk.SetFocus
End If
End If
End Sub
