VERSION 5.00
Begin VB.Form Repuestos 
   BackColor       =   &H80000005&
   Caption         =   "Trasporte EL SOL // Repuestos"
   ClientHeight    =   8670
   ClientLeft      =   2130
   ClientTop       =   990
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   Picture         =   "Repuestos.frx":0000
   ScaleHeight     =   8670
   ScaleWidth      =   10920
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
      Left            =   5880
      Picture         =   "Repuestos.frx":55AE
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2640
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
      Left            =   8640
      Picture         =   "Repuestos.frx":5B38
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton CMDbus 
      Height          =   495
      Left            =   7920
      Picture         =   "Repuestos.frx":60C2
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2640
      Width           =   495
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
      Picture         =   "Repuestos.frx":664C
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7680
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
      Picture         =   "Repuestos.frx":6BD6
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7680
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
      Picture         =   "Repuestos.frx":7160
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7680
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
      Left            =   7080
      Picture         =   "Repuestos.frx":76EA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7680
      Width           =   1815
   End
   Begin VB.TextBox TXTuti 
      Height          =   375
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   14
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox TXTmax 
      Height          =   375
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   12
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox TXTmin 
      Height          =   375
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   10
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox TXTexi 
      Height          =   375
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   8
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox TXTcos 
      Height          =   375
      Left            =   3720
      MaxLength       =   30
      TabIndex        =   5
      Top             =   4080
      Width           =   2415
   End
   Begin VB.TextBox TXTnre 
      Height          =   375
      Left            =   3720
      MaxLength       =   30
      TabIndex        =   2
      Top             =   3360
      Width           =   2895
   End
   Begin VB.TextBox TXTrep 
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
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label LBLmes 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mes (es)"
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
      Left            =   5160
      TabIndex        =   16
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Vida Útil"
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
      TabIndex        =   15
      Top             =   6960
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Stock Máximo"
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
      TabIndex        =   13
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Stock Mínimo"
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
      TabIndex        =   11
      Top             =   5520
      Width           =   1395
   End
   Begin VB.Label LBLexistencia 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Existencia"
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
      TabIndex        =   9
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label LBLbs 
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
      Index           =   0
      Left            =   6240
      TabIndex        =   7
      Top             =   4200
      Width           =   390
   End
   Begin VB.Label LBLcosto 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Costo"
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
      TabIndex        =   6
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label LBLnom 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Nombre del Repuesto"
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
      TabIndex        =   4
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label LBLrep 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Codigo del Repuesto"
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
      TabIndex        =   3
      Top             =   2640
      Width           =   2220
   End
   Begin VB.Label LBLrepuestos 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Repuestos"
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
      Left            =   8640
      TabIndex        =   0
      Top             =   720
      Width           =   2070
   End
End
Attribute VB_Name = "Repuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDbus_Click()
Buscar
End Sub
Sub activar()
Conexion.Execute " update Repuesto  set EstatusRep = 'A' where CodRep = '" & TXTrep.Text & "'"
End Sub

Sub Buscar()
Dim t As New ADODB.Recordset
If TXTrep.Text = "" Then
  MsgBox " Debe Escribir el Codigo del Repuesto", vbInformation, "Omision de Datos"
  TXTrep.SetFocus
Else
  t.Open "Select * from Repuesto where CodRep = '" & TXTrep.Text & "'", Conexion
 If t.EOF Then
  If MsgBox(" El Repuesto No esta Registrado, Desea Incluirlo ahora?", vbQuestion + vbYesNo, "Inclusion de Datos") = vbYes Then
  a = TXTrep.Text
  CMDlim_Click
  TXTrep.Text = a
  CMDinc.Enabled = True
  TXTnre.SetFocus
  Else
  CMDlim_Click
  TXTrep.SetFocus
  End If
  Else
  If t!EstatusRep <> "A" Then
  If MsgBox(" El Repuesto esta Inactivo, Desea Reactivarlo?", vbQuestion + vbYesNo, "Reactivacion de Datos") = vbYes Then
  activar
  MsgBox " El Repuesto se ha Reactivado Exitosamente ", vbInformation, "Reactivacion de Datos"
  mostrar t
  TXTnre.Enabled = False
  Else
  CMDlim_Click
  TXTrep.SetFocus
  End If
  Else
  mostrar t
  TXTnre.Enabled = False
  End If
  End If
  End If
End Sub
Private Sub CMDeli_Click()
Dim t As New ADODB.Recordset
t.Open "Select * from RepporCamion where CodRep = '" & TXTrep.Text & "' and EstatusRxc ='A'", Conexion
If Not t.EOF Then
    MsgBox "Usted No puede Eliminar este Repuesto, ya que Un(os) Camion(es) lo necesitan", vbExclamation, "Eliminacion de Datos"
    Else
    t.Close
    t.Open "Select * from Repuesto where CodRep = '" & TXTrep.Text & "'", Conexion
    If Not t.EOF Then
    Eliminar
    End If
End If
End Sub

Sub Eliminar()
If TXTrep.Text = "" Then
 MsgBox " Debe escribir un Codigo de Repuesto ", vbInformation, "Omision de Datos"
 TXTrep.SetFocus
 ElseIf TXTnre.Text = "" Then
  MsgBox " Debe Escribir el Nombre del Repuesto ", vbInformation, "Omision de Datos"
  TXTnre.SetFocus
 ElseIf TXTcos.Text = "" Then
  MsgBox " Debe escribir el costo del Repuesto ", vbInformation, "Omision de Datos"
  TXTcos.SetFocus
 ElseIf TXTexi.Text = "" Then
  MsgBox " Debe escribir la existencia del Repuesto ", vbInformation, "Omision de Datos"
  TXTexi.SetFocus
 ElseIf TXTmin.Text = "" Then
  MsgBox " Debe escribir un Stock minimo para el Repuesto ", vbInformation, "Omision de Datos"
  TXTmin.SetFocus
 ElseIf TXTmax.Text = "" Then
  MsgBox " Debe escribir un Stock maximo para el Repuesto ", vbInformation, "Omision de Datos"
  TXTmax.SetFocus
 ElseIf TXTuti.Text = "" Then
  MsgBox " Debe escribir el tiempo de vida util ", vbInformation, "Omision de Datos"
  TXTuti.SetFocus
Else
 If MsgBox(" Desea Eliminar el Repuesto?", vbQuestion + vbYesNo, "Eliminacion de Datos") = vbYes Then
 Conexion.Execute " update Repuesto set EstatusRep = 'E' where CodRep = '" & TXTrep.Text & "'"
   MsgBox " Repuesto Eliminado Satisfactoriamente", vbInformation, "Eliminacion de Datos"
   CMDlim_Click
   TXTrep.SetFocus
 End If
 End If
End Sub
Private Sub CMDinc_Click()
Dim t As New ADODB.Recordset
If TXTrep.Text = "" Then
 MsgBox " Debe escribir un Codigo de Repuesto ", vbInformation, "Omision de Datos"
 TXTrep.SetFocus
 ElseIf TXTnre.Text = "" Then
  MsgBox " Debe Escribir el Nombre del Repuesto ", vbInformation, "Omision de Datos"
  TXTnre.SetFocus
 ElseIf TXTcos.Text = "" Then
  MsgBox " Debe escribir el costo del Repuesto ", vbInformation, "Omision de Datos"
  TXTcos.SetFocus
 ElseIf TXTexi.Text = "" Then
  MsgBox " Debe escribir la existencia del Repuesto ", vbInformation, "Omision de Datos"
  TXTexi.SetFocus
 ElseIf TXTmin.Text = "" Then
  MsgBox " Debe escribir un Stock minimo para el Repuesto ", vbInformation, "Omision de Datos"
  TXTmin.SetFocus
 ElseIf TXTmax.Text = "" Then
  MsgBox " Debe escribir un Stock maximo para el Repuesto ", vbInformation, "Omision de Datos"
  TXTmax.SetFocus
 ElseIf TXTuti.Text = "" Then
  MsgBox " Debe escribir el tiempo de vida util ", vbInformation, "Omision de Datos"
  TXTuti.SetFocus
Else
    t.Open "select NomRep from Repuesto where NomRep = '" & Trim(UCase(TXTnre.Text)) & "' and EstatusRep = 'A'", Conexion
        If t.EOF Then
            Conexion.Execute "insert into Repuesto (CodRep,NomRep,Costo,Existencia,StMin,StMax,VidaUtil,EstatusRep) values ('" & TXTrep.Text & "', '" & Trim(UCase(TXTnre.Text)) & "', '" & TXTcos.Text & "', '" & TXTexi.Text & "', '" & TXTmin.Text & "', '" & TXTmax.Text & "', '" & TXTuti.Text & "', 'A')"
            MsgBox " Repuesto Incluido Exitosamente ", vbInformation, "Inclusion de Datos"
            CMDlim_Click
            TXTrep.SetFocus
        Else
           MsgBox " No Puede Incluir este Repuesto PorQue Ya Existe Un Repuesto Con Ese Nombre", vbInformation, "Duplicacion De Datos"
            CMDlim_Click
            TXTrep.SetFocus
    End If
End If
End Sub

Private Sub CMDlim_Click()
TXTrep.Text = ""
TXTnre.Text = ""
TXTcos.Text = ""
TXTmin.Text = ""
TXTmax.Text = ""
TXTexi.Text = ""
TXTuti.Text = ""
CMDinc.Enabled = False
CMDmod.Enabled = False
CMDeli.Enabled = False
TXTnre.Enabled = True
TXTrep.Enabled = True

End Sub

Private Sub mostrar(t As ADODB.Recordset)
TXTnre.Text = t!NomRep
TXTcos.Text = t!Costo
TXTmin.Text = t!StMin
TXTmax.Text = t!StMax
TXTexi.Text = t!Existencia
TXTuti.Text = t!VidaUtil
CMDinc.Enabled = False
CMDmod.Enabled = True
CMDeli.Enabled = True
TXTrep.Enabled = False
End Sub

Private Sub CMDmod_Click()
If TXTrep.Text = "" Then
 MsgBox " Debe escribir un Codigo de Repuesto ", vbInformation, "Omision de Datos"
 TXTrep.SetFocus
 ElseIf TXTnre.Text = "" Then
  MsgBox " Debe Escribir el Nombre del Repuesto ", vbInformation, "Omision de Datos"
  TXTnre.SetFocus
 ElseIf TXTcos.Text = "" Then
  MsgBox " Debe escribir el costo del Repuesto ", vbInformation, "Omision de Datos"
  TXTcos.SetFocus
 ElseIf TXTexi.Text = "" Then
  MsgBox " Debe escribir la existencia del Repuesto ", vbInformation, "Omision de Datos"
  TXTexi.SetFocus
 ElseIf TXTmin.Text = "" Then
  MsgBox " Debe escribir un Stock minimo para el Repuesto ", vbInformation, "Omision de Datos"
  TXTmin.SetFocus
 ElseIf TXTmax.Text = "" Then
  MsgBox " Debe escribir un Stock maximo para el Repuesto ", vbInformation, "Omision de Datos"
  TXTmax.SetFocus
 ElseIf TXTuti.Text = "" Then
  MsgBox " Debe escribir el tiempo de vida util ", vbInformation, "Omision de Datos"
  TXTuti.SetFocus
Else
 Conexion.Execute " update Repuesto set NomRep = '" & TXTnre.Text & "', Costo = '" & TXTcos.Text & "', Existencia = '" & TXTexi.Text & "', StMin = '" & TXTmin.Text & "', VidaUtil = '" & TXTuti.Text & "' where CodRep = '" & TXTrep.Text & "'"
  MsgBox " Repuesto Modificado Satisfactoriamente ", vbInformation, "Modificacion de Datos"
  CMDlim_Click
  TXTrep.SetFocus
End If
End Sub

Private Sub CMDnue_Click()
Dim t As New ADODB.Recordset
    t.Open " select max(codrep) from repuesto", Conexion
If Not IsNumeric(t(0)) Then
    TXTrep.Text = "001"
Else
    TXTrep.Text = Format(Str(t(0)) + 1, "000")
End If
    TXTnre.SetFocus
    CMDinc.Enabled = True
End Sub

Private Sub CMDsal_Click()
If MsgBox(" Desea realmente salir de 'Repuestos'?", vbQuestion + vbYesNo, "Salida de Pantalla") = vbYes Then
Unload Me
End If
End Sub

Private Sub Form_Load()
CMDlim_Click
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Private Sub TXTcos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(TXTcos.Text) = False Then
    MsgBox "Este Campo es Numerico, Cambie la Condición", vbExclamation, "Dato Incorrecto"
    TXTcos.Text = ""
    TXTcos.SetFocus
    ElseIf Val(TXTcos.Text) < 0 Then
MsgBox " El Costo del Repuesto no puede ser Negativo, Cambie la Condición", vbExclamation, "Dato Incorrecto"
TXTcos.Text = ""
TXTcos.SetFocus
Else
TXTexi.SetFocus
End If
End If
End Sub

Private Sub TXTexi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(TXTexi.Text) = False Then
    MsgBox "Este Campo es Numerico, Cambie la Condición", vbExclamation, "Dato Incorrecto"
    TXTexi.Text = ""
    TXTexi.SetFocus
ElseIf Val(TXTexi.Text) < 0 Then
    MsgBox " La Existencia de Un Repuesto, no puede ser Negativo, Cambie la Condición", vbExclamation, "Dato Incorrecto"
    TXTexi.Text = ""
    TXTexi.SetFocus
Else
    TXTmin.SetFocus
End If
End If
End Sub

Private Sub TXTmax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(TXTmax.Text) = False Then
    MsgBox "Este Campo es Numerico, Cambie la Condición", vbExclamation, "Dato Incorrecto"
    TXTmax.Text = ""
    TXTmax.SetFocus
ElseIf Val(TXTmax.Text) < 0 Then
    MsgBox " El Stock Minimo de un Repuesto, no puede ser Negativo, Cambie la Condición", vbExclamation, "Dato Incorrecto"
    TXTmax.Text = ""
    TXTmax.SetFocus
ElseIf Val(TXTmax.Text) < Val(TXTmin.Text) Then
    MsgBox " El Stock Maximo, no puede ser menor al Stock Minimo. Verifique Los Datos", vbExclamation, "Dato Incorrecto"
    TXTmax.Text = ""
    TXTmax.SetFocus
ElseIf Val(TXTmax.Text) < Val(TXTexi.Text) Then
    MsgBox " El Stock Maximo, no puede ser menor a la Existencia. Verifique Los Datos", vbExclamation, "Dato Incorrecto"
    TXTmax.Text = ""
    TXTmax.SetFocus
Else
    TXTuti.SetFocus
End If
End If
End Sub

Private Sub TXTmin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(TXTmin.Text) = False Then
    MsgBox "Este Campo es Numerico, Cambie la Condición", vbExclamation, "Dato Incorrecto"
    TXTmin.Text = ""
    TXTmin.SetFocus
ElseIf Val(TXTmin.Text) < 0 Then
    MsgBox " El Stock Minimo de un Repuesto, no puede ser Negativo, Cambie la Condición", vbExclamation, "Dato Incorrecto"
    TXTmin.Text = ""
    TXTmin.SetFocus
ElseIf Val(TXTmin.Text) > Val(TXTexi.Text) Then
    MsgBox " El Stock Minimo, no puede ser mayor a la Existencia. Verifique Los Datos", vbExclamation, "Dato Incorrecto"
    TXTmin.Text = ""
    TXTmin.SetFocus
Else
    TXTmax.SetFocus
End If
End If
End Sub

Private Sub TXTnre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TXTcos.SetFocus
End If
End Sub

Private Sub TXTrep_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsNumeric(TXTrep.Text) = False Then
        MsgBox "Este Campo es Numerico, Cambie la Condicion", vbExclamation, "Dato Incorrecto"
        TXTrep.Text = ""
        TXTrep.SetFocus
    ElseIf Val(TXTrep.Text) < 0 Then
        MsgBox " Este Codigo de Repuesto es Incorrecto, Cambie la Condicion", vbExclamation, "Dato Incorrecto"
        TXTrep.Text = ""
        TXTrep.SetFocus
Else
    CMDbus_Click
End If
End If
End Sub

Private Sub TXTuti_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(TXTuti.Text) = False Then
    MsgBox "Este Campo es Numerico, Cambie la Condicion", vbExclamation, "Dato Incorrecto"
    TXTuti.Text = ""
    TXTuti.SetFocus
ElseIf Val(TXTuti.Text) < 0 Then
    MsgBox " La Utilidad en Meses de un Repuesto no puede ser Negativo, Cambie la Condición", vbExclamation, "Dato Incorrecto"
    TXTuti.Text = ""
    TXTuti.SetFocus
End If
End If
End Sub
