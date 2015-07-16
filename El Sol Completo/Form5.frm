VERSION 5.00
Begin VB.Form Clientes 
   BackColor       =   &H80000005&
   Caption         =   "Transporte EL SOL // Clientes"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10485
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   7230
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CMBtel 
      Height          =   315
      Left            =   3960
      TabIndex        =   7
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox TXTtel 
      Height          =   375
      Left            =   5160
      MaxLength       =   7
      TabIndex        =   8
      Top             =   4920
      Width           =   1695
   End
   Begin VB.ComboBox CMBciu 
      Height          =   315
      Left            =   3960
      TabIndex        =   5
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox TXTciu 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      MaxLength       =   8
      TabIndex        =   6
      Top             =   4320
      Width           =   735
   End
   Begin VB.ComboBox CMBcod 
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
      ItemData        =   "Form5.frx":55AE
      Left            =   3960
      List            =   "Form5.frx":55B8
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox TXTcod 
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
      Left            =   4920
      MaxLength       =   8
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox TXTnom 
      Height          =   375
      Left            =   3960
      MaxLength       =   30
      TabIndex        =   3
      Top             =   3120
      Width           =   2895
   End
   Begin VB.TextBox TXTdir 
      Height          =   375
      Left            =   3960
      MaxLength       =   30
      TabIndex        =   4
      Top             =   3720
      Width           =   2895
   End
   Begin VB.CommandButton CMDinc 
      Caption         =   "&Incluir"
      Height          =   495
      Left            =   1560
      Picture         =   "Form5.frx":55C6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton CMDmod 
      Caption         =   "&Modificar"
      Height          =   495
      Left            =   3480
      Picture         =   "Form5.frx":5B50
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton CMDeli 
      Caption         =   "&Eliminar"
      Height          =   495
      Left            =   5400
      Picture         =   "Form5.frx":60DA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton CMDsal 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   7320
      Picture         =   "Form5.frx":6664
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton CMDbus 
      Height          =   495
      Left            =   7080
      Picture         =   "Form5.frx":6BEE
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton CMDlim 
      Caption         =   "&Limpiar"
      Height          =   495
      Left            =   7680
      Picture         =   "Form5.frx":7178
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label LBLtel 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Telefono"
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
      Left            =   1560
      TabIndex        =   19
      Top             =   4920
      Width           =   945
   End
   Begin VB.Label LBLciu 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Ciudad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1560
      TabIndex        =   18
      Top             =   4320
      Width           =   750
   End
   Begin VB.Label LBLced 
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
      Left            =   1560
      TabIndex        =   17
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label LBLnom 
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
      Left            =   1560
      TabIndex        =   16
      Top             =   3120
      Width           =   2010
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Direccion del Cliente"
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
      Left            =   1560
      TabIndex        =   13
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label LBLclientes 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Clientes"
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
      Left            =   8400
      TabIndex        =   0
      Top             =   720
      Width           =   1650
   End
End
Attribute VB_Name = "Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMBtel_Click()
Dim t As New ADODB.Recordset
t.Open "select Ciudades, CodCiudad from ciudades where codarea = '" & CMBtel.Text & "'", Conexion
  If Not t.EOF Then
     TXTciu.Text = t!CodCiudad
     CMBciu.Text = t!Ciudades
  Else
     If t.EOF Then
            MsgBox "El Codigo No Existe, por favor rectifique", vbExclamation, "Omision De Datos"
            CMBciu.Text = ""
            TXTciu.Text = ""
    End If
  End If
End Sub

Private Sub CMDsal_Click()
If MsgBox("Desea Realmente Salir de 'Clientes!'?", vbQuet + vbYesNo, "Salida de Programa") = vbYes Then
Unload Me
End If
End Sub

Private Sub Form_Load()
LlenarCombo CMBciu, " select * from Ciudades ", "Ciudades"
LlenarCombo CMBtel, " select * from Ciudades", "CodArea"
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Private Sub CMBciu_Click()
Dim t As New ADODB.Recordset
t.Open "select CodCiudad, CodArea from ciudades where ciudades = '" & CMBciu.Text & "'", Conexion
  If Not t.EOF Then
     TXTciu.Text = t!CodCiudad
     CMBtel.Text = t!CodArea
     CMBtel.Enabled = False
     Else
     If t.EOF Then
            MsgBox "El Codigo No Existe, por favor rectifique", vbExclamation, "Omision De Datos"
            CMBciu.Text = ""
            CMBtel.Text = ""
            TXTtel.SetFocus
  End If
  End If
End Sub
Private Sub CMDbus_Click()
Dim t As New ADODB.Recordset
If TXTCod.Text = "" Then
    MsgBox " Debe Ingresar Una Cedula", vbInformation, " Omision De Datos"
    TXTCod.SetFocus
Else
    TXTCod.ForeColor = &HFF&
    t.Open "select NomCli,DirCli,CodCiudad,Ciudades,TipoP, codarea,TelCli, EstatusCli from Clientes, Ciudades, TelCliente where CodCli = '" & TXTCod.Text & "' and CodCiu = CodCiudad", Conexion
    If t.EOF Then
           If MsgBox("El Cliente No Existe, Desea Registrarlo?", vbQuestion + vbYesNo, "Inclusion de Datos") = vbYes Then
          a = TXTCod.Text
          CMDinc.Enabled = True
          txtnom.SetFocus
        Else
          Limpiar
          TXTCod.SetFocus
       End If
    Else
        If t!EstatusCli <> "A" Then
           If MsgBox(" El Cliente Esta Inactivo, Desea Reactivarlo?", vbQuestion + vbYesNo, "Reactivacion de Datos") = vbYes Then
             activar
             MsgBox "El Cliente Fue Reactivado Satisfactoriamente", vbInformation, "Reactivacion de Datos"""
             mostrar t
            Else
             Limpiar
             TXTCod.SetFocus
            End If
        Else
             mostrar t
End If
End If
End If
    End Sub
Sub mostrar(t As ADODB.Recordset)
    Dim AUX As New ADODB.Recordset
    txtnom.Text = t!nomcli
    TXTdir.Text = t!DirCli
    TXTciu.Text = t!CodCiudad
    CMBciu.Text = t!Ciudades
    CMBtel.Text = t!CodArea
    CMBcod.Text = t!TipoP
    CMDinc.Enabled = False
    CMDmod.Enabled = True
    CMDeli.Enabled = True
    CMBcod.Enabled = False
    TXTCod.Enabled = False
    txtnom.Enabled = False
AUX.Open " select * from telCliente where CodClit = '" & TXTCod.Text & "'", Conexion
TXTtel.Text = AUX!TelCLi

End Sub
Sub Limpiar()
TXTCod.Text = ""
txtnom.Text = ""
TXTdir.Text = ""
TXTciu.Text = ""
CMBcod.Text = ""
CMBciu.Text = ""
TXTtel.Text = ""
CMBtel.Text = ""
CMBcod.Enabled = True
TXTCod.Enabled = True
CMBtel.Enabled = True
txtnom.Enabled = True
TXTCod.SetFocus
End Sub
Sub activar()
Conexion.Execute " update Clientes set EstatusCli = 'A' where CodCli = '" & TXTCod.Text & "'"
End Sub

Private Sub CMDeli_Click()
Dim t As New ADODB.Recordset
Dim tp As New ADODB.Recordset
If TXTCod.Text = "" Then
MsgBox "Debe Escribir Un codigo", vbInformation, "Omision De Datos"
TXTCod.SetFocus
Else
t.Open " select CodCliente from ServiciosEfectuados where CodCliente = '" & TXTCod.Text & "'", Conexion
If t.BOF Then
    tp.Open " select CodClientef from Facturas where CodClientef = '" & TXTCod.Text & "'", Conexion
    If tp.BOF Then
        If MsgBox(" Desea Realmente Eliminar este Cliente?", vbQuestion + vbYesNo, "Eliminacion de Datos") = vbYes Then
        Conexion.Execute "update Clientes set EstatusCli = 'E' where CodCli = '" & TXTCod.Text & "'"
        MsgBox "El Cliente Fue ELiminado Satisfactoriamente", vbInformation, "Eliminacion de Datos"
        Limpiar
        TXTCod.SetFocus
        End If
    Else
    MsgBox "Usted No puede Eliminar este Cliente, porque aun tiene una(s) Factura(s) Pendiente(s) por Cancelar", vbExclamation, "Eliminacion de Datos"
    Limpiar
    TXTCod.SetFocus
        End If
    Else
        MsgBox "Usted No puede Eliminar este Cliente, porque aun tiene un(os) Servicio(s) Pendiente(s) por ser Efectuados", vbExclamation, "Eliminacion de Datos"
        Limpiar
        TXTCod.SetFocus
    End If
End If
End Sub

Private Sub CMDinc_Click()
If TXTCod.Text = "" Then
    MsgBox " Debe Escribir Un Codigo", vbInformation, "Omision De Datos"
    TXTCod.SetFocus
    ElseIf txtnom.Text = "" Then
        MsgBox " Debe Escribir Un Nombre ", vbInformation, "Omision De Datos"
        txtnom.SetFocus
    ElseIf TXTdir.Text = "" Then
        MsgBox " Debe Escribir Una Direccion ", vbInformation, "Omision De Datos"
        TXTdir.SetFocus
    ElseIf TXTciu.Text = "" Then
        MsgBox " Debe Escribir Un Codigo De Ciudad ", vbInformation, "Omision De Datos"
        TXTciu.SetFocus
    ElseIf CMBciu.Text = "" Then
        MsgBox " Debe Escribir Una Ciudad", vbInformation, "Omision De Datos"
        CMBciu.SetFocus
  
    Else
         Conexion.Execute "insert into   Clientes (CodCli ,NomCli, DirCli,TipoP,CodCiu,EstatusCli) values('" & Trim(TXTCod.Text) & "', '" & Trim(UCase(txtnom.Text)) & "','" & Trim(UCase(TXTdir.Text)) & "', '" & Trim(UCase(CMBcod.Text)) & "','" & Trim(UCase(TXTciu.Text)) & "', 'A')"
         Conexion.Execute " insert into TelCliente (codClit,TelCli,EstatusTel) values ('" & TXTCod.Text & " ' , " & TXTtel.Text & " , 'A')"
        MsgBox " Cliente Incluido Satisfactoriamente ", vbInformation, "Inclusion de Datos"
        Limpiar
        TXTCod.SetFocus
    End If
End Sub

Private Sub CMDlim_Click()
Limpiar
End Sub

Private Sub CMDmod_Click()
If TXTCod.Text = "" Then
    MsgBox " Debe Escribir Un Codigo", vbInformation, "Omision De Datos"
    TXTCod.SetFocus
    ElseIf txtnom.Text = "" Then
        MsgBox " Debe Escribir Un Nombre ", vbInformation, "Omision De Datos"
        txtnom.SetFocus
    ElseIf TXTdir.Text = "" Then
        MsgBox " Debe Escribir Una Direccion ", vbInformation, "Omision De Datos"
        TXTdir.SetFocus
    ElseIf TXTciu.Text = "" Then
        MsgBox " Debe Escribir Un Codigo De Ciudad ", vbInformation, "Omision De Datos"
        TXTciu.SetFocus
    ElseIf CMBciu.Text = "" Then
        MsgBox " Debe Escribir Una Ciudad", vbInformation, "Omision De Datos"
        CMBciu.SetFocus
     CMBciu.SetFocus
    Else
        Conexion.Execute " update Clientes Set NomCli = '" & Trim(UCase(txtnom.Text)) & "', DirCli = '" & Trim(UCase(TXTdir.Text)) & "', CodCiu= '" & Trim(UCase(TXTciu.Text)) & "' where CodCli = '" & Trim(UCase(TXTCod.Text)) & "'"
        Conexion.Execute " update  telCliente set  TelCli = '" & TXTtel.Text & " 'where codclit = '" & TXTCod.Text & " '"
        MsgBox " Cliente Modificado Satisfactoriamente", vbInformation, "Modificacion de Datos"
        Limpiar
        TXTCod.SetFocus
    End If
End Sub


Private Sub TXTciu_DblClick()
Dim t As New ADODB.Recordset
If TXTciu.Text = "" Then
        MsgBox "Debe Escribir El Codigo de la Ciudad", vbExclamation, "Omision de Datos"
     Else
     If IsNumeric(TXTciu.Text) = False Then
   MsgBox " Debe Escribir Solamente Numeros", vbInformation, "Datos Incorrectos"
   TXTciu.Text = ""
   TXTciu.SetFocus
Else
        TXTciu.Text = Format(TXTciu.Text, "000")
        t.Open "select Ciudad from Ciudades where CodCiudad = '" & TXTciu.Text & "'", Conexion
        If t.EOF Then
           MsgBox "La Ciudad No Esta Registrada", vbExclamation, "Campo no Registrado"
        Else
           CMBciu.Text = t!Ciudad
           End If
        End If
     End If
End Sub

Private Sub TXTciu_KeyPress(KeyAscii As Integer)
Dim t As New ADODB.Recordset
  If KeyAscii = 13 Then
     If TXTciu.Text = "" Then
        MsgBox "Debe Escribir El Codigo de la Ciudad", vbExclamation, "Omision de Datos"
     Else
     If IsNumeric(TXTciu.Text) = False Then
   MsgBox " Debe Escribir Solamente Numeros", vbInformation, "Datos Incorrectos"
   TXTciu.Text = ""
   TXTciu.SetFocus
Else
        TXTciu.Text = Format(TXTciu.Text, "00")
        t.Open "select Ciudades from Ciudades where CodCiudad = '" & TXTciu.Text & "'", Conexion
        If t.EOF Then
           MsgBox "La Ciudad No Esta Registrada", vbExclamation, "Campo no Registrado"
        Else
           CMBciu.Text = t!Ciudades
           End If
        End If
     End If
  End If
End Sub

Private Sub TXTCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If IsNumeric(TXTCod.Text) = False Then
   MsgBox " Este Campo es Numerico, Cambie la Condicion", vbInformation, "Datos Incorrectos"
   TXTCod.Text = ""
Else
CMDbus_Click
End If
End If
End Sub
Private Sub TXTdir_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 CMBciu.SetFocus
End If
End Sub
Private Sub txtnom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If IsNumeric(txtnom.Text) = True Then
   MsgBox " Debe Escribir Solamente Letras, Cambie la Condicion", vbInformation, "Datos Incorrectos"
   txtnom.Text = ""
   txtnom.SetFocus
   Else
     TXTdir.SetFocus
End If
End If

End Sub

