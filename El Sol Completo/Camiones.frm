VERSION 5.00
Begin VB.Form Camiones 
   BackColor       =   &H80000005&
   Caption         =   "Transporte EL SOL // Camiones"
   ClientHeight    =   8010
   ClientLeft      =   2640
   ClientTop       =   1380
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Camiones.frx":0000
   ScaleHeight     =   8010
   ScaleWidth      =   11160
   Begin VB.TextBox TXTkma 
      Height          =   375
      Left            =   4080
      MaxLength       =   8
      TabIndex        =   26
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox TXTcap 
      Height          =   375
      Left            =   4080
      MaxLength       =   8
      TabIndex        =   23
      Top             =   6000
      Width           =   1455
   End
   Begin VB.TextBox TXTfad 
      Height          =   375
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   21
      Top             =   5280
      Width           =   2055
   End
   Begin VB.ComboBox CMBmar 
      Height          =   315
      Left            =   4080
      TabIndex        =   19
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox TXTmar 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      MaxLength       =   8
      TabIndex        =   18
      Top             =   4080
      Width           =   855
   End
   Begin VB.ComboBox CMBmod 
      Height          =   315
      Left            =   4080
      TabIndex        =   17
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox TXTmod 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      MaxLength       =   8
      TabIndex        =   16
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox TXTano 
      Height          =   375
      Left            =   7680
      MaxLength       =   4
      TabIndex        =   14
      Top             =   3480
      Width           =   735
   End
   Begin VB.ComboBox CMBcol 
      Height          =   315
      Left            =   4080
      TabIndex        =   13
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox TXTcol 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      MaxLength       =   8
      TabIndex        =   11
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton CMDlim 
      Caption         =   "&Limpiar"
      Height          =   495
      Left            =   6600
      Picture         =   "Camiones.frx":55AE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton CMDbus 
      Height          =   495
      Left            =   6000
      Picture         =   "Camiones.frx":5B38
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton CMDsal 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   7320
      Picture         =   "Camiones.frx":60C2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton CMDeli 
      Caption         =   "&Eliminar"
      Height          =   495
      Left            =   5400
      Picture         =   "Camiones.frx":664C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton CMDmod 
      Caption         =   "&Modificar"
      Height          =   495
      Left            =   3480
      Picture         =   "Camiones.frx":6BD6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton CMDinc 
      Caption         =   "&Incluir"
      Height          =   495
      Left            =   1560
      Picture         =   "Camiones.frx":7160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7320
      Width           =   1815
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
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   1
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label LBLkm 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Kms."
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
      Left            =   5640
      TabIndex        =   27
      Top             =   6840
      Width           =   420
   End
   Begin VB.Label LBLcap 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Kilometraje Actual"
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
      Left            =   1680
      TabIndex        =   25
      Top             =   6720
      Width           =   1890
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
      Left            =   5640
      TabIndex        =   24
      Top             =   6120
      Width           =   300
   End
   Begin VB.Label LBLcap 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Capacidad Máxima"
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
      Left            =   1680
      TabIndex        =   22
      Top             =   6000
      Width           =   2010
   End
   Begin VB.Label LBLfad 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fecha Adquisición"
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
      Left            =   1680
      TabIndex        =   20
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label LBLtel 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Año"
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
      Left            =   7200
      TabIndex        =   15
      Top             =   3600
      Width           =   420
   End
   Begin VB.Label LBLcol 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Color del Camion"
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
      Left            =   1680
      TabIndex        =   12
      Top             =   4680
      Width           =   1800
   End
   Begin VB.Label LBLmar 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Marca del Camion"
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
      Left            =   1680
      TabIndex        =   10
      Top             =   4080
      Width           =   1890
   End
   Begin VB.Label LBLmod 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Modelo del Camion"
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
      Left            =   1680
      TabIndex        =   9
      Top             =   3480
      Width           =   2025
   End
   Begin VB.Label LBLcod 
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
      Left            =   1680
      TabIndex        =   8
      Top             =   2760
      Width           =   1845
   End
   Begin VB.Label LBLcamiones 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Camiones"
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
      Left            =   8760
      TabIndex        =   0
      Top             =   720
      Width           =   2025
   End
End
Attribute VB_Name = "Camiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
LlenarCombo CMBmod, "select * from modelos", "modelo"
LlenarCombo CMBmar, "select * from marcas", "marca"
LlenarCombo CMBcol, "select * from colores", "color"
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub
Private Sub CMBcol_Click()
Dim t As New ADODB.Recordset
        t.Open "select codcol from colores where color = '" & CMBcol.Text & "'", Conexion
        TXTcol.Text = t!codcol
End Sub

Private Sub CMBmar_Click()
Dim t As New ADODB.Recordset
        t.Open "select codmarc from marcas where marca = '" & CMBmar.Text & "'", Conexion
        TXTmar.Text = t!codmarc
        CMBcol.SetFocus
End Sub

Private Sub CMBmod_Click()
Dim t As New ADODB.Recordset
        t.Open "select codmod from modelos where modelo = '" & CMBmod.Text & "'", Conexion
        TXTmod.Text = t!codmod
        TXTano.SetFocus
End Sub


Private Sub CMDbus_Click()
Dim t As New ADODB.Recordset
If TXTCod.Text = "" Then
    MsgBox "Debe Ingresar la Placa del Camion", vbInformation, "Omision de Datos"
    TXTCod.SetFocus
Else
    t.Open "select Placa,CodMod,Modelo,CodMarc,Marca,CodCol,Color,CargaMaximaKg,AnoVehic,FecAdquisicion,KmActual,EstatusCam from camiones,modelos,marcas,colores where Placa = '" & TXTCod.Text & "' and codmodelo = codmod and codcolor = codcol", Conexion
    If t.EOF Then
        If MsgBox("Este Camion No esta Registrado, Dese Registrarlo Ahora?", vbQuestion + vbYesNo, "Inclusion de Datos") = vbYes Then
            a = TXTCod.Text
            CMDlim_Click
            TXTCod.Text = a
            TXTCod.Enabled = False
            CMDinc.Enabled = True
        Else
            CMDlim_Click
            TXTCod.SetFocus
        End If
    Else
        If t!EstatusCam <> "A" Then
            If MsgBox(" El Camion Esta Inactivo, Desea Reactivarlo?", vbQuestion + vbYesNo, "Reactivacion de Datos") = vbYes Then
            activar
            MsgBox "El Camion fue Reactivado Satisfactoriamente", vbInformation, "Reactivacion de Datos"
            mostrar t
            TXTmar.Enabled = False
            CMBmar.Enabled = False
            CMBmod.Enabled = False
            TXTmod.Enabled = False
            TXTano.Enabled = False
            TXTfad.Enabled = False
            TXTkma.Enabled = False
        Else
            CMDlim_Click
            TXTCod.SetFocus
            End If
        Else
            mostrar t
            TXTmar.Enabled = False
            CMBmar.Enabled = False
            CMBmod.Enabled = False
            TXTmod.Enabled = False
            TXTano.Enabled = False
            TXTfad.Enabled = False
            TXTkma.Enabled = False
        End If
        End If
        End If
End Sub

Private Sub CMDeli_Click()
Dim t As New ADODB.Recordset
t.Open "Select * from ServiciosEfectuados where PlacaS = '" & TXTCod.Text & "' and EstatusSef = 'A'", Conexion
If Not t.EOF Then
    MsgBox "Usted No puede Eliminar este Camión, Porque tiene un(os) Servicio(s) por Efectuar", vbExclamation, "Eliminacion de Datos"
    Else
        t.Close
        t.Open "Select * from Camiones where Placa = '" & TXTCod.Text & "' and EstatusCam = 'A'", Conexion
        If Not t.EOF Then
               If MsgBox(" ¿ Desea Eliminar Este Camion?", vbQuestion + vbYesNo, "Reactivacion de Datos") = vbYes Then
            Eliminar
        Else
        CMDlim_Click
            TXTCod.SetFocus
        End If
End If
End If
End Sub

Sub Eliminar()
If TXTCod.Text = "" Then
    MsgBox "Debe escribir la Placa del Camión", vbExclamation, "Omision de Datos"
    TXTCod.SetFocus
Else
    Conexion.Execute " update camiones set estatuscam = 'E' where placa = '" & TXTCod.Text & "'"
    MsgBox "El Camion ha sido Eliminado Satisfactoriamente", vbExclamation, "Eliminacion de Datos"
    CMDlim_Click
End If
End Sub
Private Sub CMDinc_Click()
Incluir
End Sub
Sub Incluir()
If TXTCod.Text = "" Then
        MsgBox "Debe escribir un codigo", vbInformation, "Omision de Datos"
        TXTCod.SetFocus
    ElseIf CMBmod.Text = "" Then
        MsgBox "Debe escribir un modelo", vbInformation, "Omision de Datos"
        CMBmod.SetFocus
    ElseIf TXTmod.Text = "" Then
        MsgBox "Debe escribir el modelo", vbInformation, "Omision de Datos"
        TXTmod.SetFocus
    ElseIf CMBmar.Text = "" Then
        MsgBox "Debe escribir una marca", vbInformation, "Omision de Datos"
        CMBmar.SetFocus
    ElseIf TXTmar.Text = "" Then
        MsgBox "Debe escribir una marca", vbInformation, "Omision de Datos"
        TXTmar.SetFocus
    ElseIf CMBcol.Text = "" Then
        MsgBox "Debe escribir un color", vbInformation, "Omision de Datos"
        CMBcol.SetFocus
    ElseIf TXTcol.Text = "" Then
        MsgBox "Debe escribir un color", vbInformation, "Omision de Datos"
        TXTcol.SetFocus
    ElseIf TXTano.Text = "" Then
        MsgBox "Debe escribir el año del vehiculo", vbInformation, "Omision de Datos"
        TXTano.SetFocus
    ElseIf TXTcap.Text = "" Then
        MsgBox "Debe escribir la capacidad del vehiculo", vbInformation, "Omision de Datos"
        TXTano.SetFocus
    ElseIf TXTfad.Text = "" Then
        MsgBox "Debe escribir la Fecha de adquisicion", vbInformation, "Omision de Datos"
        TXTfad.SetFocus
    ElseIf TXTkma.Text = "" Then
        MsgBox "Debe escribir el kilometraje", vbInformation, "Omision de Datos"
        TXTkma.SetFocus
    Else
        Conexion.Execute "insert into camiones (placa,codmodelo,codcolor,codmarca,CargaMaximaKg,AnoVehic,FecAdquisicion,KmActual,EstatusCam) values('" & Trim(UCase(TXTCod.Text)) & "','" & Trim(UCase(TXTmod.Text)) & "', '" & Trim(UCase(TXTcol.Text)) & "','" & Trim(UCase(TXTmar.Text)) & "','" & Trim(UCase(TXTcap.Text)) & "','" & Trim(UCase(TXTano.Text)) & "','" & Trim(UCase(TXTfad.Text)) & "', '" & Trim(UCase(TXTkma.Text)) & "', 'A')"
        MsgBox " El Camion ha sido Incluido Exitosamente", vbInformation, "Inclusion de Datos"
        CMDlim_Click
    End If
End Sub

Private Sub CMDlim_Click()
TXTCod.Text = ""
CMBmod.Text = ""
TXTmod.Text = ""
CMBmar.Text = ""
TXTmar.Text = ""
TXTano.Text = ""
CMBcol.Text = ""
TXTcol.Text = ""
TXTcap.Text = ""
TXTfad.Text = ""
TXTkma.Text = ""
TXTmar.Enabled = True
CMBmar.Enabled = True
CMBmod.Enabled = True
TXTmod.Enabled = True
TXTano.Enabled = True
TXTfad.Enabled = True
TXTkma.Enabled = True
CMDinc.Enabled = False
CMDmod.Enabled = False
CMDeli.Enabled = False
TXTCod.Enabled = True
TXTCod.SetFocus
End Sub

Private Sub CMDmod_Click()
If TXTCod.Text = "" Then
        MsgBox "Debe escribir un codigo", vbInformation, "Omision de Datos"
        TXTCod.SetFocus
    ElseIf CMBmod.Text = "" Then
        MsgBox "Debe escribir un modelo", vbInformation, "Omision de Datos"
        CMBmod.SetFocus
    ElseIf TXTmod.Text = "" Then
        MsgBox "Debe escribir el modelo", vbInformation, "Omision de Datos"
        TXTmod.SetFocus
    ElseIf CMBmar.Text = "" Then
        MsgBox "Debe escribir una marca", vbInformation, "Omision de Datos"
        CMBmar.SetFocus
    ElseIf TXTmar.Text = "" Then
        MsgBox "Debe escribir una marca", vbInformation, "Omision de Datos"
        TXTmar.SetFocus
    ElseIf CMBcol.Text = "" Then
        MsgBox "Debe escribir un color", vbInformation, "Omision de Datos"
        CMBcol.SetFocus
    ElseIf TXTcol.Text = "" Then
        MsgBox "Debe escribir un color", vbInformation, "Omision de Datos"
        TXTcol.SetFocus
    ElseIf TXTano.Text = "" Then
        MsgBox "Debe escribir el año del vehiculo", vbInformation, "Omision de Datos"
        TXTano.SetFocus
    ElseIf TXTcap.Text = "" Then
        MsgBox "Debe escribir la capacidad del vehiculo", vbInformation, "Omision de Datos"
        TXTano.SetFocus
    ElseIf TXTfad.Text = "" Then
        MsgBox "Debe escribir la Fecha de adquisicion", vbInformation, "Omision de Datos"
        TXTfad.SetFocus
    ElseIf TXTkma.Text = "" Then
        MsgBox "Debe escribir el kilometraje", vbInformation, "Omision de Datos"
        TXTkma.SetFocus
    Else
        Conexion.Execute "update camiones set codmodelo = '" & Trim(UCase(TXTmod.Text)) & "',Codmarca = '" & Trim(UCase(TXTmar.Text)) & "', codcolor = '" & Trim(UCase(TXTcol.Text)) & "',CargaMaximaKg = '" & Trim(UCase(TXTcap.Text)) & "',anovehic = '" & Trim(UCase(TXTano.Text)) & "',FecAdquisicion = '" & Trim(UCase(TXTfad.Text)) & "',KmActual = '" & Trim(UCase(TXTkma.Text)) & "' where placa = '" & Trim(UCase(TXTCod.Text)) & "'"
        MsgBox " El Registro fue Modificado Exitosamente", vbInformation, "Modificacion de Datos"
        CMDlim_Click
        TXTCod.SetFocus
    End If
End Sub

Private Sub CMDsal_Click()
If MsgBox(" Desea Realmente Salir de 'Camiones'?", vbQuestion + vbYesNo, "Salida de Pantalla") = vbYes Then
Unload Me
End If
End Sub
Sub activar()
    Conexion.Execute " update camiones set EstatusCam = 'A' where placa = '" & TXTCod.Text & "'"
End Sub

Sub mostrar(t As ADODB.Recordset)
    'TexCed.Text = t!Cedula
    TXTCod.Text = t!placa
    CMBmod.Text = t!modelo
    TXTmod.Text = t!codmod
    CMBmar.Text = t!marca
    TXTmar.Text = t!codmarc
    TXTano.Text = t!anovehic
    CMBcol.Text = t!Color
    TXTcol.Text = t!codcol
    TXTcap.Text = t!CargaMaximaKg
    TXTfad.Text = t!fecadquisicion
    TXTkma.Text = t!kmactual
    CMDinc.Enabled = False
    CMDmod.Enabled = True
    CMDeli.Enabled = True
End Sub



Private Sub TXTano_Change()
If Val(TXTano.Text) < 0 Then
    MsgBox "El Año del Vehiculo no puede ser Negativo, Cambie la Condición", vbExclamation, "Dato Incorrecto"
    TXTano.Text = ""
    TXTano.SetFocus
End If
End Sub

Private Sub TXTano_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(TXTano.Text) = False Then
MsgBox "Este Campo es Numerico, Cambie la Condición", vbExclamation, "Datos Incorrectos"
TXTano.Text = ""
TXTano.SetFocus
ElseIf Val(TXTano.Text) < 0 Then
    MsgBox "El Año del Vehiculo no puede ser Negativo, Cambie la Condición", vbExclamation, "Dato Incorrecto"
    TXTano.Text = ""
    TXTano.SetFocus
    Else
    CMBmar.SetFocus
End If
End If
End Sub

Private Sub TXTcap_Change()
If Val(TXTcap.Text) < 0 Then
        MsgBox " La Capacidad del Camion en Kgs. No puede ser Negativo, Cambie la Condición", vbExclamation, "Datos Incorrectos"
        TXTcap.Text = ""
        TXTcap.SetFocus
    End If
End Sub

Private Sub TXTcap_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(TXTcap.Text) = False Then
    MsgBox " Este Campo es Numerico, Cambie la Condición", vbExclamation, "Datos Incorrectos"
    TXTcap.Text = ""
    TXTcap.SetFocus
Else
    If Val(TXTcap.Text) < 0 Then
        MsgBox " La Capacidad del Camion en Kgs. No puede ser Negativo, Cambie la Condición", vbExclamation, "Datos Incorrectos"
        TXTcap.Text = ""
        TXTcap.SetFocus
    Else
        TXTkma.SetFocus
    End If
End If
End If
End Sub

Private Sub TXTCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CMDbus_Click
End If
End Sub

Private Sub TXTcol_KeyPress(KeyAscii As Integer)
Dim t As New ADODB.Recordset
If KeyAscii = 13 Then
    If IsNumeric(TXTcol.Text) = False Then
        MsgBox " Este Campo es Numerico, Cambie la Condición", vbInformation, "Datos Incorrectos"
        TXTcol.Text = ""
        CMBcol.Text = ""
        TXTcol.SetFocus
    Else
        t.Open "select color from colores where codcol = '" & TXTcol.Text & "'", Conexion
        CMBcol.Text = t!Color
         
    End If
    End If
End Sub

Private Sub TXTfad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsDate(TXTfad.Text) = False Then
    MsgBox "Este Campo es de Tipo Fecha, Cambie la Condición", vbInformation, "Datos Incorrectos"
    TXTfad.Text = ""
    TXTfad.SetFocus
ElseIf Year(TXTfad.Text) < TXTano.Text Then
    MsgBox " El Año de Adquisición del Camión, no puede ser menor que el Año del Mismo, Cambie la Condición", vbExclamation, "Datos Incorrectos"
    TXTfad.Text = ""
    TXTfad.SetFocus
Else
    TXTcap.SetFocus
End If
End If
End Sub


Private Sub TXTkma_Change()
If Val(TXTkma.Text) < 0 Then
        MsgBox " El Kilometraje actual del Camión en Kms. No puede ser Negativo, Cambie la Condición", vbExclamation, "Datos Incorrectos"
        TXTkma.Text = ""
        TXTkma.SetFocus
    End If
End Sub

Private Sub TXTkma_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(TXTkma.Text) = False Then
    MsgBox " Este Campo es Numerico, Cambie la Condición", vbExclamation, "Datos Incorrectos"
    TXTkma.Text = ""
    TXTkma.SetFocus
Else
    If Val(TXTkma.Text) < 0 Then
        MsgBox " El Kilometraje actual del Camión en Kms. No puede ser Negativo, Cambie la Condición", vbExclamation, "Datos Incorrectos"
        TXTkma.Text = ""
        TXTkma.SetFocus
    Else
    Incluir
    End If
End If
End If
End Sub

Private Sub TXTmar_KeyPress(KeyAscii As Integer)
Dim t As New ADODB.Recordset
If KeyAscii = 13 Then
    If IsNumeric(TXTmar.Text) = False Then
        MsgBox " Este Campo es Numerico, Cambie la Condición", vbInformation, "Datos Incorrectos"
        TXTmar.Text = ""
        CMBmar.Text = ""
        TXTmar.SetFocus
    Else
        t.Open "select marca from marcas where codmarc = '" & TXTmar.Text & "'", Conexion
        CMBmar.Text = t!marca
         
    End If
    End If
End Sub

Private Sub TXTmod_KeyPress(KeyAscii As Integer)
Dim t As New ADODB.Recordset
If KeyAscii = 13 Then
    If IsNumeric(TXTmod.Text) = False Then
        MsgBox " Este Campo es Numerico, Cambie la Condición", vbInformation, "Datos Incorrectos"
        TXTmod.Text = ""
        CMBmod.Text = ""
        TXTmod.SetFocus
    Else
        t.Open "select modelo from modelos where codmod = '" & TXTmod.Text & "'", Conexion
        CMBmod.Text = t!modelo
    End If
    End If
End Sub
