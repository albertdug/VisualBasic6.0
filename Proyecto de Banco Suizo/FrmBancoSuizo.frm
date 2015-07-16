VERSION 5.00
Begin VB.Form FrmBancoSuizo 
   BackColor       =   &H00C00000&
   Caption         =   "Banco Suizo"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   FillColor       =   &H00808080&
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtAbono 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      TabIndex        =   20
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox TxtMontoARetirar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      TabIndex        =   18
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox TxtSaldo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      TabIndex        =   15
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox TxtDireccion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      TabIndex        =   14
      Top             =   4080
      Width           =   4095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "A:\Proyecto Final\BdBancoSuizo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "T_Ahorros"
      Top             =   8160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton CmbReporte 
      Caption         =   "Reporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton CmbSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton CmbModificar 
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton CmbBuscar 
      BackColor       =   &H8000000D&
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton CmbNuevo 
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3600
      TabIndex        =   7
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton CmbGuardar 
      BackColor       =   &H00C00000&
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      MaskColor       =   &H00C00000&
      TabIndex        =   6
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox TxtNombre 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      TabIndex        =   2
      Top             =   3600
      Width           =   4095
   End
   Begin VB.TextBox TxtCedula 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox TxtNumero 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7080
      TabIndex        =   24
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label LblHora 
      BackColor       =   &H00C00000&
      Caption         =   "Hora:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5880
      TabIndex        =   23
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7080
      TabIndex        =   22
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label LblFecha 
      BackColor       =   &H00C00000&
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5880
      TabIndex        =   21
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label LblAbono 
      BackColor       =   &H00C00000&
      Caption         =   "Abono:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   840
      TabIndex        =   19
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label LblMontARetirar 
      BackColor       =   &H00C00000&
      Caption         =   "Monto a retirar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label LblCtaAhorro 
      BackColor       =   &H00C00000&
      Caption         =   "CUENTA DE AHORROS"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1920
      TabIndex        =   16
      Top             =   600
      Width           =   5415
   End
   Begin VB.Label LblSaldo 
      BackColor       =   &H00C00000&
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label LblDireccion 
      BackColor       =   &H00C00000&
      Caption         =   "Direccion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label LblNombre 
      BackColor       =   &H00C00000&
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label LblCedula 
      BackColor       =   &H00C00000&
      Caption         =   "Cedula del Ahorrista:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label LblNumero 
      BackColor       =   &H00C00000&
      Caption         =   "Numero:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   2640
      Width           =   1935
   End
End
Attribute VB_Name = "FrmBancoSuizo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbBuscar_Click()
If TxtNumero.Text = "" Then
 MsgBox " Debe tipiar un Numero", vbInformation
 TxtNumero.SetFocus
 Else
   Data1.Recordset.FindFirst (" numero= '") & TxtNumero & "'"
  If Data1.Recordset.NoMatch Then
    MsgBox "No se encontró el registro", vbCritical
    TxtNumero = ""
     TxtCedula = ""
      TxtNombre = ""
       TxtDireccion = ""
        TxtSaldo = ""
         TxtMontoARetirar = ""
          TxtAbono = ""
           TxtNumero.SetFocus
       Else
    TxtCedula = Data1.Recordset("Cedula")
    TxtNombre = Data1.Recordset("Nombre")
    TxtDireccion = Data1.Recordset("Direccion")
    TxtSaldo = Data1.Recordset("Saldo")
    TxtMontoARetirar = Data1.Recordset("MontoARetirar")
    TxtAbono = Data1.Recordset("Abono")
   End If
End If
End Sub

Private Sub CmbGuardar_Click()
If TxtNumero = "" Or TxtCedula = "" Then
MsgBox "No debe dejar campos en blanco", vbInformation
    TxtNumero = ""
    TxtCedula = ""
    TxtNombre = ""
    TxtDireccion = ""
    TxtSaldo = ""
    TxtMontoARetirar = ""
    TxtAbono = ""
    TxtNumero.SetFocus
    Else
    Data1.Recordset.FindFirst ("Numero='") & TxtNumero & "'"
      If Data1.Recordset.NoMatch Then
       Data1.Recordset.AddNew
        Data1.Recordset("Numero") = TxtNumero
         Data1.Recordset("Cedula") = TxtCedula
          Data1.Recordset("Nombre") = TxtNombre
           Data1.Recordset("Direccion") = TxtDireccion
            Data1.Recordset("Saldo") = TxtSaldo
             Data1.Recordset("MontoARetirar") = TxtMontoARetirar
              Data1.Recordset("Abono") = TxtAbono
               Data1.Recordset.Update
         MsgBox "El Registro fue Guardado", vbInformation
           TxtNumero = ""
            TxtCedula = ""
             TxtNombre = ""
              TxtDireccion = ""
               TxtSaldo = ""
                TxtMontoARetirar = ""
                 TxtAbono = ""
                  TxtNumero.SetFocus
                Else
           MsgBox "El Registro ya existe", vbCritical
           TxtNumero = ""
            TxtCedula = ""
             TxtNombre = ""
              TxtDireccion = ""
               TxtSaldo = ""
                TxtMontoARetirar = ""
                 TxtAbono = ""
                  TxtNumero.SetFocus
      End If
End If
End Sub

Private Sub CmbModificar_Click()
If TxtNumero = "" Then
  MsgBox "Debe escribir un numero", vbCritical
   TxtNumero.Text = ""
    TxtCedula.Text = ""
     TxtNombre.Text = ""
      TxtDireccion.Text = ""
       TxtSaldo.Text = ""
        TxtMontoARetirar.Text = ""
         TxtAbono.Text = ""
          TxtNumero.SetFocus
          Else
           Data1.Recordset.FindFirst ("Numero='") & TxtNumero & "'"
           If Not Data1.Recordset.NoMatch Then
            Data1.Recordset.Edit
             Data1.Recordset("Numero") = TxtNumero
              Data1.Recordset("Cedula") = TxtCedula
               Data1.Recordset("Nombre") = TxtNombre
                Data1.Recordset("Direccion") = TxtDireccion
                 Data1.Recordset("Saldo") = TxtSaldo
                  Data1.Recordset("MontoARetirar") = TxtMontoARetirar
                   Data1.Recordset("Abono") = TxtAbono
                    Data1.Recordset.Update
                    MsgBox "El Registro ha sido modificado", vbInformation
                    TxtNumero.Text = ""
                     TxtCedula.Text = ""
                      TxtNombre.Text = ""
                       TxtDireccion.Text = ""
                        TxtSaldo.Text = ""
                         TxtMontoARetirar.Text = ""
                          TxtAbono.Text = ""
                           TxtNumero.SetFocus
                           Else
                           MsgBox "El Registro No Existe", vbCritical
                            TxtNumero.Text = ""
                             TxtCedula.Text = ""
                              TxtNombre.Text = ""
                               TxtDireccion.Text = ""
                                TxtSaldo.Text = ""
                                 TxtMontoARetirar.Text = ""
                                  TxtAbono.Text = ""
                                   TxtNumero.SetFocus
            End If
    End If
End Sub

Private Sub CmbNuevo_Click()
TxtNumero.Text = ""
TxtCedula.Text = ""
TxtNombre.Text = ""
TxtDireccion.Text = ""
TxtSaldo.Text = ""
TxtMontoARetirar.Text = ""
TxtAbono.Text = ""
TxtNumero.SetFocus
End Sub

Private Sub CmbReporte_Click()
DataReport1.Show
End Sub

Private Sub CmbSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label1 = Date
Label3 = Time
End Sub

Private Sub MskNumero_Change()

End Sub

