VERSION 5.00
Begin VB.Form Productos 
   Caption         =   "Transporte EL SOL // Productos a Transportar"
   ClientHeight    =   7290
   ClientLeft      =   2895
   ClientTop       =   1770
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   Picture         =   "Productos.frx":0000
   ScaleHeight     =   7290
   ScaleWidth      =   10275
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
      Left            =   6240
      Picture         =   "Productos.frx":55AE
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton CMDbus 
      Height          =   495
      Left            =   5640
      Picture         =   "Productos.frx":5B38
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
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
      Left            =   7080
      Picture         =   "Productos.frx":60C2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4200
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
      Picture         =   "Productos.frx":664C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
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
      Picture         =   "Productos.frx":6BD6
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
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
      Picture         =   "Productos.frx":7160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox TXTpro 
      Height          =   375
      Left            =   3720
      MaxLength       =   30
      TabIndex        =   2
      Top             =   3480
      Width           =   2895
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
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   1
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "A Transportar"
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
      Left            =   6840
      TabIndex        =   11
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label LBLnom 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Nombre del Producto"
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
      Top             =   3480
      Width           =   2220
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
      TabIndex        =   3
      Top             =   2880
      Width           =   2145
   End
   Begin VB.Label LBLproductos 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Productos"
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
      TabIndex        =   0
      Top             =   480
      Width           =   2085
   End
End
Attribute VB_Name = "Productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDeli_Click()
Dim t As New ADODB.Recordset
Dim tp As New ADODB.Recordset
If TXTcod.Text = "" Then
MsgBox "Debe Escribir El Codigo del Producto", vbInformation, "Omision De Datos"
TXTcod.SetFocus
ElseIf TXTpro.Text = "" Then
MsgBox "No Deben Haber Campos Vacios", vbInformation, "Omision De Datos"
TXTpro.SetFocus
Else
tp.Open " select CodProd from Facturas where CodProd = '" & TXTcod.Text & "'", Conexion
    If tp.BOF Then
t.Open " select CodProd from ServiciosEfectuados where CodProd = '" & TXTcod.Text & "'", Conexion
    If t.BOF Then
    If MsgBox(" Desea Realmente Eliminar este Producto?", vbQuestion + vbYesNo, "Eliminacion de Datos") = vbYes Then
        Conexion.Execute "update TiposProducto set EstatusTpd = 'E' where CodPro = '" & TXTcod.Text & "'"
        MsgBox " Producto Eliminado Satisfactoriamente", vbInformation, "Eliminacion de Datos"
        Limpiar
        TXTcod.SetFocus
    End If
    Else
    MsgBox " ¡El Producto No Se Puede Eliminar! Esta Siendo Usado En Un Servicio  ", vbInformation, " Alerta"
    Limpiar
    TXTcod.SetFocus
        End If
    Else
        MsgBox " ¡El Producto No Se Puede Eliminar! Esta En Una Factura aun No Cancelada ", vbInformation, " Alerta"
        Limpiar
        TXTcod.SetFocus
    End If
End If
End Sub

Private Sub CMDinc_Click()
Dim t As New ADODB.Recordset
If TXTcod.Text = "" Then
    MsgBox " Debe Escribir El Codigo Del Producto", vbInformation, "Omision De Datos"
    TXTcod.SetFocus
    ElseIf TXTpro.Text = "" Then
        MsgBox " Debe Escribir El Nombre Del Producto ", vbInformation, "Omision De Datos"
        TXTpro.SetFocus
    Else
        t.Open "select NomPro from TiposProducto where NomPro = '" & Trim(UCase(TXTpro.Text)) & "' and Estatustpd = 'A'", Conexion
        If t.EOF Then
            Conexion.Execute " insert into TiposProducto (codPro, NomPro,EstatusTpd) Values ('" & Trim(TXTcod.Text) & "', '" & Trim(UCase(TXTpro.Text)) & "', 'A') "
            MsgBox " Producto Incluido Exitosamente ", vbInformation, "Inclusion de Datos"
            Limpiar
            TXTcod.SetFocus
        Else
           MsgBox " No Puede Incluirlo PorQue Ya Existe Un Producto Con Ese Nombre", vbInformation, "Duplicacion De Datos"
            Limpiar
            TXTcod.SetFocus
    End If
End If

End Sub

Private Sub CMDmod_Click()
Dim t As New ADODB.Recordset
If TXTcod.Text = "" Then
    MsgBox " Debe Escribir El Codigo Del Producto", vbInformation, "Omision De Datos"
    TXTcod.SetFocus
    ElseIf TXTpro.Text = "" Then
        MsgBox " Debe Escribir El Nombre Del Producto ", vbInformation, "Omision De Datos"
        TXTpro.SetFocus
    Else
    t.Open "select NomPro from TiposProducto where NomPro = '" & Trim(UCase(TXTpro.Text)) & "' and Estatustpd = 'A'", Conexion
        If t.EOF Then
        Conexion.Execute " update TiposProducto Set NomPro = '" & Trim(UCase(TXTpro.Text)) & "'where CodPro = '" & Trim(UCase(TXTcod.Text)) & "'"
        MsgBox " Producto Modificado Satisfactoriamente ", vbInformation, "Modificacion de Datos"
        Limpiar
        TXTcod.SetFocus
        Else
            MsgBox " No Puede Realizar Esa Modificacion PorQue Ya Existe Un Producto Con Ese Nombre", vbInformation, "Duplicacion De Datos"
            Limpiar
            TXTcod.SetFocus
        
    End If
    End If
End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub
Private Sub CMDbus_Click()
Dim t As New ADODB.Recordset
If TXTcod.Text = "" Then
    MsgBox " Debe Ingresar El Codigo del Producto", vbInformation, " Omision De Datos"
    TXTcod.SetFocus
Else
    t.Open "select NomPro, EstatusTpd from TiposProducto where CodPro = '" & TXTcod.Text & "'", Conexion
    If t.EOF Then
            If MsgBox(" Este Producto No Esta Incluido, Desea Incluirlo Ahora?", vbQuestion + vbYesNo, "Inclusion de Datos") = vbYes Then
          a = TXTcod.Text
          Limpiar
          CMDinc.Enabled = True
          TXTcod.Text = a
          TXTpro.SetFocus
        Else
          Limpiar
          TXTcod.SetFocus
       End If
    Else
        If t!EstatusTpd <> "A" Then
           If MsgBox(" El Producto Esta Inactivo, Desea Reactivarlo Ahora?", vbQuestion + vbYesNo, "Reactivacion de Datos") = vbYes Then
             activar
             MsgBox "El Producto fue Reactivado Satisfactoriamente", vbInformation, "Reactivacion de Datos"
             mostrar t
            Else
             Limpiar
             TXTcod.SetFocus
            End If
        Else
             mostrar t
End If
End If
End If
End Sub

Sub Limpiar()
TXTcod.Text = ""
TXTpro.Text = ""
CMDinc.Enabled = False
CMDmod.Enabled = False
CMDeli.Enabled = False
End Sub

Sub activar()
Conexion.Execute " update TiposProducto set Estatustpd = 'A' where CodPro = '" & TXTcod.Text & "'"
End Sub

Sub mostrar(t As ADODB.Recordset)
TXTpro.Text = t!nompro
CMDinc.Enabled = False
CMDmod.Enabled = True
CMDeli.Enabled = True
End Sub

Private Sub CMDlim_Click()
Limpiar
End Sub

Private Sub CMDsal_Click()
If MsgBox(" Desea Realmente Salir de 'Productos'?", vbQuestion + vbYesNo, "Salida de Pantalla") = vbYes Then
Unload Me
End If
End Sub

Private Sub TXTCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If IsNumeric(TXTcod.Text) = False Then
   MsgBox " Este Campo es de Tipo Numerico, Cambie la Condición", vbInformation, "Datos Incorrectos"
   TXTcod.Text = ""
   TXTcod.SetFocus
ElseIf Val(TXTcod.Text) < 0 Then
    MsgBox "Este Codigo es Incorrecto, Cambie la Condicion", vbInformation, "Dato Incorrecto"
    TXTcod.Text = ""
   TXTcod.SetFocus
Else
CMDbus_Click
End If
End If
End Sub


Private Sub TXTpro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If IsNumeric(TXTpro.Text) = True Then
   MsgBox " Debe Escribir Solamente Letras", vbInformation, "Datos Incorrectos"
   TXTpro.Text = ""
   TXTpro.SetFocus
End If
End If
End Sub

