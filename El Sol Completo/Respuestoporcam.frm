VERSION 5.00
Begin VB.Form Repuestoporcam 
   BackColor       =   &H80000005&
   Caption         =   "Transporte EL SOL // Repuestos por Camión"
   ClientHeight    =   8685
   ClientLeft      =   2385
   ClientTop       =   870
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   Picture         =   "Respuestoporcam.frx":0000
   ScaleHeight     =   8685
   ScaleWidth      =   10500
   Begin VB.CommandButton Command1 
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
      Picture         =   "Respuestoporcam.frx":55AE
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5760
      Width           =   1815
   End
   Begin VB.TextBox TXTcrc 
      Enabled         =   0   'False
      Height          =   405
      Left            =   6000
      TabIndex        =   15
      Top             =   3240
      Width           =   855
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
      Left            =   6000
      Picture         =   "Respuestoporcam.frx":5B38
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton CMDgua 
      Caption         =   "&Guardar"
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
      Picture         =   "Respuestoporcam.frx":60C2
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5760
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
      Picture         =   "Respuestoporcam.frx":664C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5760
      Width           =   1815
   End
   Begin VB.TextBox TXTcar 
      Height          =   375
      Left            =   3720
      MaxLength       =   8
      TabIndex        =   9
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox TXThrr 
      Height          =   375
      Left            =   3720
      MaxLength       =   8
      TabIndex        =   7
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox TXTfer 
      Height          =   375
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   5
      Top             =   3720
      Width           =   2055
   End
   Begin VB.ComboBox CMBrep 
      Height          =   315
      ItemData        =   "Respuestoporcam.frx":6BD6
      Left            =   3720
      List            =   "Respuestoporcam.frx":6BD8
      TabIndex        =   3
      Top             =   3240
      Width           =   2175
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
      ItemData        =   "Respuestoporcam.frx":6BDA
      Left            =   3720
      List            =   "Respuestoporcam.frx":6BDC
      TabIndex        =   1
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Por Camión"
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
      Left            =   7680
      TabIndex        =   14
      Top             =   1200
      Width           =   2505
   End
   Begin VB.Label LBLven 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cantidad"
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
      Index           =   2
      Left            =   1320
      TabIndex        =   10
      Top             =   5040
      Width           =   945
   End
   Begin VB.Label LBLven 
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
      TabIndex        =   8
      Top             =   4440
      Width           =   525
   End
   Begin VB.Label LBLven 
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
      Top             =   3840
      Width           =   660
   End
   Begin VB.Label Label2 
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
      Top             =   3240
      Width           =   2295
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
      TabIndex        =   2
      Top             =   2640
      Width           =   1845
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
      Left            =   8040
      TabIndex        =   0
      Top             =   600
      Width           =   2070
   End
End
Attribute VB_Name = "Repuestoporcam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CMBrep_Click()
Dim t As New ADODB.Recordset
        t.Open "select CodRep from Repuesto where NomRep = '" & CMBrep.Text & "'", Conexion
        TXTcrc.Text = t!CodRep
        TXTfer.SetFocus
End Sub

Private Sub CMDeli_Click()
If CMBpla.Text = "" Then
  MsgBox " Debe Selesccionar la Placa del Camión", vbInformation, "Omision de Datos"
  CMBpla.SetFocus
  ElseIf CMBrep.Text = "" Then
    MsgBox " Debe Seleccionar El Tipo de Repuesto a Utilizar en el Camion", vbInformation, "Omision de Datos"
    CMBrep.SetFocus
  ElseIf TXTfer.Text = "" Then
    MsgBox " Debe Introducir la Fecha que va a colocar el Repuesto al Camion", vbInformation, "Omision de Datos"
    TXTfer.SetFocus
  ElseIf TXThrr.Text = "" Then
    MsgBox " Debe Introducir la Hora en que se Coloco el Repuesto al Camion", vbInformation, "Omision de Datos"
    TXThrr.SetFocus
  ElseIf TXTcar.Text = "" Then
    MsgBox " Debe Incluir la Cantidad del Repuestos Colocados al Camion", vbInformation, "Omision de Datos"
    TXTcar.SetFocus
  Else
   If MsgBox(" Desea Eliminar el Repuesto por Camion?", vbQuestion + vbYesNo, "Eliminacion de Datos") = vbYes Then
   Conexion.Execute " update RepPorCamion set EstatusRxc = 'E' where Placa = '" & CMBpla.Text & "'"
   MsgBox " Repuesto por Camion Eliminado Satisfactoriamente", vbInformation, "Eliminacion de Datos"
   CMDlim_Click
   CMBpla.SetFocus
   End If
 End If
End Sub

Private Sub CMDgua_Click()
Dim t As New ADODB.Recordset
Dim inv As New ADODB.Recordset
If CMBpla.Text = "" Then
  MsgBox " Debe Selesccionar la Placa del Camión", vbInformation, "Omision de Datos"
  CMBpla.SetFocus
  ElseIf CMBrep.Text = "" Then
    MsgBox " Debe Seleccionar El Tipo de Repuesto a Utilizar en el Camion", vbInformation, "Omision de Datos"
    CMBrep.SetFocus
  ElseIf TXTfer.Text = "" Then
    MsgBox " Debe Introducir la Fecha que va a colocar el Repuesto al Camion", vbInformation, "Omision de Datos"
    TXTfer.SetFocus
  ElseIf TXThrr.Text = "" Then
    MsgBox " Debe Introducir la Hora en que se Coloco el Repuesto al Camion", vbInformation, "Omision de Datos"
    TXThrr.SetFocus
  ElseIf TXTcar.Text = "" Then
    MsgBox " Debe Incluir la Cantidad del Repuestos Colocados al Camion", vbInformation, "Omision de Datos"
    TXTcar.SetFocus
Else

    t.Open "Select * from RepPorCamion where PlacaR = '" & CMBpla.Text & "'And CodRep = val('" & TXTcrc.Text & "') And Fecha = #" & TXTfer.Text & "# And Hora = val('" & TXThrr.Text & "') And Estatusrxc = 'A'", Conexion
        If t.EOF Then
        inv.Open " Select Existencia, Stmin from Repuesto where CodRep = '" & TXTcrc.Text & "'", Conexion
            If (inv!Existencia - Val(TXTcar.Text) < 0) Then
                    MsgBox "La Existencia de Este Producto, no puede suplir la Demanda Exigida. Revise la Existencia de este Repuesto y Verifique la Informacion", vbCritical, "Problema con el Inventario"
                    Repuestos.Show
                    Repuestos.TXTrep = TXTcrc.Text
                    Repuestos.Buscar
                    Repuestos.TXTexi.SetFocus
            ElseIf ((inv!Existencia - Val(TXTcar.Text)) < inv!StMin) Then
                If MsgBox(" La Inclusion de este Registro, Hara que la Existencia de " & CMBrep.Text & " sea Menor que el Stock Minimo. Desea Continuar? Si Presiona SI debe Ordenar la Compra Inmediata de este Repuesto para suplir la Deficiencia", vbQuestion + vbYesNo, "Problema con el Inventario") = vbYes Then
                    Conexion.Execute "insert into RepPorCamion (PlacaR,CodRep,Fecha,Hora,Cantidad,EstatusRxc) values ('" & CMBpla.Text & "', '" & TXTcrc.Text & "', '" & TXTfer.Text & "', '" & TXThrr.Text & "', " & TXTcar.Text & ", 'A')"
                    MsgBox "Repuesto por Camion Incluido Satisfactoriamente ", vbInformation, "Inclusion de Datos"
                    Conexion.Execute "update Repuesto set Existencia = Existencia - Val('" & TXTcar.Text & "') where CodRep = '" & TXTcrc.Text & "'"
                    CMDlim_Click
                    CMBpla.SetFocus
                   End If
                    Else
                    Conexion.Execute "insert into RepPorCamion (PlacaR,CodRep,Fecha,Hora,Cantidad,EstatusRxc) values ('" & CMBpla.Text & "', '" & TXTcrc.Text & "', '" & TXTfer.Text & "', '" & TXThrr.Text & "', " & TXTcar.Text & ", 'A')"
                    MsgBox "Repuesto por Camion Incluido Satisfactoriamente ", vbInformation, "Inclusion de Datos"
                    Conexion.Execute "update Repuesto set Existencia = Existencia - Val('" & TXTcar.Text & "') where CodRep = '" & TXTcrc.Text & "'"
                   CMDlim_Click
        End If
                Else
                  MsgBox " No puede Guardar ya que hay Valores Duplicados ", vbCritical, "Verifique los Datos y Vuelva a Intentarlo"
                  CMDlim_Click
                End If
            
            
            'Conexion.Execute "insert into RepPorCamion (PlacaR,CodRep,Fecha,Hora,Cantidad,EstatusRxc) values ('" & CMBpla.Text & "', '" & TXTcrc.Text & "', '" & TXTfer.Text & "', '" & TXThrr.Text & "', " & TXTcar.Text & ", 'A')"
            'MsgBox "Repuesto por Camion Incluido Satisfactoriamente ", vbInformation, "Inclusion de Datos"
            'Conexion.Execute "update Repuesto set Existencia = Existencia - Val('" & TXTcar.Text & "') where CodRep = '" & TXTcrc.Text & "'"
            'CMDlim_Click
            
            End If
    
    'MsgBox " No puede Guardar ya que hay Valores Duplicados ", vbCritical, "Verifique los Datos y Vuelva a Intentarlo"
        'CMDlim_Click

End Sub

Private Sub CMDlim_Click()
CMBpla.Text = ""
CMBrep.Text = ""
TXTcrc.Text = ""
TXTfer.Text = ""
TXThrr.Text = ""
TXTcar.Text = ""
End Sub

Private Sub CMDmod_Click()
If CMBpla.Text = "" Then
  MsgBox " Debe Selesccionar la Placa del Camión", vbInformation, "Omision de Datos"
  CMBpla.SetFocus
  ElseIf CMBrep.Text = "" Then
    MsgBox " Debe Seleccionar El Tipo de Repuesto a Utilizar en el Camionj", vbInformation, "Omision de Datos"
    CMBrep.SetFocus
  ElseIf TXTfer.Text = "" Then
    MsgBox " Debe Introducir la Fecha que va a colocar el Repuesto al Camion", vbInformation, "Omision de Datos"
    TXTfer.SetFocus
  ElseIf TXThrr.Text = "" Then
    MsgBox " Debe Introducir la Hora en que se Coloco el Repuesto al Camion", vbInformation, "Omision de Datos"
    TXThrr.SetFocus
  ElseIf TXTcar.Text = "" Then
    MsgBox " Debe Incluir la Cantidad del Repuestos Colocados al Camion", vbInformation, "Omision de Datos"
    TXTcar.SetFocus
  Else
   Conexion.Execute " update RepPorCamion set Fecha = '" & Trim(TXTfer.Text) & "', Hora = '" & Trim(TXThrr.Text) & "', Cantidad = '" & Trim(TXTcar.Text) & "' where "
    MsgBox " Repuesto por Camion Modificado Satisfactoriamente ", vbInformation, "Modificacion de Datos"
    CMDlim_Click
    CMBpla.SetFocus
End If
End Sub
Private Sub CMDsal_Click()
If MsgBox(" Desea realmente salir de 'Respuestos por Camion'?", vbQuestion + vbYesNo, "Salida de Pantalla") = vbYes Then
Unload Me
End If
End Sub

Private Sub Command1_Click()
ReporteRepCaMion.Show
End Sub

Private Sub Form_Load()
LlenarCombo CMBpla, "select * from Camiones", "Placa"
LlenarCombo CMBrep, "select * from Repuesto", "NomRep"
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Sub activar()
Conexion.Execute " update RepPorCamion set EstatusRxc = 'A' where Placa = '" & CMBpla.Text & "'"
End Sub



Private Sub Label4_Click()

End Sub

Private Sub TXTcar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsNumeric(TXTcar.Text) = False Then
        MsgBox " Este Campo es de Tipo Numerico, Cambie la Condición", vbExclamation, "Dato Incorrecto"
        TXTcar.Text = ""
        TXTcar.SetFocus
    ElseIf Val(TXTcar.Text) < 0 Then
        MsgBox "La Cantidad de Repuestos Usados, no puede ser Negativo, Cambie la Condición", vbExclamation, "Dato Incorrecto"
        TXTcar.Text = ""
        TXTcar.SetFocus
    Else
    CMDgua_Click
    End If
End If
End Sub

Private Sub TXTfer_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsDate(TXTfer.Text) = False Then
        MsgBox "Este Campo es de tipo Fecha, Cambie la Condición", vbExclamation, "Dato Incorrecto"
        TXTfer.Text = ""
        TXTfer.SetFocus
    Else
    TXThrr.SetFocus
    End If
End If
End Sub

Private Sub TXThrr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsDate(TXThrr.Text) = False Then
        MsgBox "Este Campo admite solo Horas, Cambie la Condición", vbExclamation, "Dato Incorrecto"
        TXThrr.Text = ""
        TXThrr.SetFocus
    Else
    TXTcar.SetFocus
    End If
End If
End Sub
