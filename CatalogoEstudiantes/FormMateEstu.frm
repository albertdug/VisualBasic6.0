VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormMateEstu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estudiantes x Materia"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "MeteriasxEstudiantes"
      Height          =   2535
      Left            =   600
      TabIndex        =   7
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtcodest 
         BackColor       =   &H8000000F&
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtcodmat 
         BackColor       =   &H8000000F&
         Height          =   405
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtnota 
         Height          =   375
         Left            =   2760
         MaxLength       =   3
         TabIndex        =   10
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton btnconsultare 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         Picture         =   "FormMateEstu.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "consultar clientes"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton btnconsultarm 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         Picture         =   "FormMateEstu.frx":0424
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "consultar clientes"
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo del Estudiante:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   2250
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codigo de la materia:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   2100
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nota:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   6735
      Begin VB.CommandButton btnsalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5520
         Picture         =   "FormMateEstu.frx":0848
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "salir"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton btnmodificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         Picture         =   "FormMateEstu.frx":0DC9
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Modifica el Registro Actual"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton btnincluir 
         Caption         =   "&Incluir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Picture         =   "FormMateEstu.frx":1262
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Agrega un nuevo registros"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton btncancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         Picture         =   "FormMateEstu.frx":16DF
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Elimina el registro Actual"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton btnguardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3120
         Picture         =   "FormMateEstu.frx":1929
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton btneliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4200
         Picture         =   "FormMateEstu.frx":21F3
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2040
      Top             =   4560
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\trabajo visual\bd.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\trabajo visual\bd.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "FormMateEstu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String
Dim TablaMatEstu As New ADODB.Recordset

Private Sub btnconsultare_Click()
FormCataEstudiantes.Show 1
If Escogio Then
   txtcodest.Text = CodigoCata
   Sql = "SELECT * FROM materiasxestudiante WHERE ci=" & Apost(txtcodest.Text)
   Sql = Sql & " AND codmateria=" & Apost(txtcodmat.Text)
   Call OpenTabla(TablaMatEstu, Sql)
   If Not TablaMatEstu.EOF Then
      txtnota.Text = TablaMatEstu!Nota
      txtnota.Locked = True
      btneliminar.Enabled = True
   Else
      txtnota.Text = 0
      txtnota.Locked = False
      btneliminar.Enabled = False
   End If
End If
End Sub

Private Sub btnconsultarm_Click()
FormCataMaterias.Show 1
If Escogio Then
   txtcodmat.Text = CodigoCata
   Sql = "SELECT * FROM materiasxestudiante WHERE ci=" & Apost(txtcodest.Text)
   Sql = Sql & " AND codmateria=" & Apost(txtcodmat.Text)
   Call OpenTabla(TablaMatEstu, Sql)
   If Not TablaMatEstu.EOF Then
      txtnota.Text = TablaMatEstu!Nota
      txtnota.Locked = True
      btneliminar.Enabled = True
   Else
      txtnota.Text = 0
      txtnota.Locked = False
      btneliminar.Enabled = False
   End If
End If
End Sub

Private Sub btneliminar_Click()
If Not Existe("estudiantes", "Ci", txtcodest.Text) Then
   MsgBox ("Estudiante No Existe")
   txtcodest.SetFocus
   Exit Sub
End If

If Not Existe("materias", "codmat", txtcodmat.Text) Then
   MsgBox ("Materia No Existe")
   txtcodmat.SetFocus
   Exit Sub
End If

Sql = "SELECT * FROM materiasxestudiante WHERE ci=" & Apost(txtcodest.Text)
Sql = Sql & " AND codmateria=" & Apost(txtcodmat.Text)
Call OpenTabla(TablaMatEstu, Sql)
If (Not TablaMatEstu.EOF) And (TablaMatEstu!estatus = "E") Then
   MsgBox ("Estudiante/Materia Esta Inactiva")
      txtcodest.Text = ""
      txtcodmat.Text = ""
      txtnota.Text = ""
      txtcodest.SetFocus
   Exit Sub
End If

If (Not TablaMatEstu.EOF) And (TablaMatEstu!estatus = "A") Then
   If CInt(txtnota.Text) = 0 Then
    Sql = "UPDATE materiasxestudiante SET Estatus='E' WHERE Ci=" & Apost(txtcodest.Text)
      Sql = Sql & " AND codmateria=" & Apost(txtcodmat.Text)
      Base_Datos.Execute Sql
      MsgBox ("Materia/Estudiante Eliminado Exitosamente")
      txtcodest.Text = ""
      txtcodmat.Text = ""
      txtnota.Text = ""
      txtcodest.SetFocus
      btneliminar.Enabled = False
   Else
       MsgBox ("La Nota es diferente de Cero.No Se Puede Eliminar")
       txtcodest.Text = ""
       txtcodmat.Text = ""
       txtnota.Text = ""
       txtcodest.SetFocus
   End If
End If
End Sub

Private Sub btnincluir_Click()
If Not Existe("estudiantes", "Ci", txtcodest.Text) Then
   MsgBox ("Estudiante No Existe")
   txtcodest.SetFocus
   Exit Sub
End If

If Not Existe("materias", "codmat", txtcodmat.Text) Then
   MsgBox ("Materia No Existe")
   txtcodmat.SetFocus
   Exit Sub
End If

 Sql = "SELECT * FROM materiasxestudiante WHERE ci=" & Apost(txtcodest.Text)
 Sql = Sql & " AND codmateria=" & Apost(txtcodmat.Text)
 Call OpenTabla(TablaMatEstu, Sql)
 If (Not TablaMatEstu.EOF) Then
 'And (TablaMatEstu!estatus = "E") Then
 'If MsgBox("Estudiante/Materia Esta Inactiva ,¿Desea activarla?", vbQuestion + vbYesNo) = vbYes Then
   ' Sql = "UPDATE materiasxestudiante SET Estatus='A' WHERE Ci=" & Apost(txtcodest.Text)
    '  Sql = Sql & " AND codmateria=" & Apost(txtcodmat.Text)
     ' Base_Datos.Execute Sql
     ' MsgBox ("Reguistro activado exitosamente")
     ' txtcodest.Text = ""
     ' txtcodmat.Text = ""
     ' txtnota.Text = ""
     ' txtcodest.SetFocus
    'End If
   'Exit Sub
'End If

'Sql = "SELECT * FROM materiasxestudiante WHERE ci=" & Apost(txtcodest.Text)
 'Sql = Sql & " AND codmateria=" & Apost(txtcodmat.Text)
 'Call OpenTabla(TablaMatEstu, Sql)
'If (Not TablaMatEstu.EOF) And (TablaMatEstu!estatus = "A") Then
'Sql = "SELECT * FROM materiasxestudiante WHERE ci=" & Apost(txtcodest.Text)
'Sql = Sql & " AND codmateria=" & Apost(txtcodmat.Text)
'Call OpenTabla(TablaMatEstu, Sql)
'If Not TablaMatEstu.EOF Then
   MsgBox ("Estudiante/Materia ya Incluida")
      txtcodest.Text = ""
      txtcodmat.Text = ""
      txtnota.Text = ""
      txtcodest.SetFocus
   Exit Sub
End If

Sql = "INSERT INTO materiasxestudiante (ci,codmateria,nota,estatus) VALUES("
Sql = Sql & Apost(txtcodest.Text) & ","
Sql = Sql & Apost(txtcodmat.Text) & ","
Sql = Sql & Apost(txtnota.Text) & ",'A')"
Base_Datos.Execute Sql
      MsgBox ("Materia/Estudiante incluido exitosamente")
      txtcodest.Text = ""
      txtcodmat.Text = ""
      txtnota.Text = ""
      txtcodest.SetFocus
      
End Sub

Private Sub btnsalir_Click()
If MsgBox(" Esta seguro que deseas salir ", vbYesNo + vbQuestion) = vbYes Then
       Unload Me
       End If
End Sub






Private Sub txtnota_KeyPress(KeyAscii As Integer)
Call ValidaEntero(KeyAscii)
End Sub
