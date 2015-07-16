VERSION 5.00
Begin VB.Form fmrprincipal 
   Caption         =   "ucla"
   ClientHeight    =   4785
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Height          =   5130
      Left            =   -120
      Picture         =   "fmrprincipal.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11205
   End
   Begin VB.Menu mnuestudiante 
      Caption         =   "Estudiantes"
      Enabled         =   0   'False
      Begin VB.Menu mnunuevo 
         Caption         =   "Nuevo"
      End
   End
   Begin VB.Menu mnumaterias 
      Caption         =   "Materias"
      Enabled         =   0   'False
      Begin VB.Menu mnunuevom 
         Caption         =   "Nuevo"
      End
   End
   Begin VB.Menu mnuoperaciones 
      Caption         =   "Operaciones"
      Begin VB.Menu mnumatealu 
         Caption         =   "materias por estudiante"
      End
   End
   Begin VB.Menu mnusalir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "fmrprincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Set Base_Datos = New ADODB.Connection
Base_Datos.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\bd.mdb;Persist Security Info=False"
SOpt = dbsqlpassthough
End Sub

Private Sub mnumatealu_Click()
FormMateEstu.Show 1
'fmrmateriasxestudiante.Show 1
End Sub

Private Sub mnusalir_Click()
If MsgBox(" Esta seguro que deseas salir ", vbYesNo + vbQuestion) = vbYes Then
       Unload Me
       End If
End Sub
