VERSION 5.00
Begin VB.Form RegistrarSocio 
   ClientHeight    =   6660
   ClientLeft      =   3240
   ClientTop       =   4155
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   Picture         =   "RegistrarSocios2.frx":0000
   ScaleHeight     =   6660
   ScaleWidth      =   9360
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8640
      Picture         =   "RegistrarSocios2.frx":CAE9E
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6720
      Picture         =   "RegistrarSocios2.frx":CB428
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6120
      Picture         =   "RegistrarSocios2.frx":CB9B2
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4560
      Picture         =   "RegistrarSocios2.frx":CBF3C
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1320
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   9135
      Begin VB.TextBox Text16 
         Height          =   375
         Left            =   7680
         TabIndex        =   42
         Text            =   " "
         Top             =   4440
         Width           =   1335
      End
      Begin VB.TextBox Text15 
         Height          =   375
         Left            =   7680
         TabIndex        =   36
         Text            =   " "
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   2040
         TabIndex        =   34
         Text            =   " "
         Top             =   4920
         Width           =   3735
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   7680
         TabIndex        =   32
         Text            =   " "
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   7680
         TabIndex        =   30
         Text            =   " "
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   7680
         TabIndex        =   28
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   2040
         TabIndex        =   26
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   2040
         TabIndex        =   24
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   6840
         TabIndex        =   23
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   2040
         TabIndex        =   22
         Top             =   4440
         Width           =   2055
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   2040
         TabIndex        =   19
         Top             =   3000
         Width           =   3735
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Divorciado"
         Height          =   255
         Left            =   4440
         TabIndex        =   17
         Top             =   2640
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Viudo"
         Height          =   255
         Left            =   3720
         TabIndex        =   16
         Top             =   2640
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Soltero"
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   2640
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Casado"
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   2040
         TabIndex        =   11
         Top             =   3960
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2040
         TabIndex        =   10
         Text            =   "Codigos"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3120
         TabIndex        =   9
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   2040
         TabIndex        =   8
         Top             =   1200
         Width           =   4935
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   2040
         TabIndex        =   7
         Top             =   720
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Bancaria"
         Height          =   375
         Left            =   6120
         TabIndex        =   41
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Banco Donde Realiza su Cobro"
         Height          =   495
         Left            =   6120
         TabIndex        =   35
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Lugar de Trabajo Actual"
         Height          =   495
         Left            =   720
         TabIndex        =   33
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Aporte Institucional"
         Height          =   375
         Left            =   6120
         TabIndex        =   31
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Aporte Voluntario"
         Height          =   375
         Left            =   6120
         TabIndex        =   29
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Aporte del 10% del Sueldo"
         Height          =   495
         Left            =   6120
         TabIndex        =   27
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Lugar de Nacimiento"
         Height          =   495
         Left            =   720
         TabIndex        =   25
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Sueldo Base"
         Height          =   375
         Left            =   720
         TabIndex        =   21
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Ingreso         a la Policia"
         Height          =   495
         Left            =   720
         TabIndex        =   20
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Profesion"
         Height          =   375
         Left            =   720
         TabIndex        =   18
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Civil"
         Height          =   375
         Left            =   720
         TabIndex        =   13
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Nacimiento"
         Height          =   375
         Left            =   5160
         TabIndex        =   12
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "      Rango"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "     Telefono"
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "     Direccion"
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "     Nombre"
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "      Cedula"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "RegistrarSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
Unload Me
End Sub

