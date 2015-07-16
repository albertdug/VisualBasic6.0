VERSION 5.00
Begin VB.Form Registrarpago 
   Caption         =   "Registrar Pago"
   ClientHeight    =   6615
   ClientLeft      =   3240
   ClientTop       =   3735
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Registrarpago.frx":0000
   ScaleHeight     =   6615
   ScaleWidth      =   9390
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   960
      TabIndex        =   5
      Top             =   1200
      Width           =   7455
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         Picture         =   "Registrarpago.frx":CAE9E
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   1080
         Width           =   495
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   2655
         Left            =   6600
         TabIndex        =   61
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox Text90 
         Height          =   285
         Left            =   5520
         TabIndex        =   60
         Text            =   " "
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox Text89 
         Height          =   285
         Left            =   5520
         TabIndex        =   59
         Text            =   " "
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox Text88 
         Height          =   285
         Left            =   5520
         TabIndex        =   58
         Text            =   " "
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox Text87 
         Height          =   285
         Left            =   5520
         TabIndex        =   57
         Text            =   " "
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Text86 
         Height          =   285
         Left            =   5520
         TabIndex        =   56
         Text            =   " "
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Text85 
         Height          =   285
         Left            =   5520
         TabIndex        =   55
         Text            =   " "
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text84 
         Height          =   285
         Left            =   5520
         TabIndex        =   54
         Text            =   " "
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text83 
         Height          =   285
         Left            =   5520
         TabIndex        =   53
         Text            =   " "
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text82 
         Height          =   285
         Left            =   5520
         TabIndex        =   52
         Text            =   " "
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text81 
         Height          =   285
         Left            =   5520
         TabIndex        =   51
         Text            =   " "
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text80 
         Height          =   285
         Left            =   5520
         TabIndex        =   50
         Text            =   " "
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text79 
         Height          =   285
         Left            =   4560
         TabIndex        =   49
         Text            =   " "
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox Text78 
         Height          =   285
         Left            =   4560
         TabIndex        =   48
         Text            =   " "
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox Text77 
         Height          =   285
         Left            =   4560
         TabIndex        =   47
         Text            =   " "
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox Text76 
         Height          =   285
         Left            =   4560
         TabIndex        =   46
         Text            =   " "
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox Text75 
         Height          =   285
         Left            =   4560
         TabIndex        =   45
         Text            =   " "
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox Text74 
         Height          =   285
         Left            =   4560
         TabIndex        =   44
         Text            =   " "
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox Text73 
         Height          =   285
         Left            =   4560
         TabIndex        =   43
         Text            =   " "
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Text72 
         Height          =   285
         Left            =   4560
         TabIndex        =   42
         Text            =   " "
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Text71 
         Height          =   285
         Left            =   4560
         TabIndex        =   41
         Text            =   " "
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text70 
         Height          =   285
         Left            =   4560
         TabIndex        =   40
         Text            =   " "
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text69 
         Height          =   285
         Left            =   4560
         TabIndex        =   39
         Text            =   " "
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text68 
         Height          =   285
         Left            =   3240
         TabIndex        =   38
         Text            =   " "
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox Text67 
         Height          =   285
         Left            =   3240
         TabIndex        =   37
         Text            =   " "
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Text66 
         Height          =   285
         Left            =   3240
         TabIndex        =   36
         Text            =   " "
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Text65 
         Height          =   285
         Left            =   3240
         TabIndex        =   35
         Text            =   " "
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Text64 
         Height          =   285
         Left            =   3240
         TabIndex        =   34
         Text            =   " "
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox Text63 
         Height          =   285
         Left            =   3240
         TabIndex        =   33
         Text            =   " "
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text62 
         Height          =   285
         Left            =   3240
         TabIndex        =   32
         Text            =   " "
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Text61 
         Height          =   285
         Left            =   3240
         TabIndex        =   31
         Text            =   " "
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Text55 
         Height          =   285
         Left            =   3240
         TabIndex        =   30
         Text            =   " "
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text54 
         Height          =   285
         Left            =   3240
         TabIndex        =   29
         Text            =   " "
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text53 
         Height          =   285
         Left            =   3240
         TabIndex        =   28
         Text            =   " "
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text52 
         Height          =   285
         Left            =   2040
         TabIndex        =   27
         Text            =   " "
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox Text51 
         Height          =   285
         Left            =   2040
         TabIndex        =   26
         Text            =   " "
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Text50 
         Height          =   285
         Left            =   2040
         TabIndex        =   25
         Text            =   " "
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text49 
         Height          =   285
         Left            =   2040
         TabIndex        =   24
         Text            =   " "
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox Text48 
         Height          =   285
         Left            =   2040
         TabIndex        =   23
         Text            =   " "
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox Text47 
         Height          =   285
         Left            =   2040
         TabIndex        =   22
         Text            =   " "
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text46 
         Height          =   285
         Left            =   2040
         TabIndex        =   21
         Text            =   " "
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Text45 
         Height          =   285
         Left            =   2040
         TabIndex        =   20
         Text            =   " "
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Text44 
         Height          =   285
         Left            =   2040
         TabIndex        =   19
         Text            =   " "
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text43 
         Height          =   285
         Left            =   2040
         TabIndex        =   18
         Text            =   " "
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text28 
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Text            =   " "
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   720
         TabIndex        =   16
         Text            =   " "
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   720
         TabIndex        =   15
         Text            =   " "
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   720
         TabIndex        =   14
         Text            =   " "
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   720
         TabIndex        =   13
         Text            =   " "
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Text            =   " "
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Text            =   " "
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text56 
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Text            =   " "
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Text57 
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Text            =   " "
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Text58 
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Text            =   " "
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text59 
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Text            =   " "
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text60 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Text            =   " "
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label N 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero de Solicitud"
         Height          =   495
         Left            =   2040
         TabIndex        =   68
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto"
         Height          =   495
         Left            =   3240
         TabIndex        =   67
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
         Height          =   495
         Left            =   5520
         TabIndex        =   66
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero de Cuota"
         Height          =   495
         Left            =   4560
         TabIndex        =   65
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo "
         Height          =   495
         Left            =   720
         TabIndex        =   64
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PAGOS:"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      Picture         =   "Registrarpago.frx":CB428
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      Picture         =   "Registrarpago.frx":CB9B2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8760
      Picture         =   "Registrarpago.frx":CBF3C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      Picture         =   "Registrarpago.frx":CC4C6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2760
      Picture         =   "Registrarpago.frx":CCA50
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   495
   End
End
Attribute VB_Name = "Registrarpago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command4_Click()
Unload Me
End Sub



