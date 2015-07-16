VERSION 5.00
Begin VB.Form RegistrarRetiro 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Registrar Retiro"
   ClientHeight    =   6675
   ClientLeft      =   3240
   ClientTop       =   3945
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "RegistrarRetiro.frx":0000
   ScaleHeight     =   6675
   ScaleWidth      =   9315
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   8415
      Begin VB.ComboBox Combo12 
         Height          =   315
         ItemData        =   "RegistrarRetiro.frx":CAE9E
         Left            =   1920
         List            =   "RegistrarRetiro.frx":CAEA8
         TabIndex        =   80
         Top             =   3600
         Width           =   1335
      End
      Begin VB.ComboBox Combo11 
         Height          =   315
         ItemData        =   "RegistrarRetiro.frx":CAEBC
         Left            =   1920
         List            =   "RegistrarRetiro.frx":CAEC6
         TabIndex        =   79
         Top             =   3360
         Width           =   1335
      End
      Begin VB.ComboBox Combo10 
         Height          =   315
         ItemData        =   "RegistrarRetiro.frx":CAEDA
         Left            =   1920
         List            =   "RegistrarRetiro.frx":CAEE4
         TabIndex        =   78
         Top             =   3120
         Width           =   1335
      End
      Begin VB.ComboBox Combo9 
         Height          =   315
         ItemData        =   "RegistrarRetiro.frx":CAEF8
         Left            =   1920
         List            =   "RegistrarRetiro.frx":CAF02
         TabIndex        =   77
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         ItemData        =   "RegistrarRetiro.frx":CAF16
         Left            =   1920
         List            =   "RegistrarRetiro.frx":CAF20
         TabIndex        =   76
         Top             =   2640
         Width           =   1335
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         ItemData        =   "RegistrarRetiro.frx":CAF34
         Left            =   1920
         List            =   "RegistrarRetiro.frx":CAF3E
         TabIndex        =   75
         Top             =   2400
         Width           =   1335
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "RegistrarRetiro.frx":CAF52
         Left            =   1920
         List            =   "RegistrarRetiro.frx":CAF5C
         TabIndex        =   74
         Top             =   2160
         Width           =   1335
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "RegistrarRetiro.frx":CAF70
         Left            =   1920
         List            =   "RegistrarRetiro.frx":CAF7A
         TabIndex        =   73
         Top             =   1920
         Width           =   1335
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "RegistrarRetiro.frx":CAF8E
         Left            =   1920
         List            =   "RegistrarRetiro.frx":CAF98
         TabIndex        =   72
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "RegistrarRetiro.frx":CAFAC
         Left            =   1920
         List            =   "RegistrarRetiro.frx":CAFB6
         TabIndex        =   71
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "RegistrarRetiro.frx":CAFCA
         Left            =   1920
         List            =   "RegistrarRetiro.frx":CAFD4
         TabIndex        =   70
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text60 
         Height          =   285
         Left            =   600
         TabIndex        =   62
         Text            =   " "
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text59 
         Height          =   285
         Left            =   600
         TabIndex        =   61
         Text            =   " "
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text58 
         Height          =   285
         Left            =   600
         TabIndex        =   60
         Text            =   " "
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text57 
         Height          =   285
         Left            =   600
         TabIndex        =   59
         Text            =   " "
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text56 
         Height          =   285
         Left            =   600
         TabIndex        =   58
         Text            =   " "
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   600
         TabIndex        =   57
         Text            =   " "
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   600
         TabIndex        =   56
         Text            =   " "
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   600
         TabIndex        =   55
         Text            =   " "
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   600
         TabIndex        =   54
         Text            =   " "
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   600
         TabIndex        =   53
         Text            =   " "
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   600
         TabIndex        =   52
         Text            =   " "
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   3240
         TabIndex        =   51
         Text            =   " "
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   3240
         TabIndex        =   50
         Text            =   " "
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   3240
         TabIndex        =   49
         Text            =   " "
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   3240
         TabIndex        =   48
         Text            =   " "
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   3240
         TabIndex        =   47
         Text            =   " "
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   3240
         TabIndex        =   46
         Text            =   " "
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   3240
         TabIndex        =   45
         Text            =   " "
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   3240
         TabIndex        =   44
         Text            =   " "
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   3240
         TabIndex        =   43
         Text            =   " "
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   3240
         TabIndex        =   42
         Text            =   " "
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   3240
         TabIndex        =   41
         Text            =   " "
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   4440
         TabIndex        =   40
         Text            =   " "
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   4440
         TabIndex        =   39
         Text            =   " "
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text25 
         Height          =   285
         Left            =   4440
         TabIndex        =   38
         Text            =   " "
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text26 
         Height          =   285
         Left            =   4440
         TabIndex        =   37
         Text            =   " "
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   4440
         TabIndex        =   36
         Text            =   " "
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text28 
         Height          =   285
         Left            =   4440
         TabIndex        =   35
         Text            =   " "
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text29 
         Height          =   285
         Left            =   4440
         TabIndex        =   34
         Text            =   " "
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text30 
         Height          =   285
         Left            =   4440
         TabIndex        =   33
         Text            =   " "
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text31 
         Height          =   285
         Left            =   4440
         TabIndex        =   32
         Text            =   " "
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Text32 
         Height          =   285
         Left            =   4440
         TabIndex        =   31
         Text            =   " "
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox Text33 
         Height          =   285
         Left            =   4440
         TabIndex        =   30
         Text            =   " "
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox Text34 
         Height          =   285
         Left            =   5760
         TabIndex        =   29
         Text            =   " "
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text35 
         Height          =   285
         Left            =   5760
         TabIndex        =   28
         Text            =   " "
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox Text36 
         Height          =   285
         Left            =   5760
         TabIndex        =   27
         Text            =   " "
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text37 
         Height          =   285
         Left            =   5760
         TabIndex        =   26
         Text            =   " "
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox Text38 
         Height          =   285
         Left            =   5760
         TabIndex        =   25
         Text            =   " "
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Text39 
         Height          =   285
         Left            =   5760
         TabIndex        =   24
         Text            =   " "
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox Text40 
         Height          =   285
         Left            =   5760
         TabIndex        =   23
         Text            =   " "
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox Text41 
         Height          =   285
         Left            =   5760
         TabIndex        =   22
         Text            =   " "
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox Text42 
         Height          =   285
         Left            =   5760
         TabIndex        =   21
         Text            =   " "
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox Text43 
         Height          =   285
         Left            =   5760
         TabIndex        =   20
         Text            =   " "
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox Text44 
         Height          =   285
         Left            =   5760
         TabIndex        =   19
         Text            =   " "
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox Text45 
         Height          =   285
         Left            =   6720
         TabIndex        =   18
         Text            =   " "
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text46 
         Height          =   285
         Left            =   6720
         TabIndex        =   17
         Text            =   " "
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text47 
         Height          =   285
         Left            =   6720
         TabIndex        =   16
         Text            =   " "
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text48 
         Height          =   285
         Left            =   6720
         TabIndex        =   15
         Text            =   " "
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text49 
         Height          =   285
         Left            =   6720
         TabIndex        =   14
         Text            =   " "
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox Text50 
         Height          =   285
         Left            =   6720
         TabIndex        =   13
         Text            =   " "
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text51 
         Height          =   285
         Left            =   6720
         TabIndex        =   12
         Text            =   " "
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox Text52 
         Height          =   285
         Left            =   6720
         TabIndex        =   11
         Text            =   " "
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox Text53 
         Height          =   285
         Left            =   6720
         TabIndex        =   10
         Text            =   " "
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox Text54 
         Height          =   285
         Left            =   6720
         TabIndex        =   9
         Text            =   " "
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox Text55 
         Height          =   285
         Left            =   6720
         TabIndex        =   8
         Text            =   " "
         Top             =   3600
         Width           =   1095
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2655
         Left            =   7800
         TabIndex        =   7
         Top             =   1200
         Width           =   255
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         Picture         =   "RegistrarRetiro.frx":CAFE8
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RETIROS:"
         Height          =   255
         Left            =   480
         TabIndex        =   69
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo "
         Height          =   375
         Left            =   600
         TabIndex        =   68
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cedulla"
         Height          =   375
         Left            =   5760
         TabIndex        =   67
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Disponibilidad"
         Height          =   375
         Left            =   6720
         TabIndex        =   66
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
         Height          =   375
         Left            =   4440
         TabIndex        =   65
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo  "
         Height          =   375
         Left            =   1920
         TabIndex        =   64
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto"
         Height          =   375
         Left            =   3240
         TabIndex        =   63
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      Picture         =   "RegistrarRetiro.frx":CB572
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      Picture         =   "RegistrarRetiro.frx":CBAFC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8760
      Picture         =   "RegistrarRetiro.frx":CC086
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2760
      Picture         =   "RegistrarRetiro.frx":CC610
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      Picture         =   "RegistrarRetiro.frx":CCB9A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   495
   End
End
Attribute VB_Name = "RegistrarRetiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
Unload Me
End Sub

