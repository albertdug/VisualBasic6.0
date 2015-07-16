VERSION 5.00
Begin VB.Form RegistrarAporte 
   ClientHeight    =   6660
   ClientLeft      =   3240
   ClientTop       =   3735
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "aportes.frx":0000
   ScaleHeight     =   6660
   ScaleWidth      =   9375
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   8415
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         Picture         =   "aportes.frx":CAE9E
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   1200
         Width           =   495
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2655
         Left            =   7800
         TabIndex        =   80
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox Text55 
         Height          =   285
         Left            =   6720
         TabIndex        =   79
         Text            =   " "
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox Text54 
         Height          =   285
         Left            =   6720
         TabIndex        =   78
         Text            =   " "
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox Text53 
         Height          =   285
         Left            =   6720
         TabIndex        =   77
         Text            =   " "
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox Text52 
         Height          =   285
         Left            =   6720
         TabIndex        =   76
         Text            =   " "
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox Text51 
         Height          =   285
         Left            =   6720
         TabIndex        =   75
         Text            =   " "
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox Text50 
         Height          =   285
         Left            =   6720
         TabIndex        =   74
         Text            =   " "
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text49 
         Height          =   285
         Left            =   6720
         TabIndex        =   73
         Text            =   " "
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox Text48 
         Height          =   285
         Left            =   6720
         TabIndex        =   72
         Text            =   " "
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text47 
         Height          =   285
         Left            =   6720
         TabIndex        =   71
         Text            =   " "
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text46 
         Height          =   285
         Left            =   6720
         TabIndex        =   70
         Text            =   " "
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text45 
         Height          =   285
         Left            =   6720
         TabIndex        =   69
         Text            =   " "
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text44 
         Height          =   285
         Left            =   5760
         TabIndex        =   68
         Text            =   " "
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox Text43 
         Height          =   285
         Left            =   5760
         TabIndex        =   67
         Text            =   " "
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox Text42 
         Height          =   285
         Left            =   5760
         TabIndex        =   66
         Text            =   " "
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox Text41 
         Height          =   285
         Left            =   5760
         TabIndex        =   65
         Text            =   " "
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox Text40 
         Height          =   285
         Left            =   5760
         TabIndex        =   64
         Text            =   " "
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox Text39 
         Height          =   285
         Left            =   5760
         TabIndex        =   63
         Text            =   " "
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox Text38 
         Height          =   285
         Left            =   5760
         TabIndex        =   62
         Text            =   " "
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Text37 
         Height          =   285
         Left            =   5760
         TabIndex        =   61
         Text            =   " "
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox Text36 
         Height          =   285
         Left            =   5760
         TabIndex        =   60
         Text            =   " "
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text35 
         Height          =   285
         Left            =   5760
         TabIndex        =   59
         Text            =   " "
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox Text34 
         Height          =   285
         Left            =   5760
         TabIndex        =   58
         Text            =   " "
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text33 
         Height          =   285
         Left            =   4440
         TabIndex        =   57
         Text            =   " "
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox Text32 
         Height          =   285
         Left            =   4440
         TabIndex        =   56
         Text            =   " "
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox Text31 
         Height          =   285
         Left            =   4440
         TabIndex        =   55
         Text            =   " "
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Text30 
         Height          =   285
         Left            =   4440
         TabIndex        =   54
         Text            =   " "
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text29 
         Height          =   285
         Left            =   4440
         TabIndex        =   53
         Text            =   " "
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text28 
         Height          =   285
         Left            =   4440
         TabIndex        =   52
         Text            =   " "
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   4440
         TabIndex        =   51
         Text            =   " "
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text26 
         Height          =   285
         Left            =   4440
         TabIndex        =   50
         Text            =   " "
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text25 
         Height          =   285
         Left            =   4440
         TabIndex        =   49
         Text            =   " "
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   4440
         TabIndex        =   48
         Text            =   " "
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   4440
         TabIndex        =   47
         Text            =   " "
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   3240
         TabIndex        =   46
         Text            =   " "
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   3240
         TabIndex        =   45
         Text            =   " "
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   3240
         TabIndex        =   44
         Text            =   " "
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   3240
         TabIndex        =   43
         Text            =   " "
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   3240
         TabIndex        =   42
         Text            =   " "
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   3240
         TabIndex        =   41
         Text            =   " "
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   3240
         TabIndex        =   40
         Text            =   " "
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   3240
         TabIndex        =   39
         Text            =   " "
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   3240
         TabIndex        =   38
         Text            =   " "
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   3240
         TabIndex        =   37
         Text            =   " "
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   3240
         TabIndex        =   36
         Text            =   " "
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   600
         TabIndex        =   35
         Text            =   " "
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   600
         TabIndex        =   34
         Text            =   " "
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   600
         TabIndex        =   33
         Text            =   " "
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   600
         TabIndex        =   32
         Text            =   " "
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   600
         TabIndex        =   31
         Text            =   " "
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   600
         TabIndex        =   30
         Text            =   " "
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   600
         TabIndex        =   29
         Text            =   " "
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   600
         TabIndex        =   28
         Text            =   " "
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   600
         TabIndex        =   27
         Text            =   " "
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   600
         TabIndex        =   26
         Text            =   " "
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   600
         TabIndex        =   25
         Text            =   " "
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox Combo11 
         Height          =   315
         ItemData        =   "aportes.frx":CB428
         Left            =   1920
         List            =   "aportes.frx":CB432
         TabIndex        =   24
         Top             =   3600
         Width           =   1335
      End
      Begin VB.ComboBox Combo10 
         Height          =   315
         ItemData        =   "aportes.frx":CB446
         Left            =   1920
         List            =   "aportes.frx":CB450
         TabIndex        =   23
         Top             =   3360
         Width           =   1335
      End
      Begin VB.ComboBox Combo9 
         Height          =   315
         ItemData        =   "aportes.frx":CB464
         Left            =   1920
         List            =   "aportes.frx":CB46E
         TabIndex        =   22
         Top             =   3120
         Width           =   1335
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         ItemData        =   "aportes.frx":CB482
         Left            =   1920
         List            =   "aportes.frx":CB48C
         TabIndex        =   21
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         ItemData        =   "aportes.frx":CB4A0
         Left            =   1920
         List            =   "aportes.frx":CB4AA
         TabIndex        =   20
         Top             =   2640
         Width           =   1335
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "aportes.frx":CB4BE
         Left            =   1920
         List            =   "aportes.frx":CB4C8
         TabIndex        =   19
         Top             =   2400
         Width           =   1335
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "aportes.frx":CB4DC
         Left            =   1920
         List            =   "aportes.frx":CB4E6
         TabIndex        =   18
         Top             =   2160
         Width           =   1335
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "aportes.frx":CB4FA
         Left            =   1920
         List            =   "aportes.frx":CB504
         TabIndex        =   17
         Top             =   1920
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "aportes.frx":CB518
         Left            =   1920
         List            =   "aportes.frx":CB522
         TabIndex        =   16
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "aportes.frx":CB536
         Left            =   1920
         List            =   "aportes.frx":CB540
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "aportes.frx":CB554
         Left            =   1920
         List            =   "aportes.frx":CB55E
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto  "
         Height          =   375
         Left            =   3240
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo  "
         Height          =   375
         Left            =   1920
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Height          =   375
         Left            =   4440
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
         Height          =   375
         Left            =   6720
         TabIndex        =   11
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cedula"
         Height          =   375
         Left            =   5760
         TabIndex        =   10
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo "
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "APORTES:"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   870
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      Picture         =   "aportes.frx":CB572
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3600
      Picture         =   "aportes.frx":CBAFC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8880
      Picture         =   "aportes.frx":CC086
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2400
      Picture         =   "aportes.frx":CC610
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      Picture         =   "aportes.frx":CCB9A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8400
      Picture         =   "aportes.frx":CD124
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   495
   End
End
Attribute VB_Name = "RegistrarAporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command4_Click()
Unload Me
End Sub

