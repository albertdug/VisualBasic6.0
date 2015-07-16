VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form registrarcuentabancaria 
   Caption         =   "Registrar Cuenta Bancaria"
   ClientHeight    =   6660
   ClientLeft      =   3240
   ClientTop       =   3540
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "registrarcuentabancaria.frx":0000
   ScaleHeight     =   6660
   ScaleWidth      =   9375
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      Picture         =   "registrarcuentabancaria.frx":CAE9E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3480
      Picture         =   "registrarcuentabancaria.frx":CB428
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8760
      Picture         =   "registrarcuentabancaria.frx":CB9B2
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      Picture         =   "registrarcuentabancaria.frx":CBF3C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      Picture         =   "registrarcuentabancaria.frx":CC4C6
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4440
      Picture         =   "registrarcuentabancaria.frx":CCA50
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1920
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   8895
      Begin MSACAL.Calendar Calendar1 
         Height          =   2175
         Left            =   6240
         TabIndex        =   10
         Top             =   360
         Width           =   2295
         _Version        =   524288
         _ExtentX        =   4048
         _ExtentY        =   3836
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2009
         Month           =   11
         Day             =   28
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Text            =   " "
         Top             =   2640
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "registrarcuentabancaria.frx":CCFDA
         Left            =   2040
         List            =   "registrarcuentabancaria.frx":CCFE4
         TabIndex        =   6
         Text            =   " "
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Text            =   " "
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Text            =   " "
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Apertura"
         Height          =   375
         Left            =   4800
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cuenta"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de Cuenta"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
   End
End
Attribute VB_Name = "registrarcuentabancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
Unload Me
End Sub

