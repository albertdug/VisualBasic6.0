VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Ayuda 
   Caption         =   "Transporte EL SOL // Ayuda"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Ayuda.frx":0000
   ScaleHeight     =   8025
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Sobre Los Menus"
      TabPicture(0)   =   "Ayuda.frx":55AE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Sobre los Datos"
      TabPicture(1)   =   "Ayuda.frx":55CA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label13"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label31"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label32"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label35"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label36"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label37"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label38"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label39"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label40"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Command2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Sobre Los Controles"
      TabPicture(2)   =   "Ayuda.frx":55E6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command3"
      Tab(2).Control(1)=   "Label30"
      Tab(2).Control(2)=   "Label29"
      Tab(2).Control(3)=   "Label28"
      Tab(2).Control(4)=   "Label27"
      Tab(2).Control(5)=   "Label26"
      Tab(2).Control(6)=   "Label25"
      Tab(2).Control(7)=   "Label24"
      Tab(2).Control(8)=   "Label23"
      Tab(2).Control(9)=   "Label22"
      Tab(2).Control(10)=   "Label21"
      Tab(2).Control(11)=   "Line1"
      Tab(2).Control(12)=   "Label20"
      Tab(2).Control(13)=   "Label19"
      Tab(2).Control(14)=   "Label18"
      Tab(2).Control(15)=   "Label17"
      Tab(2).Control(16)=   "Label16"
      Tab(2).Control(17)=   "Label15"
      Tab(2).Control(18)=   "Label14"
      Tab(2).Control(19)=   "Label12"
      Tab(2).Control(20)=   "Label11"
      Tab(2).Control(21)=   "Label10"
      Tab(2).Control(22)=   "Label9"
      Tab(2).Control(23)=   "Label8"
      Tab(2).Control(24)=   "Label7"
      Tab(2).ControlCount=   25
      Begin VB.CommandButton Command2 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -66960
         Picture         =   "Ayuda.frx":5602
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4680
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         Picture         =   "Ayuda.frx":5B8C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4680
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -66960
         Picture         =   "Ayuda.frx":6116
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4680
         Width           =   2175
      End
      Begin VB.Label Label40 
         Caption         =   "Clientes: Tanto Naturales -N- Como Juridicos -J-, Su codigo es la Cedula o RIF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   42
         Top             =   2400
         Width           =   9615
      End
      Begin VB.Label Label39 
         Caption         =   "Productos: El Codigo es Numerico, contentivo de Dos (2) Digitos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   41
         Top             =   2640
         Width           =   7575
      End
      Begin VB.Label Label38 
         Caption         =   "Repuestos: El Codigo en esta tabla, tambien es Numerico, Contiene Tres (3) Digitos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   40
         Top             =   2880
         Width           =   8535
      End
      Begin VB.Label Label37 
         Caption         =   $"Ayuda.frx":66A0
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74880
         TabIndex        =   39
         Top             =   3120
         Width           =   9615
      End
      Begin VB.Label Label36 
         Caption         =   $"Ayuda.frx":6754
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74880
         TabIndex        =   38
         Top             =   3600
         Width           =   9255
      End
      Begin VB.Label Label35 
         Caption         =   $"Ayuda.frx":6810
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74880
         TabIndex        =   37
         Top             =   4200
         Width           =   7695
      End
      Begin VB.Label Label32 
         Caption         =   $"Ayuda.frx":6911
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74880
         TabIndex        =   36
         Top             =   1080
         Width           =   9615
      End
      Begin VB.Label Label31 
         Caption         =   "Este Apéndice, quiere referirse a los Datos a Incluir en el Sistema."
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
         Left            =   -74880
         TabIndex        =   35
         Top             =   840
         Width           =   9855
      End
      Begin VB.Label Label13 
         Caption         =   "Los Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   34
         Top             =   480
         Width           =   9735
      End
      Begin VB.Label Label30 
         Caption         =   "Alt+N  Nueva (Factura)"
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
         Left            =   -70560
         TabIndex        =   1
         Top             =   4200
         Width           =   3255
      End
      Begin VB.Label Label29 
         Caption         =   "Alt+B   Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70560
         TabIndex        =   32
         Top             =   3960
         Width           =   3255
      End
      Begin VB.Label Label28 
         Caption         =   "Alt+L   Limpiar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70560
         TabIndex        =   31
         Top             =   3720
         Width           =   3255
      End
      Begin VB.Label Label27 
         Caption         =   "Alt+S   Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70560
         TabIndex        =   30
         Top             =   3480
         Width           =   3255
      End
      Begin VB.Label Label26 
         Caption         =   "Alt+E   Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70560
         TabIndex        =   29
         Top             =   3240
         Width           =   3255
      End
      Begin VB.Label Label25 
         Caption         =   "Alt+M   Modificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70560
         TabIndex        =   28
         Top             =   3000
         Width           =   3255
      End
      Begin VB.Label Label24 
         Caption         =   "Alt+G   Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70560
         TabIndex        =   27
         Top             =   2760
         Width           =   3255
      End
      Begin VB.Label Label23 
         Caption         =   "Alt+I   Incluir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70560
         TabIndex        =   26
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Label Label22 
         Caption         =   $"Ayuda.frx":6AF4
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -70560
         TabIndex        =   25
         Top             =   960
         Width           =   5655
      End
      Begin VB.Label Label21 
         Caption         =   "Comandos de Botones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70560
         TabIndex        =   24
         Top             =   600
         Width           =   5535
      End
      Begin VB.Line Line1 
         X1              =   -70800
         X2              =   -70800
         Y1              =   720
         Y2              =   4560
      End
      Begin VB.Label Label20 
         Caption         =   "Ctrl+F  Facturas Cobradas"
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
         Left            =   -74880
         TabIndex        =   23
         Top             =   4080
         Width           =   3255
      End
      Begin VB.Label Label19 
         Caption         =   "Ctrl+R  Repuestos Por Camión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   22
         Top             =   3840
         Width           =   3255
      End
      Begin VB.Label Label18 
         Caption         =   "Ctrl+U  Mantenimientos Requeridos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   3600
         Width           =   3255
      End
      Begin VB.Label Label17 
         Caption         =   "Ctrl+M  Mantenimientos Por Camión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   20
         Top             =   3360
         Width           =   3255
      End
      Begin VB.Label Label16 
         Caption         =   "Ctrl+T  Tabulador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   3120
         Width           =   3255
      End
      Begin VB.Label Label15 
         Caption         =   "Ctrl+P  Productos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   18
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Label Label14 
         Caption         =   "Ctrl+E  Servicios Efectuados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   17
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label Label12 
         Caption         =   "F5   Reportes"
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
         Left            =   -74880
         TabIndex        =   16
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "F4   Facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   15
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "F3   Repuestos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   14
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "F2   Camiones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   13
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "F1   Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Menu Principal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   11
         Top             =   600
         Width           =   9735
      End
      Begin VB.Label Label1 
         Caption         =   "Menu Principal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   9735
      End
      Begin VB.Label Label2 
         Caption         =   "El Menu principal, ofrece las opciones más importantes de este Sistema de Gestión."
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
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   9855
      End
      Begin VB.Label Label3 
         Caption         =   $"Ayuda.frx":6C31
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   9615
      End
      Begin VB.Label Label4 
         Caption         =   $"Ayuda.frx":6DAC
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   9615
      End
      Begin VB.Label Label5 
         Caption         =   "Menu de Seguridad Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3600
         Width           =   9615
      End
      Begin VB.Label Label6 
         Caption         =   $"Ayuda.frx":6F4F
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   5
         Top             =   3960
         Width           =   9615
      End
   End
   Begin VB.Label LBLayuda 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Ayuda"
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
      Left            =   8880
      TabIndex        =   0
      Top             =   720
      Width           =   1350
   End
End
Attribute VB_Name = "Ayuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub


