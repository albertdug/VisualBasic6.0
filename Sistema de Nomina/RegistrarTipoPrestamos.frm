VERSION 5.00
Begin VB.Form RegistrarTipoPrestamo 
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Tasa de Interes"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "RegistrarTipoPrestamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Unload Me
End Sub
