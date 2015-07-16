VERSION 5.00
Begin VB.Form GuardarTipoPrestamo 
   Caption         =   "Registrar Nuevo Prestamo"
   ClientHeight    =   6675
   ClientLeft      =   3030
   ClientTop       =   3945
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "GuardarTipoPrestamo.frx":0000
   ScaleHeight     =   6675
   ScaleWidth      =   9405
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      Picture         =   "GuardarTipoPrestamo.frx":CAE9E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      Picture         =   "GuardarTipoPrestamo.frx":CB428
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8880
      Picture         =   "GuardarTipoPrestamo.frx":CB9B2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5760
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2760
      Picture         =   "GuardarTipoPrestamo.frx":CBF3C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      Picture         =   "GuardarTipoPrestamo.frx":CC4C6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   8535
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3960
         Picture         =   "GuardarTipoPrestamo.frx":CCA50
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   2040
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de Cuotas"
         Height          =   495
         Left            =   720
         TabIndex        =   4
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tasa de Interes"
         Height          =   495
         Left            =   720
         TabIndex        =   2
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   495
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "GuardarTipoPrestamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Unload Me
End Sub
