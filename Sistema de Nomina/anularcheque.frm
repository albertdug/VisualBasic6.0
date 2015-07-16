VERSION 5.00
Begin VB.Form anularcheque 
   Caption         =   "Anular Cheque"
   ClientHeight    =   6660
   ClientLeft      =   3030
   ClientTop       =   4155
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "anularcheque.frx":0000
   ScaleHeight     =   6660
   ScaleWidth      =   9375
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      ForeColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   8655
      Begin VB.TextBox Text8 
         Height          =   405
         Left            =   6360
         TabIndex        =   23
         Text            =   " "
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   2040
         TabIndex        =   14
         Text            =   " "
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Text            =   " "
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Text            =   " "
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Text            =   " "
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Text            =   " "
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Text            =   " "
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3840
         Picture         =   "anularcheque.frx":CAE9E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6360
         TabIndex        =   7
         Text            =   " "
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Anulacion"
         Height          =   375
         Left            =   4680
         TabIndex        =   22
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
         Height          =   375
         Left            =   480
         TabIndex        =   21
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de Cuenta"
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Emision"
         Height          =   375
         Left            =   4680
         TabIndex        =   19
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
         Height          =   375
         Left            =   480
         TabIndex        =   18
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "A Nombre de"
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo de Chequera"
         Height          =   495
         Left            =   480
         TabIndex        =   16
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo del Cheque"
         Height          =   495
         Left            =   480
         TabIndex        =   15
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4680
      Picture         =   "anularcheque.frx":CB428
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4080
      Picture         =   "anularcheque.frx":CB9B2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8760
      Picture         =   "anularcheque.frx":CBF3C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3480
      Picture         =   "anularcheque.frx":CC4C6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      Picture         =   "anularcheque.frx":CCA50
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      Picture         =   "anularcheque.frx":CCFDA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   495
   End
End
Attribute VB_Name = "anularcheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
Unload Me
End Sub
