VERSION 5.00
Begin VB.Form Generarusuario 
   Caption         =   "Usuario"
   ClientHeight    =   6660
   ClientLeft      =   3240
   ClientTop       =   3945
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Generarusuario.frx":0000
   ScaleHeight     =   6660
   ScaleWidth      =   9360
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8760
      Picture         =   "Generarusuario.frx":CAE9E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6000
      Picture         =   "Generarusuario.frx":CB428
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   840
      TabIndex        =   0
      Top             =   1680
      Width           =   7695
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Text            =   " "
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Text            =   " "
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "las claves no son iguales"
         Height          =   375
         Left            =   5280
         TabIndex        =   7
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirmar Clave"
         Height          =   495
         Left            =   840
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave"
         Height          =   495
         Left            =   1200
         TabIndex        =   3
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "Generarusuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Unload Me
End Sub
