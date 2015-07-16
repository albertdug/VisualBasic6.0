VERSION 5.00
Begin VB.Form generarestadodecuenta 
   Caption         =   "Generar Estado de Cuenta"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "generarestadodecuenta.frx":0000
   ScaleHeight     =   6675
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4920
      Picture         =   "generarestadodecuenta.frx":CAE9E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Text            =   " "
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cedula"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "generarestadodecuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
EstadodeCuenta.Show
End Sub
