VERSION 5.00
Begin VB.Form simularprestamo 
   Caption         =   "Simular Prestamo"
   ClientHeight    =   6630
   ClientLeft      =   3240
   ClientTop       =   3945
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "simularprestamo.frx":0000
   ScaleHeight     =   6630
   ScaleWidth      =   9375
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8760
      Picture         =   "simularprestamo.frx":CAE9E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7920
      Picture         =   "simularprestamo.frx":CB428
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5280
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " "
      Height          =   4575
      Left            =   720
      TabIndex        =   0
      Top             =   1200
      Width           =   7815
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Text            =   " "
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   4800
         TabIndex        =   12
         Text            =   " "
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Text            =   " "
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   4200
         TabIndex        =   8
         Text            =   " "
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Text            =   " "
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Text            =   " "
         Top             =   2760
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "simularprestamo.frx":CB9B2
         Left            =   1920
         List            =   "simularprestamo.frx":CB9C2
         TabIndex        =   2
         Text            =   " "
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto de la Cuota"
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "El Monto del Cheque sara de"
         Height          =   375
         Left            =   3600
         TabIndex        =   11
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Disponibilidad Neta"
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Taza de Interes"
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de cuotas"
         Height          =   615
         Left            =   480
         TabIndex        =   5
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Prestamo"
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   1440
         Width           =   1455
      End
   End
End
Attribute VB_Name = "simularprestamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
If Combo1.Text = "Corto Plazo" Then
 Text2.Text = 12
 Text3.Text = 8
 End If
 
 If Combo1.Text = "Mediano Plazo" Then
 Text2.Text = 24
 Text3.Text = 12
 End If
 
 If Combo1.Text = "Largo Plazo" Then
 Text2.Text = 60
 Text3.Text = 10
 End If
 
 If Combo1.Text = "Especial" Then
 Text2.Text = 36
 Text3.Text = 10
 End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
