VERSION 5.00
Begin VB.Form RegistrarBancos 
   ClientHeight    =   6660
   ClientLeft      =   3240
   ClientTop       =   3735
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "RegistrarBancos.frx":0000
   ScaleHeight     =   6660
   ScaleWidth      =   9360
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2400
      Picture         =   "RegistrarBancos.frx":CAE9E
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      Picture         =   "RegistrarBancos.frx":CB428
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8760
      Picture         =   "RegistrarBancos.frx":CB9B2
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3480
      Picture         =   "RegistrarBancos.frx":CBF3C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      Picture         =   "RegistrarBancos.frx":CC4C6
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      Picture         =   "RegistrarBancos.frx":CCA50
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   8535
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Text            =   " "
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Text            =   "Codigos"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   2640
         TabIndex        =   8
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   2040
         Width           =   4575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo del Banco"
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "    Direccion"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "  Rif"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "        Nombre"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   1455
      End
   End
End
Attribute VB_Name = "RegistrarBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command4_Click()
Unload Me
End Sub
