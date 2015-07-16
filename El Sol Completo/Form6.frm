VERSION 5.00
Begin VB.Form Acercade 
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   4890
   ClientLeft      =   3780
   ClientTop       =   2400
   ClientWidth     =   8625
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   6887.324
   ScaleMode       =   0  'User
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   501
      Left            =   6240
      Picture         =   "Form6.frx":6E9C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   2000
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   6360
      Picture         =   "Form6.frx":7426
      Top             =   1200
      Width           =   1605
   End
End
Attribute VB_Name = "Acercade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

