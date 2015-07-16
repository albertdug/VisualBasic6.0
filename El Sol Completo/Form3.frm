VERSION 5.00
Begin VB.Form Objetivos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Transporte EL SOL // Objetivos"
   ClientHeight    =   4590
   ClientLeft      =   3360
   ClientTop       =   3435
   ClientWidth     =   8505
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   4266.166
   ScaleMode       =   0  'User
   ScaleWidth      =   8505
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
      Height          =   533
      Left            =   6240
      Picture         =   "Form3.frx":84F8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   2000
   End
End
Attribute VB_Name = "Objetivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub
