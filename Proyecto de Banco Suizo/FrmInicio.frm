VERSION 5.00
Begin VB.Form FrmInicio 
   BackColor       =   &H00C00000&
   Caption         =   "Inicio"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10035
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmbEntrar 
      BackColor       =   &H00C00000&
      Caption         =   "Entrar"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox TxtContraseña 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label LblInsertClave 
      BackColor       =   &H00C00000&
      Caption         =   "Inserte Clave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Lbl1 
      BackColor       =   &H00C00000&
      Caption         =   "Bank Suizo"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1695
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   6495
   End
End
Attribute VB_Name = "FrmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cont As Byte
Private Sub CmbEntrar_Click()
If TxtContraseña.Text = "1976" Then
   FrmBancoSuizo.Show
   Unload Me
   Else
      cont = cont + 1
      restan = 3 - cont
      MsgBox "Le quedan:0" & restan & "opciones"
      TxtContraseña = ""
      TxtContraseña.SetFocus
      If cont = 3 Then
         MsgBox "No sabe la contraseña"
         Unload Me
      End If
End If
End Sub

