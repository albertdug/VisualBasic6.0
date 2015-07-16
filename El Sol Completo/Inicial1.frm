VERSION 5.00
Begin VB.Form Inicial1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Transporte EL SOL // Ingrese Usuario"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Inicial1.frx":0000
   ScaleHeight     =   4491.743
   ScaleMode       =   0  'User
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMBsal 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Picture         =   "Inicial1.frx":6DC2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdlimpiar 
      Caption         =   "&Limpiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Picture         =   "Inicial1.frx":734C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdEntrar 
      Caption         =   "&Entrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Picture         =   "Inicial1.frx":78D6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtclav 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3000
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3120
      Width           =   2895
   End
   Begin VB.TextBox txtnom 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2640
      Width           =   2895
   End
End
Attribute VB_Name = "Inicial1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cont As Integer

Private Sub cmblimpiar_Click()
TXTnom.Text = ""
txtclav.Text = ""
End Sub

Private Sub CMBsal_Click()
If MsgBox(" Aun no ha empezado Sesion, Desea Realmemte Salir?", vbQuestion + vbYesNo, "Salida de Programa") = vbYes Then
Unload Me
End If
End Sub

Private Sub cmdEntrar_Click()
Dim t As New ADODB.Recordset
t.Open "select * from usuarios where nombre= '" & TXTnom.Text & "' and codigo = '" & txtclav.Text & "'", Conexion
If TXTnom.Text = "" Or txtclav.Text = "" Then
    MsgBox " No Debe Dejar Espacios En Blanco", vbInformation, "Omision de Datos"
Else
    If t.EOF Then
        MsgBox "Acceso Denegado: Clave o Nombre De Usuario Incorrectos", vbInformation, "Datos Incorrectos"
        cont = cont + 1
        restan = 3 - cont
        If MsgBox("Le Restan " & restan & " Oportunidades para Acceder, Reintentelo", vbRetryCancel + vbCritical, "Entrada al Sistema") = vbRetry Then
        cmblimpiar_Click
        TXTnom.SetFocus
        Else
        MsgBox "Se Cierra Sesión del Sistema Transporte EL SOL", vbSystemModal, "Cierre de Sesión"
        Unload Me
        End If
    If cont = 3 Then
        MsgBox "Usted No Sabe u Olvido Su Clave... El Sistema se Cerrará por Seguridad "
        Unload Me
    End If
    Else
    Unload Me
    Inicial2.Show
End If
End If
End Sub

Private Sub Form_Load()
Conectar
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub
Private Sub txtclav_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdEntrar_Click
End If
End Sub
Private Sub txtnom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtclav.SetFocus
End If
End Sub
