VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form AsignarPrestamo 
   Caption         =   "Generar Prestamo"
   ClientHeight    =   6630
   ClientLeft      =   3240
   ClientTop       =   3945
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "RegistrarSocios.frx":0000
   ScaleHeight     =   6630
   ScaleWidth      =   9390
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4200
      Picture         =   "RegistrarSocios.frx":CAE9E
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6240
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4800
      Picture         =   "RegistrarSocios.frx":CB428
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6240
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8160
      Picture         =   "RegistrarSocios.frx":CB9B2
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5640
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8880
      Picture         =   "RegistrarSocios.frx":CBF3C
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5760
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3600
      Picture         =   "RegistrarSocios.frx":CC4C6
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6240
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3600
      Picture         =   "RegistrarSocios.frx":CCA50
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      Picture         =   "RegistrarSocios.frx":CCFDA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   8535
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1680
         TabIndex        =   24
         Top             =   4440
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   7200
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   3120
         TabIndex        =   20
         Top             =   3360
         Width           =   495
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   1680
         TabIndex        =   18
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox Text8 
         Height          =   405
         Left            =   1680
         TabIndex        =   16
         Top             =   3360
         Width           =   375
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "RegistrarSocios.frx":CD564
         Left            =   3720
         List            =   "RegistrarSocios.frx":CD574
         TabIndex        =   13
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   2760
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "RegistrarSocios.frx":CD5BC
         Left            =   1680
         List            =   "RegistrarSocios.frx":CD5CC
         TabIndex        =   9
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   1680
         TabIndex        =   7
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   1680
         TabIndex        =   3
         Top             =   840
         Width           =   4935
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   2055
         Left            =   5520
         TabIndex        =   31
         Top             =   1440
         Width           =   2295
         _Version        =   524288
         _ExtentX        =   4048
         _ExtentY        =   3625
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2009
         Month           =   11
         Day             =   28
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimir Cheque"
         Height          =   195
         Left            =   6600
         TabIndex        =   32
         Top             =   4680
         Width           =   1125
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto del Cheque"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de Solicitud"
         Height          =   375
         Left            =   5640
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Solicitud"
         Height          =   615
         Left            =   4200
         TabIndex        =   21
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Tasa"
         Height          =   375
         Left            =   2400
         TabIndex        =   19
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto a Debitar"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuotas"
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Solicitado"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Prestamo"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   1245
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Disponibilidad"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "     Nombre"
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "      Cedula"
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "AsignarPrestamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_Click()
If Combo2.Text = "Especial" Then
Combo3.Visible = True
Label9.Visible = True
End If



If Combo2.Text = "Corto Plazo" Then
 Label10.Visible = True
 Text8.Visible = True
 Text8.Text = 12
 Label12.Visible = True
 Text10.Visible = True
 Text10.Text = 8
 Combo3.Visible = False
 Label9.Visible = False
 End If
 
 
 If Combo2.Text = "Largo Plazo" Then
 Combo3.Visible = False
 Label9.Visible = False
 Label10.Visible = True
 Text8.Visible = True
 Text8.Text = 60
 Label12.Visible = True
 Text10.Visible = True
 Text10.Text = 10
 End If
 
 If Combo2.Text = "Mediano Plazo" Then
 Combo3.Visible = False
 Label9.Visible = False
 Label10.Visible = True
 Text8.Visible = True
 Text8.Text = 24
 Label12.Visible = True
 Text10.Visible = True
 Text10.Text = 12
 End If
 
 If Combo2.Text = "Especial" Then
 Label10.Visible = True
 Text8.Visible = True
 Text8.Text = 36
 Label12.Visible = True
 Text10.Visible = True
 Text10.Text = 10
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Combo3.Visible = False
Label9.Visible = False
'Label10.Visible = False
'Label11.Visible = False
'Text8.Visible = False
'Text9.Visible = False
'Label12.Visible = False
'Text10.Visible = False
'Text4.Visible = False
' Label5.Visible = False
End Sub






Private Sub Text7_Change()
If Text7.Text <> " " And Combo2.Text <> "" Then
 Text9.Visible = True
 Label11.Visible = True
 Text4.Visible = True
 Label5.Visible = True
 
 End If
 
End Sub
