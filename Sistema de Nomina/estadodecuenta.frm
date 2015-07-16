VERSION 5.00
Begin VB.Form EstadodeCuenta 
   Caption         =   "Historial"
   ClientHeight    =   10320
   ClientLeft      =   120
   ClientTop       =   660
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "estadodecuenta.frx":0000
   ScaleHeight     =   10320
   ScaleWidth      =   15120
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   14400
      Picture         =   "estadodecuenta.frx":204C42
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9360
      Width           =   495
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DISPONIBILIDAD NETA............................:                       2337.108"
      Height          =   195
      Left            =   5040
      TabIndex        =   24
      Top             =   9480
      Width           =   4785
   End
   Begin VB.Line Line7 
      X1              =   8040
      X2              =   9840
      Y1              =   9360
      Y2              =   9360
   End
   Begin VB.Line Line6 
      X1              =   10680
      X2              =   12000
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line5 
      X1              =   9120
      X2              =   10440
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL PRESTAMOS................................:                            .00"
      Height          =   195
      Left            =   5040
      TabIndex        =   23
      Top             =   9120
      Width           =   4545
   End
   Begin VB.Line Line4 
      X1              =   3120
      X2              =   12120
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Line Line3 
      X1              =   3120
      X2              =   12120
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"estadodecuenta.frx":2051CC
      Height          =   195
      Left            =   3120
      TabIndex        =   22
      Top             =   8160
      Width           =   8970
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION     PRESTAMO     MONTO     PRESTAMO     MONTO          SALDO     FECHA     FECHA     CUOTA     NROS"
      Height          =   195
      Left            =   3120
      TabIndex        =   21
      Top             =   7800
      Width           =   8955
   End
   Begin VB.Line Line2 
      X1              =   3120
      X2              =   12120
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10/02/2005"
      Height          =   195
      Left            =   7560
      TabIndex        =   20
      Top             =   4560
      Width           =   870
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Albert R. Durán G."
      Height          =   195
      Left            =   6960
      TabIndex        =   19
      Top             =   4080
      Width           =   1305
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "03/12/09"
      Height          =   195
      Left            =   4320
      TabIndex        =   18
      Top             =   3000
      Width           =   690
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Sargento"
      Height          =   375
      Left            =   4200
      TabIndex        =   17
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "       19.432.407"
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   $"estadodecuenta.frx":205262
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   14
      Top             =   7200
      Width           =   9735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   $"estadodecuenta.frx":2052FC
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   13
      Top             =   6960
      Width           =   9615
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "AHORROS ACUMULADOS VOLUNTARIOS....:                               .00                          .00                             .00"
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   12
      Top             =   6360
      Width           =   9375
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   $"estadodecuenta.frx":205390
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   6120
      Width           =   9375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   $"estadodecuenta.frx":205419
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   5880
      Width           =   9375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "HAB. REALES"
      Height          =   255
      Left            =   10680
      TabIndex        =   9
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "HAB. TOTALES"
      Height          =   255
      Left            =   9120
      TabIndex        =   8
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "CUOTA"
      Height          =   255
      Left            =   7800
      TabIndex        =   7
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "AHORROS SOCIOS ACTUALIZADOS AL...:   03/12/09"
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   5280
      Width           =   4455
   End
   Begin VB.Line Line1 
      X1              =   3120
      X2              =   12120
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "           FECHA:"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Ingreso:"
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Rango:"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   6000
      TabIndex        =   2
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cedula:"
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   """Estado de Cuenta de Personal"""
      Height          =   255
      Left            =   4680
      TabIndex        =   0
      Top             =   3480
      Width           =   6375
   End
End
Attribute VB_Name = "EstadodeCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DBGrid1_Click()

End Sub

