VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Inicial2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Transporte EL SOL // Ingrese Fecha de Hoy"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Inicial2.frx":0000
   ScaleHeight     =   4890
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTFechaIn 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   2880
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   68419585
      CurrentDate     =   39947
   End
   Begin VB.CommandButton Accept 
      Caption         =   "&Aceptar"
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
      Picture         =   "Inicial2.frx":6F52
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label3 
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
      Left            =   5640
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
End
Attribute VB_Name = "Inicial2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tempfact As New ADODB.Recordset
Dim t As New ADODB.Recordset
Private Sub Accept_Click()
GLOBALDATE = DTFechaIn.Value

    Conexion.Execute " update ServiciosEfectuados set EstatusSef = 'R' where #" & GLOBALDATE & "# >= FecServicio and EstatusSef = 'A' "
    Conexion.Execute " update ServiciosEfectuados set EstatusSef = 'A' where #" & GLOBALDATE & "# <= FecServicio and EstatusSef = 'R' "
    Conexion.Execute " update CamionesTemp set EstatusCP = 'R' where #" & GLOBALDATE & "# > FechaP and EstatusCP = 'A'"
If Format(GLOBALDATE, "dddd") = "Viernes" Then
    MsgBox (" Comenzara El Proceso De Facturacion Automatico, Tendra Que Esperar Un Momento ")
    tempfact.Open "select * from ServiciosEfectuados ", Conexion
    t.Open " select MAX(NumFact) from facturas", Conexion
    If Not IsNumeric(t(0)) Then
       cont = "001"
       Else
     cont = Format(Val(t(0)) + 1, "000")
    End If
    
While Not tempfact.EOF
  
    
  If (tempfact(2) - GLOBALDATE) <= 0 Then

    Conexion.Execute "insert into facturas(Numfact,CodClienteF,FecVencim,CodCiudOrig,CodCiudDest,CodProd,CantKg,FechaSolicitud,FechaServicio,Placa,EstatusFac) values(" _
    & " '" & cont & "','" & tempfact(1) & "','" & (tempfact(2) + 15) & "','" & tempfact(3) & "','" & tempfact(4) & "','" & tempfact(5) & "','" & tempfact(6) & "','" & tempfact(7) & "','" & tempfact(2) & "','" & tempfact(8) & "','A')"
    Conexion.Execute "delete * from ServiciosEfectuados where EstatusSef = 'R'"
  cont = Format(cont + 1, "000")
  End If
tempfact.MoveNext

 Wend
 MsgBox " Operacion Terminada", vbInformation, " Proceso Terminado"
 End If
 
 Conexion.Execute "delete * from CamionesTemp where EstatusCp = 'R'"

Principal.Show
Unload Me

End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
Label3 = Date
End Sub


