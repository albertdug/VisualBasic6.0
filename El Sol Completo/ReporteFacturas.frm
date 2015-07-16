VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ReporteFacturas 
   Caption         =   "Reporte De Facturas"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   Picture         =   "ReporteFacturas.frx":0000
   ScaleHeight     =   8625
   ScaleWidth      =   12870
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTP2 
      Height          =   375
      Left            =   8520
      TabIndex        =   6
      Top             =   2640
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   23855105
      CurrentDate     =   39958
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   23855105
      CurrentDate     =   39958
   End
   Begin MSAdodcLib.Adodc Adodcfac 
      Height          =   495
      Left            =   2280
      Top             =   8280
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\El Sol Completo\Base de datos1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\El Sol Completo\Base de datos1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select Numfact, CodClientef, NomCli, FecVenCim from Facturas, Clientes where CodClienteF = CodCli "
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton CMDMos 
      Caption         =   "&Mostrar "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Picture         =   "ReporteFacturas.frx":55AE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton CMDsal 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      Picture         =   "ReporteFacturas.frx":5B38
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "ReporteFacturas.frx":60C2
      Height          =   3735
      Left            =   1920
      TabIndex        =   2
      Top             =   4200
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6588
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Reporte De Facturas"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Numfact"
         Caption         =   "Nro Factura"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "CodClientef"
         Caption         =   "Codigo Cliente"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "NomCli"
         Caption         =   "Nombre Cliente"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "FecVenCim"
         Caption         =   "Fecha Vencimiento"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "de Facturas"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8880
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label LBLclientes 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Reporte: Vencimiento"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6840
      TabIndex        =   0
      Top             =   360
      Width           =   4530
   End
End
Attribute VB_Name = "ReporteFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDMos_Click()
 fecha5 = Str(DTP1.Year) & "/" & Str(DTP1.Month) & "/" & Str(DTP1.Day)
 fecha6 = Str(DTP2.Year) & "/" & Str(DTP2.Month) & "/" & Str(DTP2.Day)
Adodcfac.RecordSource = " Select Numfact, CodClientef, NomCli, FecVenCim from Facturas, Clientes where CodClienteF = CodCli and fecvencim between  # " & fecha5 & " # and # " & fecha6 & " #  "
Adodcfac.Refresh
End Sub



Private Sub CMDsal_Click()
If MsgBox(" Esta Seguro que Desea Salir de Este Reporte?", vbQuestion + vbYesNo, "Salida de Pantalla") = vbYes Then
Unload Me
End If
End Sub

Private Sub DTP1_Change()
DTP2.Value = DTP1.Value + 3
End Sub

Private Sub Form_Load()
DTP1.Value = GLOBALDATE

Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Private Sub Command3_Click()
 
 Else
  adofacven.RecordSource = " select numfact, facturas.codcli, nomcli, i.ciudad, d.ciudad, nomprod, cantkg, fecsolicitud, fecprobserv, fecvencim from facturas, clientes, ciudades as i, ciudades as d, productos where clientes.codcli = facturas.codcli and i.codciudad=facturas.codciudadinicio and d.codciudad=facturas.codciudaddestino and productos.codprod = facturas.codprod and (facturas.fecvencim >= #" & fecha5 & "#) and (facturas.fecvencim <= #" & fecha6 & "#) "
  adofacven.Refresh
  End If
End Sub
