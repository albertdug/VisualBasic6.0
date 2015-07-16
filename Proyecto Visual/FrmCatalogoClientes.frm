VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCatalogoClientes 
   BackColor       =   &H00FF00FF&
   Caption         =   "Catalogo de Clientes"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   5520
      TabIndex        =   4
      Top             =   4320
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5520
      TabIndex        =   3
      Top             =   3360
      Width           =   3255
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   1935
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   3413
      _Version        =   393216
      FixedCols       =   0
      FormatString    =   "Cedula Del Cliente          |         Nombre Del Cliente"
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Nombre Del Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Cedula Del Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   3360
      Width           =   2775
   End
End
Attribute VB_Name = "FrmCatalogoClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cn As New ADODB.Connection
Dim Res As Byte
Dim CedulaA As String
Function cantrecord(Cad) As Integer
Dim i As Integer
Dim TSQL1 As String
Dim TB1 As New ADODB.Recordset
Dim cont As Integer


    TSQL1 = Cad

    TB1.Open TSQL1, Cn
    
    If Not TB1.EOF Then
      
        
        
        cont = 0
        
        While Not TB1.EOF
          cont = cont + 1
          TB1.MoveNext
        Wend
        
    Else
      cont = 0
    End If
    
cantrecord = cont
End Function
Sub colocarfrmclientes(cadcedula)
Dim TSQL2 As String
Dim TB2 As New ADODB.Recordset
Dim cadsexo As String
Dim cadestadoc As String

TSQL2 = "SELECT * FROM Clientes Where Cedula='" + cadcedula + "'"
TB2.Open TSQL2, Cn
If Not TB2.EOF Then
   FrmClientes.Text1.Text = TB2("Cedula")
   FrmClientes.Text2.Text = TB2("Nombre")
   FrmClientes.DTFecha.Value = TB2("Fecnac")
   cadsexo = TB2("Sexo")
   
   If cadsexo = "M" Then
      FrmClientes.OptionM.Value = True
   Else
      FrmClientes.OptionF.Value = True
   End If
   
   cadestadoc = TB2("Estadoc")
   
   Select Case cadestadoc
          Case "Soltero"
               FrmClientes.OptionS.Value = True
               FrmClientes.OptionC.Value = False
               FrmClientes.OptionV.Value = False
               FrmClientes.OptionD.Value = False
          Case "Casado"
               FrmClientes.OptionS.Value = False
               FrmClientes.OptionC.Value = True
               FrmClientes.OptionV.Value = False
               FrmClientes.OptionD.Value = False
          Case "Viudo"
               FrmClientes.OptionS.Value = False
               FrmClientes.OptionC.Value = False
               FrmClientes.OptionV.Value = True
               FrmClientes.OptionD.Value = False
         Case Else
               FrmClientes.OptionS.Value = False
               FrmClientes.OptionC.Value = False
               FrmClientes.OptionV.Value = False
               FrmClientes.OptionD.Value = True
      End Select
      
      FrmClientes.Text3.Text = TB2("Telefono")
      FrmClientes.Text4.Text = TB2("Direccion")

End If
TB2.Close

End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
  Cn.Open "DSN=sisfact; UID=Admin; PWD=19324551;"
  Dim i As Byte
  Dim TSQL As String
   Dim TB As New ADODB.Recordset
   


    TSQL = "SELECT * FROM Clientes ORDER BY Val(Cedula)"

    TB.Open TSQL, Cn
     FG.Rows = cantrecord("select * from Clientes") + 1
     
    If Not TB.EOF Then
      
        
        
        i = 1
        
        While Not TB.EOF
        
          FG.TextMatrix(i, 0) = TB("Cedula")
          FG.TextMatrix(i, 1) = TB("Nombre")
          TB.MoveNext
          i = i + 1
        Wend
        
              
    
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Pos As Integer
Pos = FG.Row
CedulaA = FG.TextMatrix(Pos, 0)
colocarfrmclientes (CedulaA)
Cn.Close
Unload Me
FrmClientes.Visible = True
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then
   Dim i As Byte
   Dim TSQL As String
   Dim TB As New ADODB.Recordset
   Dim Cad As String
   
  
   

   TSQL = "SELECT * FROM Clientes WHERE Cedula LIKE  '" & Text1.Text & "%' Order by Val(Cedula)"


   TB.Open TSQL, Cn
   FG.Rows = cantrecord(TSQL) + 1
   
   If Not TB.EOF Then
      
        
        
        i = 1
        
        While Not TB.EOF
        
          FG.TextMatrix(i, 0) = TB("Cedula")
          FG.TextMatrix(i, 1) = TB("Nombre")
          TB.MoveNext
          i = i + 1
        Wend
    Else
     Res = MsgBox("No Hay Datos", 64, "Información")
     Text1.Text = Mid(Text1.Text, 1, Len(Text1.Text) - 1)
     Text1.SelStart = Len(Text1.Text)
    End If

End If
End Sub

Private Sub Text2_Change()
Cadena = QBlancos(Text2.Text)
Text2.Text = Cadena
Text2.SelStart = Len(Text2.Text)
If Text2.Text <> "" Then
   Dim i As Byte
   Dim TSQL As String
   Dim TB As New ADODB.Recordset
   Dim Cad As String
   
  
   

   TSQL = "SELECT * FROM Clientes WHERE Nombre LIKE  '" & Text2.Text & "%' Order by Nombre"


   TB.Open TSQL, Cn
   FG.Rows = cantrecord(TSQL) + 1
   
   If Not TB.EOF Then
      
        
        
        i = 1
        
        While Not TB.EOF
        
          FG.TextMatrix(i, 0) = TB("Cedula")
          FG.TextMatrix(i, 1) = TB("Nombre")
          TB.MoveNext
          i = i + 1
        Wend
    Else
     Res = MsgBox("No Hay Datos", 64, "Información")
     Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
     Text2.SelStart = Len(Text2.Text)
    End If

End If
End Sub
