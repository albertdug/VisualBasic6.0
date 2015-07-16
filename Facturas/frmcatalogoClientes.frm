VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmcatalogoClientes 
   Caption         =   "Catalogo de Clientes"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   9795
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   4080
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   615
      Left            =   3120
      Picture         =   "frmcatalogoClientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   2895
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5106
      _Version        =   393216
      FixedCols       =   0
      FormatString    =   "Cedula del  Cliente          |                                                         Nombre  del   Cliente"
   End
   Begin VB.Label Label1 
      Caption         =   "Cedula del  Cliente:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre del  Cliente:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   4080
      Width           =   1695
   End
End
Attribute VB_Name = "frmcatalogoClientes"
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
   frmclientes.Text1.Text = TB2("Cedula")
   frmclientes.Text2.Text = TB2("Nombre")
   frmclientes.DTFecha.Value = TB2("Fecnac")
   cadsexo = TB2("Sexo")
   
   If cadsexo = "M" Then
      frmclientes.OptionM.Value = True
   Else
      frmclientes.OptionF.Value = True
   End If
   
   cadestadoc = TB2("Estadoc")
   
   Select Case cadestadoc
          Case "Soltero"
               frmclientes.OptionS.Value = True
               frmclientes.OptionC.Value = False
               frmclientes.OptionV.Value = False
               frmclientes.OptionD.Value = False
          Case "Casado"
               frmclientes.OptionS.Value = False
               frmclientes.OptionC.Value = True
               frmclientes.OptionV.Value = False
               frmclientes.OptionD.Value = False
          Case "Viudo"
               frmclientes.OptionS.Value = False
               frmclientes.OptionC.Value = False
               frmclientes.OptionV.Value = True
               frmclientes.OptionD.Value = False
         Case Else
               frmclientes.OptionS.Value = False
               frmclientes.OptionC.Value = False
               frmclientes.OptionV.Value = False
               frmclientes.OptionD.Value = True
      End Select
      
      frmclientes.Text3.Text = TB2("Telefono")
      frmclientes.Text4.Text = TB2("Direccion")

End If
TB2.Close

End Sub

Private Sub Command1_Click()
Unload Me

End Sub
Private Sub Form_Load()
  Cn.Open "DSN=sisfact; UID=Admin; PWD=123;"
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
frmclientes.Visible = True
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
