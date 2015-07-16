VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmcatalogoArticulos 
   Caption         =   "Catalogo de Articulos"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   615
      Left            =   2760
      Picture         =   "frmcatalogoArticulos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   3720
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3120
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   2895
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5106
      _Version        =   393216
      FixedCols       =   0
      FormatString    =   "Codigo   del  Articulo   |                                                         Descripción del Articulo"
   End
   Begin VB.Label Label2 
      Caption         =   "Descripción del Articulo:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo del Articulo:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
End
Attribute VB_Name = "frmcatalogoArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cn As New ADODB.Connection
Dim Res As Byte
Dim CodigoA As String
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
Sub colocarfrmarticulos(cadcodigo)
Dim TSQL2 As String
Dim TB2 As New ADODB.Recordset
Dim cadsexo As String
Dim cadestadoc As String

TSQL2 = "SELECT * FROM Articulos Where Codigo='" + cadcodigo + "'"
TB2.Open TSQL2, Cn
If Not TB2.EOF Then
   frmarticulos.Text1.Text = TB2("Codigo")
   frmarticulos.Text2.Text = TB2("Descripcion")
   frmarticulos.Text3.Text = CLng(TB2("Existencia"))
   frmarticulos.Text4.Text = CCur(TB2("Preciou"))

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
   


    TSQL = "SELECT * FROM Articulos ORDER BY Val(Codigo)"

    TB.Open TSQL, Cn
     FG.Rows = cantrecord("select * from Articulos") + 1
     
    If Not TB.EOF Then
      
        
        
        i = 1
        
        While Not TB.EOF
        
          FG.TextMatrix(i, 0) = TB("Codigo")
          FG.TextMatrix(i, 1) = TB("Descripcion")
          TB.MoveNext
          i = i + 1
        Wend
        
              
    
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Pos As Integer
Pos = FG.Row
CodigoA = FG.TextMatrix(Pos, 0)
colocarfrmarticulos (CodigoA)
Cn.Close
Unload Me
frmarticulos.Visible = True
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then
   Dim i As Byte
   Dim TSQL As String
   Dim TB As New ADODB.Recordset
   Dim Cad As String
   
  
   

   TSQL = "SELECT * FROM Articulos WHERE Codigo LIKE  '" & Text1.Text & "%' Order by Val(Codigo)"


   TB.Open TSQL, Cn
   FG.Rows = cantrecord(TSQL) + 1
   
   If Not TB.EOF Then
      
        
        
        i = 1
        
        While Not TB.EOF
        
          FG.TextMatrix(i, 0) = TB("Codigo")
          FG.TextMatrix(i, 1) = TB("Descripcion")
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
   
  
   

   TSQL = "SELECT * FROM Articulos WHERE Descripcion LIKE  '" & Text2.Text & "%' Order by Descripcion"


   TB.Open TSQL, Cn
   FG.Rows = cantrecord(TSQL) + 1
   
   If Not TB.EOF Then
      
        
        
        i = 1
        
        While Not TB.EOF
        
          FG.TextMatrix(i, 0) = TB("Codigo")
          FG.TextMatrix(i, 1) = TB("Descripcion")
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

