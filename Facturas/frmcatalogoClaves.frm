VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmcatalogoClaves 
   Caption         =   "Actualización de Claves"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   495
      Left            =   3360
      Picture         =   "frmcatalogoClaves.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   4560
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   3120
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   2175
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      FormatString    =   " Clave        |  Usuario                      |                                                        Nombre del  Usuario"
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre del  Usuario :"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Usuario :"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Clave del  Usuario :"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   3120
      Width           =   1575
   End
End
Attribute VB_Name = "frmcatalogoClaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cn As New ADODB.Connection
Dim Res As Byte
Dim ClaveA As String
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
Sub colocarfrmclaves(cadclave)
Dim TSQL2 As String
Dim TB2 As New ADODB.Recordset
Dim cadnivel As String

TSQL2 = "SELECT * FROM Usuarios Where Clave='" + cadclave + "'"
TB2.Open TSQL2, Cn
If Not TB2.EOF Then
   frmclaves.Text1.Text = TB2("Clave")
   frmclaves.Text2.Text = TB2("Usuario")
   frmclaves.Text3.Text = TB2("Nombre")
   cadnivel = TB2("Nivel")
   
   If cadnivel = "1" Then
      frmclaves.Option1.Value = True
   Else
      frmclaves.Option2.Value = True
   End If

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
   


    TSQL = "SELECT * FROM Usuarios ORDER BY Val(Clave)"

    TB.Open TSQL, Cn
     FG.Rows = cantrecord("select * from Usuarios") + 1
     
    If Not TB.EOF Then
      
        
        
        i = 1
        
        While Not TB.EOF
        
          FG.TextMatrix(i, 0) = TB("Clave")
          FG.TextMatrix(i, 1) = TB("Usuario")
          FG.TextMatrix(i, 2) = TB("Nombre")
          TB.MoveNext
          i = i + 1
        Wend
        
              
    
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Pos As Integer
Pos = FG.Row
ClaveA = FG.TextMatrix(Pos, 0)
colocarfrmclaves (ClaveA)
Cn.Close
Unload Me
frmclaves.Visible = True
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then
   Dim i As Byte
   Dim TSQL As String
   Dim TB As New ADODB.Recordset
   Dim Cad As String
   
  
   

   TSQL = "SELECT * FROM Usuarios WHERE Clave LIKE  '" & Text1.Text & "%' Order by Val(Clave)"


   TB.Open TSQL, Cn
   FG.Rows = cantrecord(TSQL) + 1
   
   If Not TB.EOF Then
      
        
        
        i = 1
        
        While Not TB.EOF
        
          FG.TextMatrix(i, 0) = TB("Clave")
          FG.TextMatrix(i, 1) = TB("Usuario")
          FG.TextMatrix(i, 2) = TB("Nombre")
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
If Text2.Text <> "" Then
   Dim i As Byte
   Dim TSQL As String
   Dim TB As New ADODB.Recordset
   Dim Cad As String
   
  
   

   TSQL = "SELECT * FROM Usuarios WHERE Usuario LIKE  '" & Text2.Text & "%' Order by Usuario"


   TB.Open TSQL, Cn
   FG.Rows = cantrecord(TSQL) + 1
   
   If Not TB.EOF Then
      
        
        
        i = 1
        
        While Not TB.EOF
        
          FG.TextMatrix(i, 0) = TB("Clave")
          FG.TextMatrix(i, 1) = TB("Usuario")
          FG.TextMatrix(i, 2) = TB("Nombre")
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
Private Sub Text3_Change()
Cadena = QBlancos(Text3.Text)
Text3.Text = Cadena
Text3.SelStart = Len(Text3.Text)
If Text3.Text <> "" Then
   Dim i As Byte
   Dim TSQL As String
   Dim TB As New ADODB.Recordset
   Dim Cad As String
   
  
   

   TSQL = "SELECT * FROM Usuarios WHERE Nombre LIKE  '" & Text3.Text & "%' Order by Nombre"


   TB.Open TSQL, Cn
   FG.Rows = cantrecord(TSQL) + 1
   
   If Not TB.EOF Then
      
        
        
        i = 1
        
        While Not TB.EOF
        
          FG.TextMatrix(i, 0) = TB("Clave")
          FG.TextMatrix(i, 1) = TB("Usuario")
          FG.TextMatrix(i, 2) = TB("Nombre")
          TB.MoveNext
          i = i + 1
        Wend
    Else
     Res = MsgBox("No Hay Datos", 64, "Información")
     Text3.Text = Mid(Text3.Text, 1, Len(Text3.Text) - 1)
     Text3.SelStart = Len(Text3.Text)
    End If

End If


End Sub
