VERSION 5.00
Begin VB.MDIForm menusistema 
   BackColor       =   &H8000000C&
   ClientHeight    =   6660
   ClientLeft      =   3030
   ClientTop       =   1980
   ClientWidth     =   9360
   LinkTopic       =   "MDIForm1"
   Picture         =   "menusistema.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.Menu menubasicas 
      Caption         =   "Basicas"
      Begin VB.Menu menuprestamo 
         Caption         =   "Prestamo"
         Begin VB.Menu menuregistrarprestamo 
            Caption         =   "Registrar Tipo de Prestamo"
         End
      End
      Begin VB.Menu menusocio 
         Caption         =   "Registrar Socio"
      End
      Begin VB.Menu menubanco 
         Caption         =   "Banco"
         Begin VB.Menu menuregistrarbanco 
            Caption         =   "Registrar Banco"
         End
         Begin VB.Menu menuregistrarcuenta 
            Caption         =   "Registrar Cuenta Bancaria"
         End
         Begin VB.Menu menuregistrarcheque 
            Caption         =   "Registrar Chequera"
         End
         Begin VB.Menu menuanular 
            Caption         =   "Anular Cheque"
         End
      End
   End
   Begin VB.Menu menumovimientos 
      Caption         =   "Movimientos"
      Begin VB.Menu menumovprestamo 
         Caption         =   "Prestamo"
         Begin VB.Menu menuregistrarpago 
            Caption         =   "Registar Pago"
         End
         Begin VB.Menu menuasignar 
            Caption         =   "Asignar Prestamo"
         End
         Begin VB.Menu menusimularprestamo 
            Caption         =   "Simular Prestamo"
         End
      End
      Begin VB.Menu menuaporte 
         Caption         =   "Registrar Aporte"
      End
      Begin VB.Menu menuretiro 
         Caption         =   "Registrar Retiro"
      End
      Begin VB.Menu menugenerarestado 
         Caption         =   "Generar Estado de Cuenta"
      End
   End
   Begin VB.Menu menuseguridad 
      Caption         =   "Seguridad"
      Begin VB.Menu menugenerarusuario 
         Caption         =   "Generar Usuario"
      End
   End
   Begin VB.Menu menusalir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "menusistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub menuaportepatrono_Click()
RegistrarAporte.Show
End Sub

Private Sub menuaportesocio_Click()
RegistrarAporte.Show
End Sub

Private Sub MDIForm_Load()
If usuario = "1" Then
menuprestamo.Enabled = False
menuprestamo.Visible = False
menuregistrarprestamo.Enabled = False
menuregistrarprestamo.Visible = False
menusocio.Enabled = False
End If

End Sub

Private Sub menuanular_Click()
anularcheque.Show
End Sub

Private Sub menuaporte_Click()
RegistrarAporte.Show
End Sub

Private Sub menuasignar_Click()
AsignarPrestamo.Show
End Sub



Private Sub menugenerarestado_Click()
EstadodeCuenta.Show
End Sub

Private Sub menugenerarusuario_Click()
Generarusuario.Show
End Sub

Private Sub menuregistrarbanco_Click()
RegistrarBancos.Show
End Sub

Private Sub menuregistrarcheque_Click()
registrarcheque.Show
End Sub

Private Sub menuregistrarcuenta_Click()
registrarcuentabancaria.Show
End Sub

Private Sub menuregistrarpago_Click()
Registrarpago.Show
End Sub

Private Sub menuregistrarprestamo_Click()
GuardarTipoPrestamo.Show
End Sub

Private Sub menuregistrarsocio_Click()

End Sub



Private Sub menuretiro_Click()
RegistrarRetiro.Show
End Sub

Private Sub menusalir_Click()
End
End Sub

Private Sub menusimularprestamo_Click()
simularprestamo.Show
End Sub

Private Sub menusocio_Click()
RegistrarSocio.Show
End Sub
