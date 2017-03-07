VERSION 5.00
Begin VB.Form Salir 
   BackColor       =   &H8000000C&
   Caption         =   "Form7"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6285
   LinkTopic       =   "Form7"
   ScaleHeight     =   3510
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "¿seguro que quieres salir?"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   1815
   End
End
Attribute VB_Name = "Salir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Salir.Hide
Menu.Hide
Actor.Hide
Alquiller.Hide
Cassette.Hide
Clientes.Hide
Pelicula.Hide
Tipo_pelicula.Hide


End Sub

Private Sub Command2_Click()
Menu.Show (abrir)

Salir.Hide

End Sub
