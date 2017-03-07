VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H008080FF&
   Caption         =   "Form1"
   ClientHeight    =   9585
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080FF80&
      Caption         =   "Adios!!!!"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   3615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FF8080&
      Caption         =   "Cassette."
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   3615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF8080&
      Caption         =   "Actor."
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   3615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF8080&
      Caption         =   "Pelicula"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Tipo pelicula."
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Alquiler."
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "Clientes,"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   3405
      Left            =   7560
      Picture         =   "Menu.frx":0000
      Top             =   2760
      Width           =   6855
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Tipo_pelicula.Show (abrir)

Menu.Hide
End Sub

Private Sub Command2_Click()
Alquiller.Show (abrir)

Menu.Hide
End Sub

Private Sub Command3_Click()
Clientes.Show (abrir)

Menu.Hide
End Sub

Private Sub Command4_Click()
Pelicula.Show (abrir)

Menu.Hide
End Sub

Private Sub Command5_Click()
Actor.Show (abrir)

Menu.Hide
End Sub

Private Sub Command6_Click()
Cassette.Show (abrir)

Menu.Hide
End Sub

Private Sub Command7_Click()

If MsgBox("¿ENCERIO, QUIERES SALIR WEY?", _
vbQuestion + vbYesNo, "SALIR") = vbYes Then
End
Else
Cancel = Value
End If





End Sub
