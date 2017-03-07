VERSION 5.00
Begin VB.Form Alquiller 
   BackColor       =   &H00C00000&
   Caption         =   "Form4"
   ClientHeight    =   14160
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   18960
   LinkTopic       =   "Form4"
   ScaleHeight     =   14160
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Menu."
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   16800
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Data bele 
      Caption         =   "                            Tipo de pelicula"
      Connect         =   "Access"
      DatabaseName    =   "E:\100% no fake\Disks.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   14400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ALQUILERES"
      Top             =   8040
      Width           =   4215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Siguiente registro."
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10680
      TabIndex        =   20
      Top             =   9120
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ultimo registro."
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      TabIndex        =   19
      Top             =   9120
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Anterior registro."
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   18
      Top             =   9120
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eliminar."
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10680
      TabIndex        =   17
      Top             =   7920
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar registro."
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      TabIndex        =   16
      Top             =   7920
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo registro."
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   15
      Top             =   7920
      Width           =   3255
   End
   Begin VB.TextBox Text8 
      DataField       =   "cantidad"
      DataSource      =   "bele"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   14
      Top             =   7080
      Width           =   8655
   End
   Begin VB.TextBox Text7 
      DataField       =   "valor de alquiler"
      DataSource      =   "bele"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   12
      Top             =   6120
      Width           =   8655
   End
   Begin VB.TextBox Text6 
      DataField       =   "fecha devolucion"
      DataSource      =   "bele"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   10
      Top             =   5160
      Width           =   8655
   End
   Begin VB.TextBox Text5 
      DataField       =   "fecha de alquiler"
      DataSource      =   "bele"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   8
      Top             =   4200
      Width           =   8655
   End
   Begin VB.TextBox Text4 
      DataField       =   "cod_clientes"
      DataSource      =   "bele"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   6
      Top             =   3240
      Width           =   8655
   End
   Begin VB.TextBox Text3 
      DataField       =   "Codigo de cassette"
      DataSource      =   "bele"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   4
      Top             =   2280
      Width           =   8655
   End
   Begin VB.TextBox Text2 
      DataField       =   "codigo"
      DataSource      =   "bele"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   3
      Top             =   1320
      Width           =   8655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      TabIndex        =   0
      Text            =   "        Alquiler."
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label7 
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   13
      Top             =   7080
      Width           =   5175
   End
   Begin VB.Label Label6 
      Caption         =   "Valor_alquiler."
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   11
      Top             =   6120
      Width           =   5175
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha_devolucion."
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   9
      Top             =   5160
      Width           =   5175
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha_alquiler"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   7
      Top             =   4200
      Width           =   5175
   End
   Begin VB.Label Label3 
      Caption         =   "Cod_cliente."
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   5
      Top             =   3240
      Width           =   5175
   End
   Begin VB.Label l 
      Caption         =   "Cod_cassette."
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   2
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo."
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   5175
   End
End
Attribute VB_Name = "Alquiller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
bele.Recordset.AddNew
End Sub

Private Sub Command2_Click()
bele.Recordset.Update
End Sub

Private Sub Command3_Click()
bele.Recordset.Delete
End Sub

Private Sub Command4_Click()
bele.Recordset.MovePrevious
If bele.Recordset.BOF Then
bele.Recordset.MoveNext
End If


End Sub

Private Sub Command5_Click()
bele.Recordset.MoveLast

End Sub

Private Sub Command6_Click()
bele.Recordset.MoveNext
If bele.Recordset.EOF Then
bele.Recordset.MovePrevious
End If

End Sub



Private Sub Command7_Click()
Menu.Show (abrir)

Alquiller.Hide
End Sub
