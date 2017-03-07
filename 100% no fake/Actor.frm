VERSION 5.00
Begin VB.Form Actor 
   BackColor       =   &H00FF8080&
   Caption         =   "Form3"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   18960
   LinkTopic       =   "Form3"
   ScaleHeight     =   10950
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
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6120
      Width           =   1815
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
      TabIndex        =   12
      Top             =   6000
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
      TabIndex        =   11
      Top             =   6000
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
      TabIndex        =   10
      Top             =   6000
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
      TabIndex        =   9
      Top             =   7200
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
      TabIndex        =   8
      Top             =   7200
      Width           =   3255
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
      TabIndex        =   7
      Top             =   7200
      Width           =   3255
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
      Height          =   1500
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "actor"
      Top             =   8640
      Width           =   12135
   End
   Begin VB.TextBox Text4 
      DataField       =   "fecha de nacimiento"
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
      Height          =   855
      Left            =   7320
      TabIndex        =   6
      Top             =   4800
      Width           =   8655
   End
   Begin VB.TextBox Text3 
      DataField       =   "nombre"
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
      Height          =   855
      Left            =   7320
      TabIndex        =   4
      Top             =   3360
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
      Height          =   855
      Left            =   7320
      TabIndex        =   3
      Top             =   2040
      Width           =   8655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFF80&
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
      Left            =   5040
      TabIndex        =   0
      Text            =   "        ACTOR"
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha de nacimiento."
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   5
      Top             =   4800
      Width           =   5895
   End
   Begin VB.Label Label2 
      Caption         =   "   Nombre."
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   2
      Top             =   3360
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "    Codigo."
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   1
      Top             =   2040
      Width           =   5895
   End
End
Attribute VB_Name = "Actor"
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

Actor.Hide
End Sub
