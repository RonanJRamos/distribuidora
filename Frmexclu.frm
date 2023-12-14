VERSION 5.00
Begin VB.Form FrmExcluiItem 
   Caption         =   "Item a Excluir"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5580
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdFechar 
      Caption         =   "Fechar"
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtro"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
      Begin VB.OptionButton Todos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Individual 
         Caption         =   "Individual"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Produto a Excluir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2520
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "FrmExcluiItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFechar_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
