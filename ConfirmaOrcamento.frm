VERSION 5.00
Begin VB.Form ConfirmaOrcamento 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3105
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "ConfirmaOrcamento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Fechar"
      Default         =   -1  'True
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "ConfirmaOrcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private a As Integer

Private Sub Command1_Click()
On Error Resume Next
Unload DadosOrcamento
Unload Me
GlFormA.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
Select Case GlFormA.Name
  Case Is = "Orcamento"
     Label1.Caption = orcamento.Natureza.Text & " Gerado com o Número:"
     Label2.Caption = orcamento.Documento.Text
  Case Is = "FrmProposta"
     Label1.Caption = "Proposta Comercial Gerado com o Número:"
     Label2.Caption = FrmProposta.Txt(0).Text
End Select
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub
