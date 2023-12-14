VERSION 5.00
Begin VB.Form ErrosEncontrados 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Erros de Processamento."
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdFechar 
      Caption         =   "Fechar"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   1815
   End
   Begin VB.ListBox Erro 
      Height          =   5520
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Os seguintes Erros foram encontrados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "ErrosEncontrados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFechar_Click()
Unload Me
End Sub
