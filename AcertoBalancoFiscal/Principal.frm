VERSION 5.00
Begin VB.Form Principal 
   BackColor       =   &H00DCDDBF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   2580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSair 
      BackColor       =   &H00DCDDBF&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1230
      Width           =   2295
   End
   Begin VB.CommandButton CmdTranferencia 
      BackColor       =   &H00DCDDBF&
      Caption         =   "Transferência"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   735
      Width           =   2295
   End
   Begin VB.CommandButton CmdBalanco 
      BackColor       =   &H00DCDDBF&
      Caption         =   "Balanço"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBalanco_Click()
BalancoGalpao.Show , Me
End Sub

Private Sub CmdSair_Click()
End
End Sub

Private Sub CmdTranferencia_Click()
Transferencia.Show , Me
End Sub
