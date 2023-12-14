VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmDataSisema 
   BackColor       =   &H00A7A3FE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Atual"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3480
   Icon            =   "Frmdatasisema.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3267.531
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox txt 
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13,5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99/99/9999"
      PromptChar      =   " "
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK F2"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   630
   End
End
Attribute VB_Name = "frmDataSisema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private a As Integer



Private Sub CmdOk_Click()
On Error Resume Next
Dim RsEmpresa As Recordset
Dim LcEmpresa As String
If IsDate(Txt.Text) Then
   GlDataSistema = Format(Txt, "dd/mm/yyyy")
   Date() = CDate(Format(Txt.Text, "dd/mm/yyyy"))
   FrmPrincipal.Show
   Unload frmDataSisema
Else
   MsgBox "A data Digitada não é Válida...", 64, "Data Incorreta."
   Txt.SetFocus
End If
End Sub

Private Sub Form_Load()
On Error Resume Next


If GlDataSistema = "00:00:00" Then
   Txt = Format(Date, "dd/mm/yyyy")
Else
  Txt = Format(GlDataSistema, "dd/mm/yyyy")
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
End Sub

