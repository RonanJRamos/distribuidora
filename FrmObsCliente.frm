VERSION 5.00
Begin VB.Form FrmObsCliente 
   BackColor       =   &H00F0FFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Observações do Cliente"
   ClientHeight    =   4935
   ClientLeft      =   11355
   ClientTop       =   750
   ClientWidth     =   7050
   Icon            =   "FrmObsCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7050
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "Salvar"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox Comercial 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2640
      Width           =   6735
   End
   Begin VB.TextBox Financeira 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   6735
   End
   Begin VB.Label Codigo 
      BackColor       =   &H00F0FFE3&
      Height          =   255
      Left            =   11160
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F0FFE3&
      Caption         =   "Comercial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F0FFE3&
      Caption         =   "Financeira"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "FrmObsCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSalvar_Click()
On Error GoTo ErroS
Dim LcStrSql As String
Dim LcREgistro As Integer

AbreBase
LcStrSql = "Update ALID001 set ObsFinaceiro='" & Replace(Financeira.Text, "'", "''") & "',ObsComercial='" & Replace(Comercial.Text, "'", "''") & "' where CODIGO='" & Codigo.Caption & "'"
Dbbase.Execute (LcStrSql)
LcREgistro = Dbbase.RecordsAffected
If LcREgistro <> 0 Then
   MsgBox "Registro Salvo com Sucesso!", vbInformation, "Informação"
Else
   MsgBox "Erro salvando o Registro!", vbInformation, "Informação"
End If
Exit Sub
ErroS:
MsgBox "Ocorreu o seguinte erro salvando as informações do cliente: " & err.Description, vbCritical, "Erro Nº:" & err.Number
End Sub

Private Sub Form_Activate()
'Exit Sub
On Error Resume Next
Dim Rs As Recordset
AbreBase
Set Rs = Dbbase.OpenRecordset("Select ObsFinaceiro,ObsComercial from alid001 where CODIGO='" & Codigo.Caption & "'")
If Not Rs.EOF Then
   Financeira.Text = Rs!ObsFinaceiro & ""
   Comercial.Text = Rs!ObsComercial & ""
Else
   Financeira.Text = ""
   Comercial.Text = ""
End If
If Financeira.Locked Then Comercial.SetFocus
Rs.Close
Dbbase.Close
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim Rs As Recordset
AbreBase
Set Rs = Dbbase.OpenRecordset("Select ObsFinaceiro,ObsComercial where CODIGO='" & Codigo.Caption & "'")
If Not Rs.EOF Then
   Financeira.Text = Rs!ObsFinaceiro & ""
   Comercial.Text = Rs!ObsComercial & ""
Else
   Financeira.Text = ""
   Comercial.Text = ""
End If
Rs.Close
Dbbase.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
GlPodeAbrirOBS = False
End Sub
