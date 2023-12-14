VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form DadosTransp 
   BackColor       =   &H00E6E4D2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dados Finais da Nota Fiscal"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   11
      Left            =   1920
      TabIndex        =   13
      Top             =   4440
      Width           =   5655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   10
      Left            =   1920
      TabIndex        =   12
      Top             =   4200
      Width           =   5655
   End
   Begin VB.TextBox VlTotalSdesconto 
      Height          =   375
      Left            =   7080
      TabIndex        =   21
      Top             =   6120
      Width           =   1815
   End
   Begin VB.TextBox DescontoEst 
      Height          =   375
      Left            =   7080
      TabIndex        =   22
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox Dias 
      Height          =   375
      Index           =   2
      Left            =   4830
      TabIndex        =   17
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox Dias 
      Height          =   375
      Index           =   1
      Left            =   3375
      TabIndex        =   16
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox Dias 
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   15
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox quantidade 
      Height          =   375
      Left            =   1920
      TabIndex        =   39
      Top             =   6840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   375
      Left            =   1920
      TabIndex        =   23
      Top             =   6480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   18
      Top             =   6000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   375
      Left            =   7920
      TabIndex        =   25
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Imprime 
      Caption         =   "&Imprimir F2"
      Height          =   375
      Left            =   7920
      TabIndex        =   24
      Top             =   3000
      Width           =   1335
   End
   Begin VB.ComboBox TipoMonetario 
      Height          =   315
      Left            =   1920
      TabIndex        =   14
      Text            =   "BOLETO"
      Top             =   5040
      Width           =   3255
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   9
      Left            =   1920
      MaxLength       =   60
      TabIndex        =   11
      Top             =   3930
      Width           =   5655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   8
      Left            =   1920
      MaxLength       =   60
      TabIndex        =   10
      Top             =   3645
      Width           =   5655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   7
      Left            =   1920
      MaxLength       =   60
      TabIndex        =   9
      Top             =   3360
      Width           =   5655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   6
      Left            =   1920
      MaxLength       =   12
      TabIndex        =   8
      Top             =   2400
      Width           =   5655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   5
      Left            =   8280
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   4
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   6
      Top             =   2040
      Width           =   5655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   3
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1560
      Width           =   7095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   6120
      MaxLength       =   18
      TabIndex        =   4
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin MSMask.MaskEdBox Placa 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "AAA-9999"
      Mask            =   "AAA-9999"
      PromptChar      =   " "
   End
   Begin VB.ComboBox Tipo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "DadosTransp.frx":0000
      Left            =   7440
      List            =   "DadosTransp.frx":000A
      TabIndex        =   1
      Text            =   "1- CIF"
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   1
      Left            =   3375
      TabIndex        =   19
      Top             =   6000
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   2
      Left            =   4830
      TabIndex        =   20
      Top             =   6000
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ordem Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   15
      Left            =   120
      TabIndex        =   45
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End. Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   120
      TabIndex        =   44
      Top             =   4200
      Width           =   1350
   End
   Begin VB.Line Line4 
      X1              =   6480
      X2              =   6480
      Y1              =   5400
      Y2              =   7440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vl Total  Sem Desc."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6960
      TabIndex        =   43
      Top             =   5880
      Width           =   2070
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vl do Icms (do Desc.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6840
      TabIndex        =   42
      Top             =   6600
      Width           =   2235
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto P/ o Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   6840
      TabIndex        =   41
      Top             =   5520
      Width           =   2280
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   120
      TabIndex        =   40
      Top             =   5640
      Width           =   495
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   7800
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line2 
      X1              =   7800
      X2              =   9360
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      X1              =   7800
      X2              =   7800
      Y1              =   2760
      Y2              =   5400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   14
      Left            =   120
      TabIndex        =   38
      Top             =   6480
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   120
      TabIndex        =   37
      Top             =   6120
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Monetário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   120
      TabIndex        =   36
      Top             =   5040
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Informações Complementares"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   9
      Left            =   2640
      TabIndex        =   35
      Top             =   2880
      Width           =   4185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inscr. Est."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   120
      TabIndex        =   34
      Top             =   2400
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   7800
      TabIndex        =   33
      Top             =   2040
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Município"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   120
      TabIndex        =   32
      Top             =   2040
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   31
      Top             =   1560
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C..G.C./C.P.F.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   4560
      TabIndex        =   30
      Top             =   960
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   3240
      TabIndex        =   29
      Top             =   960
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Placa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   28
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   6840
      TabIndex        =   27
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transportadora "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "DadosTransp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcNatureza As String
Private LcNota, LcBoleta, LcEspaco, LcLinha, LcEspC As String
Private LcSalto, LcQuant As Integer, a As Integer

Private Sub Command2_Click()
On Error Resume Next
Me.Visible = False
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub


Private Sub Dias_Change(Index As Integer)
On Error Resume Next
If IsNumeric(Dias(Index).Text) Then
   If Dias(Index).Text = "0" Then
      Vencimento(Index).Text = "  /  /  "
   Else
      Vencimento(Index).Text = Format(Date + (CInt(Dias(Index).Text)), "dd/mm/yy")
   End If
End If
LcDivide = 1
If Vencimento(0).Text <> "  /  /  " Then
   If Vencimento(1).Text <> "  /  /  " Then
      If Vencimento(2).Text <> "  /  /  " Then
          LcDivide = 3
      Else
          LcDivide = 2
      End If
   Else
      LcDivide = 1
   End If
End If
valor.Text = AcertaNumero(CStr(CCur(AcertaNumero(FrmSaidaProduto.Txt(16).Text, 2)) / LcDivide), 2)

End Sub

Private Sub Dias_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Dias_LostFocus(Index As Integer)
If Len(Dias(Index).Text) = 0 Then Exit Sub
If Not IsNumeric(Dias(Index).Text) Then
   MsgBox "Digite um Valor Numérico...", 64, "Aviso"
   Dias(Index).SetFocus
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
If FrmSaidaProduto.Natureza.Text = "ORG PUBL. EST." Then
   DescontoEst.Visible = True
   Label2.Visible = True
End If
Dim LcVer As Integer
Dim LcNat As String
On Error Resume Next
Dim a As Integer
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
HabilitaPag

Select Case GlFormA.Natureza.Text
    Case Is = "VENDAS A VISTA"
         LcVer = False
    Case Is = "VENDAS A PRAZO"
         LcVer = True
End Select
For a = 1 To 2
    Vencimento(a).Visible = LcVer
Next
CarregaTipoMonetario
BuscaDados
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub
Function BuscaDados()
On Error Resume Next
Dim RsOrc As Recordset
AbreBase
'verifica dados da nota no cadastro de cliente
If GlFormA.Name = "FrmSaidaProduto" Then
   LcSql2 = "Select DadosNota FROM Alid001 Where Codigo='" & FrmSaidaProduto.Txt(8).Text & "'"
Else
   LcSql2 = "Select DadosNota FROM Alid001 Where Codigo='" & FrmProposta.Txt(8).Text & "'"
End If
Set RsOrc = Dbbase.OpenRecordset(LcSql2)
If Not RsOrc.EOF Then
    Txt(7).Text = UCase(RsOrc!dadosnota) & ""
End If
RsOrc.Close
Set RsOrc = Nothing
If GlFormA.Name = "FrmSaidaProduto" Then
   If Len(FrmSaidaProduto.proposta.Text) > 0 Then
      LcSql2 = "Select * from Proposta where NUMNF='" & FrmSaidaProduto.proposta.Text & "'"
   Else
      LcSql2 = "Select * from alid050 where NUMNF='" & FrmSaidaProduto.Txt(0).Text & "'"
   End If
Else
   LcSql2 = "Select * from Proposta where NUMNF='" & FrmProposta.Txt(0).Text & "'"
End If
Set RsOrc = Dbbase.OpenRecordset(LcSql2)
If Not RsOrc.EOF Then
   Txt(0).Text = RsOrc!Transp
   If RsOrc!TIPOTRANS = 1 Then
      Tipo.Text = "1- CIF"
   Else
      Tipo.Text = "2 - FOB"
   End If
   Placa.Text = RsOrc!PLACATRANS
   Txt(1).Text = RsOrc!UFTRANS
   Txt(2).Text = RsOrc!CGCCPFTRAN
   Txt(3).Text = RsOrc!ENDTRANS
   Txt(4).Text = RsOrc!MUNICTRANS
'   txt(6).Text = RsOrc!cidade
'   txt(12).Text = RsOrc!Cep
   Txt(5).Text = RsOrc!UFMUNIC
 '  txt(10).Text = RsOrc!fonetransp & ""
   Txt(7).Text = IIf(Len(Txt(7).Text) = 0, RsOrc!OBS02 & "", Txt(7).Text)
   Txt(8).Text = RsOrc!OBS03 & ""
   Txt(9).Text = RsOrc!OBS04 & ""
   Dias(0).Text = RsOrc!dias1 & ""
   Dias(1).Text = RsOrc!dias2 & ""
   Dias(2).Text = RsOrc!dias3 & ""
   Txt(10).Text = RsOrc!EnderecoEntrega & ""
   Txt(11).Text = RsOrc!OC & ""
'   TipoPag.Text = RsOrc!CondPag & ""
   If Len(RsOrc!formapag) > 0 Then TipoMonetario.Text = RsOrc!formapag & ""
   'txt(11).Text = RsOrc!dias
   If IsNull(RsOrc!Vencimento1) Then Vencimento(0).Text = "  /  /  " Else Vencimento(0).Text = Format(RsOrc!Vencimento1, "dd/mm/yy")
   If IsNull(RsOrc!vencimento2) Then Vencimento(1).Text = "  /  /  " Else Vencimento(1).Text = Format(RsOrc!vencimento2, "dd/mm/yy")
   If IsNull(RsOrc!vencimento3) Then Vencimento(2).Text = "  /  /  " Else Vencimento(2).Text = Format(RsOrc!vencimento3, "dd/mm/yy")
 
End If
If GlFormA.Natureza = "VENDAS A VISTA" Then
   Vencimento(1).Visible = False
   Vencimento(2).Visible = False
Else
   Vencimento(1).Visible = True
   Vencimento(2).Visible = True
End If
GeraValor
RsOrc.Close
Set RsOrc = Nothing

End Function
Function HabilitaPag()
Dim Exibe, ExibeMonetario As Integer

On Error Resume Next
Select Case GlFormA.Natureza.Text
Case Is = "VENDAS A VISTA"
     LcNatureza = "VENDAS A VISTA"
     Exibe = False
     ExibeMonetario = True
   
Case Is = "VENDAS A PRAZO"
     LcNatureza = "VENDAS A PRAZO"
     Exibe = True
     ExibeMonetario = False
   
Case Is = "SR - Simples Remessa"
     LcNatureza = "Simples Remessa"
     ExibeMonetario = False
     Exibe = False
   
Case Is = "ND - Nota Devolucao"
    LcNatureza = "Nota Devolução"
    ExibeMonetario = False
    Exibe = False
  
End Select
If ExibeMonetario Then CarregaTipoMonetario
TipoMonetario.Visible = True
Label1(10).Visible = ExibeMonetario

End Function
Function CarregaTipoMonetario()
Dim RsMoney As Recordset
TipoMonetario.Clear
AbreBase
Set RsMoney = Dbbase.OpenRecordset("Select * from alid008 where VENDA='S' order by XTPMONET")
Do Until RsMoney.EOF
   TipoMonetario.AddItem RsMoney("XTPMONET")
   RsMoney.MoveNext
Loop
RsMoney.Close
Dbbase.Close
Set RsMoney = Nothing
Set Dbbase = Nothing
TipoMonetario.Text = "BOLETO"

End Function


Function GeraValor() As Currency
Dim LcValor As Currency
If Len(GlFormA.Txt(16).Text) = 0 Then GlFormA.Txt(16).Text = 0
If Vencimento(2).Text = "  /  /  " Then
    If Vencimento(1).Text = "  /  /  " Then
        If Vencimento(0).Text = "  /  /  " Then
        Else
          valor.Text = CCur(GlFormA.Txt(16).Text)
          LcQuant = 1
        End If
    Else
       valor.Text = CCur(GlFormA.Txt(16).Text) / 2
       LcQuant = 2
   End If
Else
   valor.Text = CCur(GlFormA.Txt(16).Text) / 3
   LcQuant = 3
End If
Quantidade.Text = LcQuant
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmSaidaProduto.SetFocus
End Sub

Private Sub Imprime_Click()
'===> Checando se os Valores Batem
On Error Resume Next
Imprime.Enabled = False
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando a nota Fiscal."
'Area.BeginTrans
If Vencimento(0).Text = "  /  /  " Then Vencimento(0).Text = Format(Date, "dd/mm/yy")

If GlFormA.Name = "FrmSaidaProduto" Then
     If Len(FrmSaidaProduto.Txt(8).Text) = 0 Then
        MsgBox "O Cliente Selecionado nâo é Valido.", 64, "Aviso"
        Exit Sub
     End If

    If Not IsNumeric(Quantidade.Text) Then Quantidade.Text = 1
    If Not IsNumeric(valor.Text) Then
       MsgBox "O Valor das parcelas deve ser Um Valor Numerico.", 64, "Aviso"
       Exit Sub
    End If
    If Not IsDate(Vencimento(0).Text) Then Vencimento(0).Text = "  /  /  "
    If Not IsDate(Vencimento(1).Text) Then Vencimento(1).Text = "  /  /  "
    If Not IsDate(Vencimento(2).Text) Then Vencimento(2).Text = "  /  /  "
    
    If Len(FrmSaidaProduto.Txt(17).Text) = 0 Then FrmSaidaProduto.Txt(17).Text = 0
    If Not IsNumeric(FrmSaidaProduto.Txt(15).Text) Then
       MsgBox "O valor total dos produtos deve ser Numerico", 64, "Aviso"
       Exit Sub
    End If
    If Not IsNumeric(FrmSaidaProduto.Txt(16).Text) Then
       MsgBox "O valor total da Nota deve ser Numerico", 64, "Aviso"
       Exit Sub
    End If
     FrmSaidaProduto.processanota
     
 'Else
      ' MsgBox "Ocorreu um erro no lançamento da nota, Nota não Lançada.", 64, "Aviso"
       'GoTo desfaz
  ' End If
   Unload Me
   FrmSaidaProduto.Txt(0).SetFocus
Else

    If GlFormA.Name = "FrmSaidaProdutoAlternativo" Then
         If Len(FrmSaidaProdutoAlternativo.Txt(8).Text) = 0 Then
            MsgBox "O Cliente Selecionado nâo é Valido.", 64, "Aviso"
            Exit Sub
         End If
    
        If Not IsNumeric(Quantidade.Text) Then Quantidade.Text = 1
        If Not IsNumeric(valor.Text) Then
           MsgBox "O Valor das parcelas deve ser Um Valor Numerico.", 64, "Aviso"
           Exit Sub
        End If
        If Not IsDate(Vencimento(0).Text) Then Vencimento(0).Text = "  /  /  "
        If Not IsDate(Vencimento(1).Text) Then Vencimento(1).Text = "  /  /  "
        If Not IsDate(Vencimento(2).Text) Then Vencimento(2).Text = "  /  /  "
        
        If Len(FrmSaidaProdutoAlternativo.Txt(17).Text) = 0 Then FrmSaidaProdutoAlternativo.Txt(17).Text = 0
        If Not IsNumeric(FrmSaidaProdutoAlternativo.Txt(15).Text) Then
           MsgBox "O valor total dos produtos deve ser Numerico", 64, "Aviso"
           Exit Sub
        End If
        If Not IsNumeric(FrmSaidaProdutoAlternativo.Txt(16).Text) Then
           MsgBox "O valor total da Nota deve ser Numerico", 64, "Aviso"
           Exit Sub
        End If
         FrmSaidaProdutoAlternativo.processanota
         
     'Else
          ' MsgBox "Ocorreu um erro no lançamento da nota, Nota não Lançada.", 64, "Aviso"
           'GoTo desfaz
      ' End If
       Unload Me
       FrmSaidaProdutoAlternativo.Txt(0).SetFocus
    Else
       
           FrmProposta.SalvaNota
           FrmProposta.Imprime
        
           FrmProposta.limpanota
           Unload Me
           FrmProposta.Txt(0).SetFocus
     End If
End If
'Area.CommitTrans
Me.Caption = LcCap
Imprime.Enabled = True

Exit Sub
Me.Caption = LcCap
Imprime.Enabled = True
desfaz:
'Area.Rollback
'Call FrmSaidaProduto.excluilancamentos
Exit Sub

End Sub

Private Sub Imprime_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Placa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Tipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub TipoMonetario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Txt_LostFocus(Index As Integer)
On Error Resume Next
Txt(Index).Text = UCase(Txt(Index).Text)
End Sub

Private Sub valor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub valor_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub Vencimento_Change(Index As Integer)
GeraValor
End Sub

Private Sub Vencimento_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

