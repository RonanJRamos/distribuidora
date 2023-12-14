VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form DadosEntradaNota 
   BackColor       =   &H00B3E9FD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dados Finais da Nota Fiscal"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox quantidade 
      Height          =   375
      Left            =   5640
      TabIndex        =   20
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   3600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   1440
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
      Left            =   7800
      TabIndex        =   15
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Imprime 
      Caption         =   "&Salvar F2"
      Height          =   375
      Left            =   7800
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox TipoMonetario 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   3255
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   2
      Top             =   1440
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
      Left            =   5640
      TabIndex        =   3
      Top             =   1440
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
      Index           =   3
      Left            =   1800
      TabIndex        =   4
      Top             =   1920
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
      Index           =   4
      Left            =   3720
      TabIndex        =   5
      Top             =   1920
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
      Index           =   5
      Left            =   5640
      TabIndex        =   6
      Top             =   1920
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
      Index           =   6
      Left            =   1800
      TabIndex        =   7
      Top             =   2400
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
      Index           =   7
      Left            =   3720
      TabIndex        =   8
      Top             =   2400
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
      Index           =   8
      Left            =   5640
      TabIndex        =   9
      Top             =   2400
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
      Index           =   9
      Left            =   1800
      TabIndex        =   10
      Top             =   2880
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
      Index           =   10
      Left            =   3720
      TabIndex        =   11
      Top             =   2880
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
      Index           =   11
      Left            =   5640
      TabIndex        =   12
      Top             =   2880
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
      Caption         =   "Parcelas"
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
      Left            =   4320
      TabIndex        =   21
      Top             =   3720
      Width           =   945
   End
   Begin VB.Line Line3 
      X1              =   -120
      X2              =   7680
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      X1              =   7680
      X2              =   9240
      Y1              =   -120
      Y2              =   -120
   End
   Begin VB.Line Line1 
      X1              =   7680
      X2              =   7680
      Y1              =   -120
      Y2              =   4200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
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
      TabIndex        =   19
      Top             =   3720
      Width           =   570
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
      TabIndex        =   18
      Top             =   1560
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
      Left            =   0
      TabIndex        =   17
      Top             =   840
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fechamento Nota de Entrada"
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
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   4020
   End
End
Attribute VB_Name = "DadosEntradaNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcNatureza As String
Private LcNota, LcBoleta, LcEspaco, LcLinha, LcEspC As String
Private LcSalto, LcQuant, a  As Integer

Private Sub Command2_Click()
On Error Resume Next
Unload Me

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Form_Load()
Dim LcVer, a As Integer
Valor.Text = FrmEntradaProduto.Txt(16).Text
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
HabilitaPag
Select Case FrmEntradaProduto.Natureza.Text
    Case Is = "A VISTA"
         LcVer = False
    Case Is = "A PRAZO"
         LcVer = True
End Select
For a = 1 To 11
    Vencimento(a).Visible = LcVer
Next
quantidade.Visible = LcVer
Label1(0).Visible = LcVer
'Label1(13).Visible = LcVer
'Label1(14).Visible = LcVer
End Sub
Function HabilitaPag()
Dim Exibe, ExibeMonetario As Integer


Select Case FrmEntradaProduto.Natureza.Text
Case Is = "A VISTA"
     LcNatureza = "VENDAS A VISTA"
     Exibe = False
     ExibeMonetario = True
   
Case Is = "A PRAZO"
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
 CarregaTipoMonetario
'TipoMonetario.Visible = ExibeMonetario
Label1(10).Visible = ExibeMonetario

End Function
Function CarregaTipoMonetario()
Dim RsMoney As Recordset
TipoMonetario.Clear
AbreBase
Set RsMoney = Dbbase.OpenRecordset("Select * from alid008 where VENDA='S' order by XTPMONET", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Do Until RsMoney.EOF
   TipoMonetario.AddItem RsMoney("XTPMONET")
   RsMoney.MoveNext
Loop
RsMoney.Close
Dbbase.Close
Set RsMoney = Nothing
Set Dbbase = Nothing


End Function


Function GeraValor() As Currency
Dim LcValor As Currency
If Vencimento(11).Text = "  /  /  " Then
  If Vencimento(10).Text = "  /  /  " Then
     If Vencimento(9).Text = "  /  /  " Then
       If Vencimento(8).Text = "  /  /  " Then
            If Vencimento(7).Text = "  /  /  " Then
             If Vencimento(6).Text = "  /  /  " Then
                   If Vencimento(5).Text = "  /  /  " Then
                      If Vencimento(4).Text = "  /  /  " Then
                         If Vencimento(3).Text = "  /  /  " Then
                            If Vencimento(2).Text = "  /  /  " Then
                               If Vencimento(1).Text = "  /  /  " Then
                                  If Vencimento(0).Text = "  /  /  " Then
                                  Else
                                     Valor.Text = CCur(FrmEntradaProduto.Txt(16).Text)
                                     LcQuant = 1
                                  End If
                              Else
                                 Valor.Text = CCur(FrmEntradaProduto.Txt(16).Text) / 2
                                 LcQuant = 2
                              End If
                          Else
                             Valor.Text = CCur(FrmEntradaProduto.Txt(16).Text) / 3
                             LcQuant = 3
                         End If
                       Else
                          Valor.Text = CCur(FrmEntradaProduto.Txt(16).Text) / 4
                          LcQuant = 4
                      End If
                    Else
                       Valor.Text = CCur(FrmEntradaProduto.Txt(16).Text) / 5
                       LcQuant = 5
                    End If
                   Else
                     Valor.Text = CCur(FrmEntradaProduto.Txt(16).Text) / 6
                     LcQuant = 6
                   End If
              Else
               Valor.Text = CCur(FrmEntradaProduto.Txt(16).Text) / 7
               LcQuant = 7
              End If
            Else
               Valor.Text = CCur(FrmEntradaProduto.Txt(16).Text) / 8
               LcQuant = 8
            End If
         Else
            Valor.Text = CCur(FrmEntradaProduto.Txt(16).Text) / 9
            LcQuant = 9
         End If
       Else
            Valor.Text = CCur(FrmEntradaProduto.Txt(16).Text) / 10
            LcQuant = 10
         End If
        Else
            Valor.Text = CCur(FrmEntradaProduto.Txt(16).Text) / 11
            LcQuant = 11
        End If
      Else
         Valor.Text = CCur(FrmEntradaProduto.Txt(16).Text) / 12
         LcQuant = 12
     End If
quantidade.Text = LcQuant
End Function

Private Sub Imprime_Click()
'FrmSaidaProduto.ImprimeNota
If Len(TipoMonetario.Text) = 0 Then
    MsgBox "Informe o tipo monetario para o Lancamento.", 64, "Aviso"
Else
   If Not IsDate(Vencimento(0).Text) Then
       MsgBox "Informe pelo menos um vencimento para o lançamento.", 64, "Aviso"
    Else
        FrmEntradaProduto.processanota
        Unload Me
    End If
    
End If

End Sub

Private Sub Imprime_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Quantidade_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub TipoMonetario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub valor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{S}"
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
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Vencimento_LostFocus(Index As Integer)
If Vencimento(Index).Text = "  /  /  " Then Exit Sub
If Not IsDate(Vencimento(Index).Text) Then
   MsgBox "O Valor digitado deve Ser uma Data...", 64, "Aviso"
   Vencimento(Index).Text = "  /  /  "
   Vencimento(Index).SetFocus
End If
End Sub
