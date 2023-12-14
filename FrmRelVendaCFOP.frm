VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmRelVendaCFOP 
   BackColor       =   &H00D8C5B6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório Vernda CFOP"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Produto 
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   1800
      Width           =   7095
   End
   Begin VB.TextBox Cliente 
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   1080
      Width           =   7095
   End
   Begin VB.TextBox CFOP 
      Height          =   375
      Left            =   3840
      TabIndex        =   23
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E6E4D2&
      Caption         =   "Saída"
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   1815
      Begin VB.OptionButton Video 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Video"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton impressora 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirmar F3"
      Height          =   495
      Left            =   6000
      TabIndex        =   14
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   6000
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E6E4D2&
      Caption         =   "Status"
      Height          =   1335
      Left            =   2040
      TabIndex        =   9
      Top             =   2280
      Width           =   1815
      Begin VB.OptionButton OptCanceladas 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Canceladas"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   630
         Width           =   1335
      End
      Begin VB.OptionButton OptEmitidas 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Emitida"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptTodas 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Todas"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E6E4D2&
      Caption         =   "Situação"
      Height          =   1335
      Left            =   3960
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Todas"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton OptTrans 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Transmitida"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Não Transmitida"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   630
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E6E4D2&
      Caption         =   "Notas de"
      Height          =   1335
      Left            =   5880
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
      Begin VB.OptionButton OptSaida 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Saida"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptEntrada 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Entrada"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Todas"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.ListBox FormaPag 
      Height          =   960
      ItemData        =   "FrmRelVendaCFOP.frx":0000
      Left            =   120
      List            =   "FrmRelVendaCFOP.frx":000D
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   4080
      Width           =   2175
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   360
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
   Begin MSMask.MaskEdBox Dataf 
      Height          =   375
      Left            =   2160
      TabIndex        =   19
      Top             =   360
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
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
      Left            =   120
      TabIndex        =   28
      Top             =   1560
      Width           =   825
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
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
      Left            =   120
      TabIndex        =   26
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CFOP"
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
      Left            =   3840
      TabIndex        =   24
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Final"
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
      Left            =   2160
      TabIndex        =   22
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial"
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
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   1185
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forma de Pagamento"
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
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   2250
   End
End
Attribute VB_Name = "FrmRelVendaCFOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Rs        As ADODB.Recordset
Private Rel       As New CrysVendaCFOP
Private Sub Command1_Click()
On Error Resume Next
Dim StrSql          As String
If Not IsDate(Datai.Text) And Not IsDate(Dataf.Text) Then
    MsgBox "Informe o Periodo para Visualizar o Relatorio.", 64, "Aviso"
    Datai.SetFocus
    Exit Sub
End If
LcCap = Me.Caption
Me.Caption = "Aguarde, processando os dados..."
Screen.MousePointer = 11
Screen.MousePointer = 0
Me.Caption = LcCap
Screen.MousePointer = vbHourglass
StrSql = "SELECT alid050.NUMNF, alid050.DTEMIS, alid050.NATUREZA, alid050.CLIENTE, alid050.CondPag, alid050.status, alid050.NomeCliente,"
StrSql = StrSql & " alid050.finalidadeEmissao, alid050.FomaEmissao, alid052.ITEM, alid052.QTDE, alid052.VALUNIT, alid052.descricao,"
StrSql = StrSql & "alid052.codProd, alid052.CFOP, alid052.CST, alid052.ValorIcms as icms, alid052.desconto, alid052.frete, alid052.Seguro,"
StrSql = StrSql & "alid052.despAcessorias, alid052.CEST FROM alid050 INNER JOIN alid052 ON alid050.NUMNF = alid052.NUMNF"
StrSql = StrSql & " WHERE (((alid050.DTEMIS) Between '" & Format(CDate(Datai.Text), "yyyy-mm-dd") & "' And '" & Format(CDate(Dataf.Text), "yyyy-mm-dd") & "'))"
If Len(CFOP.Text) > 0 Then
   StrSql = StrSql & " and alid052.CFOP='" & CFOP.Text & "'"
End If
If Len(Cliente.Text) > 0 Then
   StrSql = StrSql & " and alid050.NomeCliente like'%" & Cliente.Text & "%'"
End If
If Len(Produto.Text) > 0 Then
   StrSql = StrSql & " and alid052.descricao like'%" & Produto.Text & "%'"
End If

If Option1.Value = False Then
    If OptTrans.Value = True Then
       StrSql = StrSql & " and (Status like 'Autorizado%')"
    End If
    If OptTrans.Value = False Then
       StrSql = StrSql & " and alid050.transmitida=0"
    End If
End If
If OptTodas.Value = False Then
    If OptCanceladas.Value = True Then
       StrSql = StrSql & " And alid050.Status='CANCELADA'"
    End If
    If OptCanceladas.Value = False Then
       StrSql = StrSql & " And alid050.Status<>'CANCELADA'"
    End If
End If
If OptSaida.Value = True Then
   StrSql = StrSql & " And alid050.TipoOperacao='1 - Saida'"
End If
If OptEntrada.Value = True Then
   StrSql = StrSql & " And alid050.TipoOperacao='0 - Entrada'"
End If

For a = 0 To FormaPag.ListCount - 1
    If FormaPag.Selected(a) Then
     If Len(LcIn) > 0 Then LcIn = LcIn & ","
       LcIn = LcIn & "'" & FormaPag.List(a) & "'"
    End If
    
Next
If Len(LcIn) > 0 Then
   LcIn = " alid050.condpag in(" & LcIn & ")"
   StrSql = StrSql & " And " & LcIn
End If
Debug.Print StrSql
Set Rs = AbreRecordset(StrSql, True)
Load Relatorios
With Relatorios
     Rel.DiscardSavedData
     Rel.Database.SetDataSource Rs
     Rel.Subreport1.OpenSubreport.DiscardSavedData
     Rel.Subreport1.OpenSubreport.Database = Rs
     Rel.Subreport1.OpenSubreport.Database.SetDataSource Rs
     .CRViewer1.ReportSource = Rel
     setaformula
      .CRViewer1.ViewReport
End With
Relatorios.Show
Screen.MousePointer = vbDefault
End Sub
Sub setaformula()
Dim a As Integer
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset
Dim RsOpcao As Recordset
Dim LcValor As Double
Dim LcEmpresa, LcEndereco, LcFone, LcCelular, Lccelular1, Lcemail, LcVer, LcCap, LcVer1 As String
Dim lctitulo As String
Dim StrSql As String
Dim bb     As Database

Set db = OpenDatabase(GLBase)
Set RsEmpresa = db.OpenRecordset("Select * from EMPRESA")

If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao & ""
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone & ""
   Lcemail = RsEmpresa!Email & ""
   
End If
Set RsEmpresa = Nothing
lctitulo = "Relatorio de Vendas por CFOP"
With Rel
'Exit Sub
For a = 1 To .FormulaFields.Count
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Fone") Then .FormulaFields(a).Text = "totext('" & LcFone & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("EMPRESA") Then .FormulaFields(a).Text = "totext('" & LcEmpresa & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("ENDERECOEMPRESA") Then .FormulaFields(a).Text = "totext('" & LcEndereco & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = "TIPO" Then .FormulaFields(a).Text = "totext('" & Tipo & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("email") Then .FormulaFields(a).Text = "totext('" & Lcemail & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Titulo") Then
           .FormulaFields(a).Text = "totext('" & lctitulo & "')"
        End If
    Next
End With
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
