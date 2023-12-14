VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmRelSaidaEstPeriodo 
   BackColor       =   &H00E6E4D2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Saída de Estoque Período"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox Dataf 
      Height          =   375
      Left            =   1680
      TabIndex        =   29
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.ComboBox Comissao 
      Height          =   315
      ItemData        =   "FrmRelSaidaEstPeriodo.frx":0000
      Left            =   3240
      List            =   "FrmRelSaidaEstPeriodo.frx":0002
      TabIndex        =   25
      Top             =   540
      Width           =   3855
   End
   Begin VB.TextBox codigo 
      Height          =   405
      Left            =   5760
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox IncluirTrasnsferencias 
      BackColor       =   &H00E6E4D2&
      Caption         =   "Incluir Transferências"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   23
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CheckBox IncluirDevolucao 
      BackColor       =   &H00E6E4D2&
      Caption         =   "Incluir Devoluções"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   22
      Top             =   2880
      Width           =   2535
   End
   Begin VB.ListBox FormaPag 
      Height          =   960
      ItemData        =   "FrmRelSaidaEstPeriodo.frx":0004
      Left            =   120
      List            =   "FrmRelSaidaEstPeriodo.frx":0011
      Style           =   1  'Checkbox
      TabIndex        =   20
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E6E4D2&
      Caption         =   "Notas de"
      Height          =   1335
      Left            =   5880
      TabIndex        =   16
      Top             =   1080
      Width           =   1335
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Todas"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptEntrada 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Entrada"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton OptSaida 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Saida"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E6E4D2&
      Caption         =   "Situação"
      Height          =   1335
      Left            =   3960
      TabIndex        =   12
      Top             =   1080
      Width           =   1815
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Não Transmitida"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   630
         Width           =   1575
      End
      Begin VB.OptionButton OptTrans 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Transmitida"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Todas"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E6E4D2&
      Caption         =   "Status"
      Height          =   1335
      Left            =   2040
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
      Begin VB.OptionButton OptTodas 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Todas"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptEmitidas 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Emitida"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptCanceladas 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Canceladas"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   630
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   2520
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirmar F3"
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox copias 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Text            =   "1"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E6E4D2&
      Caption         =   "Saída"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
      Begin VB.OptionButton impressora 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Video 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Video"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
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
      Left            =   3240
      TabIndex        =   27
      Top             =   120
      Width           =   1035
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
      Left            =   1680
      TabIndex        =   26
      Top             =   120
      Width           =   1080
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
      TabIndex        =   21
      Top             =   2520
      Width           =   2250
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
      TabIndex        =   7
      Top             =   120
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copias"
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
      Left            =   4920
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "FrmRelSaidaEstPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TipoVend
      codigo As String
      Nome As String
End Type
Private LcTamanho, a As Integer
Private MtVendedor() As TipoVend
Private Sub Cbo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub analitico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub
Function GeraNota()
On Error Resume Next
Dim RsNota As ADODB.Recordset
Dim RsNotaMdb As Recordset
Dim LcSql As String
Dim LcNome As String
Dim LcIn As String
'StrSql = "Update  Alid050 set valorproduto=0,ValorNota=0,DESCONTO=0,baseicms=0,valoricms=0,acrescimo=0,basepis=0,valorpis=0,basecofins=0,valorcofins=0,BaseCalculoIcmsSubst=0,valorIcmsSubst=0,valorIpi=0,outrasdespesas=0,frete=0 ,seguro=0 where (LENGTH(protocolonfe)=0) AND (DTEMIS Between '" & Format(Datai.Text, "yyyy-mm-dd") & "' and '" & Format(Dataf.Text, "yyyy-mm-dd") & "')"
'conexaoAdo.Execute StrSql
'StrSql = "Update Alid050 set valorproduto=0,ValorNota=0,DESCONTO=0,baseicms=0,valoricms=0,acrescimo=0,basepis=0,valorpis=0,basecofins=0,valorcofins=0,BaseCalculoIcmsSubst=0,valorIcmsSubst=0,valorIpi=0,outrasdespesas=0,frete=0 ,seguro=0 where (LENGTH(ProtocoloCancelamento)>0) AND (DTEMIS Between '" & Format(Datai.Text, "yyyy-mm-dd") & "' and '" & Format(Dataf.Text, "yyyy-mm-dd") & "')"
'conexaoAdo.Execute StrSql


'StrSql = "Update Alid050 set valorproduto=0,ValorNota=0,DESCONTO=0,baseicms=0,valoricms=0,acrescimo=0,basepis=0,valorpis=0,basecofins=0,valorcofins=0,BaseCalculoIcmsSubst=0,valorIcmsSubst=0,valorIpi=0,outrasdespesas=0,frete=0 ,seguro=0 where (natureza='TRANSFERENCIA' or natureza='TR') AND (DTEMIS Between '" & Format(Datai.Text, "yyyy-mm-dd") & "' and '" & Format(Dataf.Text, "yyyy-mm-dd") & "') "
'conexaoAdo.Execute StrSql
'StrSql = "Update Alid050 set valorproduto=0,ValorNota=0,DESCONTO=0,baseicms=0,valoricms=0,acrescimo=0,basepis=0,valorpis=0,basecofins=0,valorcofins=0,BaseCalculoIcmsSubst=0,valorIcmsSubst=0,valorIpi=0,outrasdespesas=0,frete=0 ,seguro=0 where (STATUS='INUTILIZADA') AND (DTEMIS Between '" & Format(Datai.Text, "yyyy-mm-dd") & "' and '" & Format(Dataf.Text, "yyyy-mm-dd") & "') "
'conexaoAdo.Execute StrSql
'StrSql = "Update Alid050 set valorproduto=0,ValorNota=0,DESCONTO=0,baseicms=0,valoricms=0,acrescimo=0,basepis=0,valorpis=0,basecofins=0,valorcofins=0,BaseCalculoIcmsSubst=0,valorIcmsSubst=0,valorIpi=0,outrasdespesas=0,frete=0 ,seguro=0 where (STATUS  LIKE '%DENEG%') AND (DTEMIS Between '" & Format(Datai.Text, "yyyy-mm-dd") & "' and '" & Format(Dataf.Text, "yyyy-mm-dd") & "') "
'conexaoAdo.Execute StrSql
LcSql = "Select * from alid050 where DTEMIS Between '" & Format(Datai.Text, "yyyy-mm-dd") & "' And '" & Format(Dataf.Text, "yyyy-mm-dd") & "'"
If Option1.Value = False Then
    If OptTrans.Value = True Then
       LcSql = LcSql & " and (Status like 'Autorizado%')"
    End If
    If OptTrans.Value = False Then
       LcSql = LcSql & " and transmitida=0"
    End If
End If
If OptTodas.Value = False Then
    If OptCanceladas.Value = True Then
       LcSql = LcSql & " And Status='CANCELADA'"
    End If
    If OptCanceladas.Value = False Then
       LcSql = LcSql & " And Status<>'CANCELADA'"
    End If
End If
If OptSaida.Value = True Then
   LcSql = LcSql & " And TipoOperacao like '1%'"
End If
If OptEntrada.Value = True Then
   LcSql = LcSql & " And TipoOperacao like '0%'"
End If
If IncluirDevolucao.Value = 0 Then
   LcSql = LcSql & " and finalidadeEmissao not like '4%'"
End If
If IncluirTrasnsferencias.Value = 0 Then
   LcSql = LcSql & " and natureza not like 'trans%'"
End If
If IsNumeric(codigo.Text) Then
   LcSql = LcSql & "  and (vendedor=" & codigo.Text & ")"
End If
For a = 0 To FormaPag.ListCount - 1
    If FormaPag.Selected(a) Then
     If Len(LcIn) > 0 Then LcIn = LcIn & ","
       LcIn = LcIn & "'" & FormaPag.List(a) & "'"
    End If
Next
If Len(LcIn) > 0 Then
   LcIn = " condpag in(" & LcIn & ")"
   LcSql = LcSql & " And " & LcIn
End If

AbreBase

'Debug.Print LcSql
Set RsNota = AbreRecordsetRel(LcSql, RsNota)
Set RsNotaMdb = Dbbase.OpenRecordset("Select * from alid050", dbOpenDynaset, dbSeeChanges, dbOptimistic)
RsNota.Requery
'===> Apagando Registros antigos
Do Until RsNotaMdb.EOF
    DoEvents
    RsNotaMdb.Delete
    RsNotaMdb.MoveNext
Loop

RsNota.MoveFirst
Do Until RsNota.EOF
    RsNotaMdb.AddNew
     Dim Zera As Boolean
     Zera = False
    For C = 0 To RsNota.Fields.Count - 1
       
        LcNome = RsNota.Fields(C).Name
        '====> Verifica as Situaçoes para zerar
        If InStr(1, UCase(LcNome), UCase("protocolonfe"), vbTextCompare) > 0 Then If Len(RsNota.Fields(C)) = 0 Then Zera = True
        If InStr(1, UCase(LcNome), UCase("ProtocoloCancelamento"), vbTextCompare) > 0 Then If Len(RsNota.Fields(C)) > 0 Then Zera = True
        If InStr(1, UCase(LcNome), UCase("natureza"), vbTextCompare) > 0 Then If RsNota.Fields(C) = "TRANSFERENCIA" Or RsNota.Fields(C) = "TR" Then Zera = True
        If InStr(1, UCase(LcNome), UCase("STATUS"), vbTextCompare) > 0 Then If RsNota.Fields(C) = "INUTILIZADA" Then Zera = True
        If InStr(1, UCase(LcNome), UCase("DENEG"), vbTextCompare) > 0 Then Zera = True
        
        If InStr(1, UCase(LcNome), "STATUS", vbTextCompare) > 0 Then
          If InStr(1, UCase(RsNota.Fields(C)), "DENEG", vbTextCompare) > 0 Then
             RsNotaMdb(LcNome) = "Denegada"
          Else
             RsNotaMdb(LcNome) = RsNota.Fields(C)
          End If
           
        Else
           RsNotaMdb(LcNome) = RsNota.Fields(C)
        End If
        DoEvents
    Next
    If Zera Then
        RsNotaMdb("valorproduto") = 0
        RsNotaMdb("ValorNota") = 0
        RsNotaMdb("DESCONTO") = 0
        RsNotaMdb("baseicms") = 0
        RsNotaMdb("valoricms") = 0
        'RsNotaMdb("acrescimo") = 0
        'RsNotaMdb("basepis") = 0
        'RsNotaMdb("valorpis") = 0
       ' RsNotaMdb("basecofins") = 0
        'RsNotaMdb("valorcofins") = 0
        'RsNotaMdb("BaseCalculoIcmsSubst") = 0
        'RsNotaMdb("valorIcmsSubst") = 0
        'RsNotaMdb("valorIpi") = 0
        'RsNotaMdb("outrasdespesas") = 0
        'RsNotaMdb("frete") = 0
        'RsNotaMdb("seguro") = 0
    End If
    RsNotaMdb.Update
    RsNota.MoveNext
    DoEvents
Loop
RsNota.Close
RsNotaMdb.Close

End Function
Function AbreRecordsetRel(LcSql As String, RsAtual As ADODB.Recordset) As ADODB.Recordset

On Error GoTo ErroAbreRs
LcComentario = "- AbreRecordset - Criando Nova Instancia do RecordSet."
Set RsAtual = New ADODB.Recordset
LcComentario = "- AbreRecordset - Setando os Parametros do Recordset."
RsAtual.CursorType = adOpenDynamic ' adOpenStatic
RsAtual.CursorLocation = adUseClient
RsAtual.LockType = adLockReadOnly
RsAtual.Source = LcSql
RsAtual.ActiveConnection = conexaoAdo
Debug.Print conexaoAdo.ConnectionString

LcComentario = "- AbreRecordset - Abrindo o Recordset."
RsAtual.Open
Set AbreRecordsetRel = RsAtual
Exit Function

ErroAbreRs:
'If err.Number = 3709 Then
'   'abreconexao
'   Resume 0
'End If
'If LcExibemsg Then ErrosSistema = MsgBox(msg, 64, "erro Abrindo Tabela. ") Else ErrosSistema = 0
'MsgBox err.Description & err.Number
'Resume 0
logErro err.Number, err.Description, LcComentario
Resume Next
End Function




Private Sub Comissao_Click()
For a = 0 To LcTamanho
    If MtVendedor(a).Nome = Comissao.Text Then
       codigo.Text = MtVendedor(a).codigo
       Exit For
    End If
Next
If Len(Comissao.Text) = 0 Then Comissao.Text = ""
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, RsOpcao As Recordset
Dim LcEmpresa, LcEndereco, LcFone, LcCelular, Lccelular1, Lcemail, LcVer, LcVer1, LcCap As String

LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
GeraNota
AbreBase
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsOpcao = Dbbase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone
End If
If Not RsOpcao.EOF Then
   LcVer = RsOpcao!Msg
   LcVer1 = RsOpcao!Msg1
End If

    'Abertura do relatório de vendas
        
    CryRelatorio.DataFiles(0) = GLBase
    'If analitico Then
       'lctitulo = "Relatório de Comissões << ANALÍTICO >>"
     If GlImprimeSemLinha Then
        CryRelatorio.ReportFileName = App.Path & "\NotaSaida.rpt"
     Else
        CryRelatorio.ReportFileName = App.Path & "\NotaSaidasl.rpt"
     End If
    'Else
       'lctitulo = "Relatório de Comissões << SINTÉTICO >>"
   ' End If
    CryRelatorio.SortFields(0) = "+{ALID050.NUMNF}"
    
    CryRelatorio.CopiesToPrinter = Val(Txt1.Text)
    'If Comissao.Text <> "TODOS" Then
       'LcFormula = "{ALID201.VENDEDOR} = '" & codigo.Text & "'"
    'End If

  '== Inicio Filtro
  strData = CDate(Format(Datai.Text, "dd/mm/yyyy"))
  LcAno = Year(strData)
  LcMes = Month(strData)
  LcDia = Day(strData)
  LcDataInicio = LcAno & "," & LcMes & "," & LcDia
  LcChav1 = " date(" & LcDataInicio & ")"
         
  strData = CDate(Format(Dataf.Text, "dd/mm/yyyy"))
  LcAno = Year(strData)
  LcMes = Month(strData)
  LcDia = Day(strData)
  LcDataInicio = LcAno & "," & LcMes & "," & LcDia
  LcChav2 = " date(" & LcDataInicio & ")"
  If Len(LcFormula) <> 0 Then LcFormula = LcFormula & " And "
  LcFormula = LcFormula & "{ALID050.DTEMIS} >=" & LcChav1 & " And {ALID050.DTEMIS} <=" & LcChav2
  'LcFormula = LcFormula & " AND {ALID050.NATUREZA} <>'TR'"

'== fim filtro
'== fim filtro
CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Saída de Estoque por Período"

 CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
 CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
 CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'"
 CryRelatorio.Formulas(3) = "titulo='Relatório de Saída de Estoque por Período'"
 'CryRelatorio.Formulas(4) = "Versiculo='" & LcVer & "'"
 'CryRelatorio.Formulas(5) = "Versiculo1='" & LcVer1 & "'"
 CryRelatorio.Formulas(5) = "Celular='" & LcCelular & "'"
 CryRelatorio.Formulas(4) = "Celular1='" & Lccelular1 & "'"
 CryRelatorio.Formulas(6) = "email='" & Lcemail & "'"

If Impressora Then
   LcTipoSaida = 1
Else
   LcTipoSaida = 0
End If
CryRelatorio.SelectionFormula = LcFormula

CryRelatorio.Destination = LcTipoSaida
CryRelatorio.PrintReport
Me.Caption = LcCap
RsOpcao.Close
RsEmpresa.Close
Dbbase.Close
Set RsOpcao = Nothing
Set RsEmpresa = Nothing
Set Dbbase = Nothing
If CryRelatorio.LastErrorNumber > 0 Then MsgBox CryRelatorio.LastErrorString


End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Copias_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Dataf_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Dataf_LostFocus()
If Not IsDate(Dataf.Text) And Dataf.Text <> "  /  /  " Then
      MsgBox "A data digitada não é Válida...", 48, "Aviso"
      Dataf.SetFocus
End If
End Sub

Private Sub Datai_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
   
End Sub

Private Sub Datai_LostFocus()
If Not IsDate(Datai.Text) And Datai.Text <> "  /  /  " Then
      MsgBox "A data digitada não é Válida...", 48, "Aviso"
      Datai.SetFocus
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
DataS.Text = Format(GlDataSistema, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
'Top = Screen.Height / 2 - Height / 2
'Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
CarregaVendedor
'Me.Height = 2835
'Me.Width = 5370
'abreconexao
End Sub
Function CarregaVendedor()
On Error GoTo errc
Dim RsVendedor As Recordset
AbreBase
Set RsVendedor = Dbbase.OpenRecordset("ALID200", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcTamanho = 0
Do Until RsVendedor.EOF
   If Not IsNull(RsVendedor!Nome) Then
      ReDim Preserve MtVendedor(LcTamanho)
      MtVendedor(LcTamanho).codigo = RsVendedor!codigo
      MtVendedor(LcTamanho).Nome = RsVendedor!Nome
      Comissao.AddItem RsVendedor!Nome
      
      LcTamanho = LcTamanho + 1
   End If
   RsVendedor.MoveNext
Loop
If LcTamanho > 0 Then LcTamanho = LcTamanho - 1
'Comissao.AddItem "TODOS"
'Comissao.Text = "TODOS"
RsVendedor.Close
Set RsVendedor = Nothing
Exit Function
errc:
MsgBox err.Description & err.Number
'Resume 0
Exit Function
End Function
Function RecuperaNomeVendedor(codigo As String) As String
On Error GoTo errc
Dim RsVendedor As Recordset
AbreBase
Set RsVendedor = Dbbase.OpenRecordset("Select * from ALID200 where Codigo='" & codigo & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Dim LcNome As String
If Not RsVendedor.EOF Then
   If Not IsNull(RsVendedor!Nome) Then
      LcNome = RsVendedor!Nome
    End If
End If
RsVendedor.Close
Set RsVendedor = Nothing

RecuperaNomeVendedor = LcNome

Exit Function
errc:
MsgBox err.Description & err.Number
End Function
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'FechaConexao
End Sub

Private Sub Impressora_Click()
Copias.Visible = True
Label3.Visible = True
End Sub

Private Sub Impressora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub sintetico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub OptEntrada_Click()
If OptEntrada.Value = True Then
   IncluirDevolucao.Value = 1
End If
End Sub

Private Sub Video_Click()
Copias.Visible = False
Label3.Visible = False
End Sub

Private Sub Video_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub
