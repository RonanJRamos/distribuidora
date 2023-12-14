VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form ContabilSaida 
   BackColor       =   &H00B3E9FD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatorio contabil de Saidas"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdGerar 
      Caption         =   "GerarRel"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   2880
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSMask.MaskEdBox Dataf 
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      Height          =   195
      Left            =   1440
      TabIndex        =   5
      Top             =   480
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo de "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "ContabilSaida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Rs       As ADODB.Recordset
Private Rel      As New CryRelSaida
Private MtErro() As String
Private Sub CmdFechar_Click()
On Error Resume Next
Unload Me
End Sub
Sub BuscaDados()
On Error GoTo erroBusca
Dim RsFor As Recordset
Dim RsNota As ADODB.Recordset
Dim RsdadosNota As ADODB.Recordset
Dim StrSql As String
Dim CFOP_Saida As String
Dim NomeForn As String

Dim db As Database
Set db = OpenDatabase(GLBase)

'==> Exclui dados antigo do rel
StrSql = "Delete from RelsaidaContabil"
afetados = ExecutaSql(StrSql)
'==> Abre a tb de notas

StrSql = "SELECT alid050.Desconto,alid050.NUMNF, alid050.DTEMIS, alid050.CLIENTE, alid052.VALUNIT, alid052.QTDE, alid052.icms, alid052.CST, alid052.CFOP,alid052.valoricms,alid052.valorpis,alid052.valorcofins,alid052.baseicms,alid052.vICMSSubstituto, (qtde*valunit)-alid052.DESCONTO + alid052.despAcessorias+ alid052.Seguro + alid052.frete AS Total " & _
         "FROM alid050 INNER JOIN alid052 ON alid050.NUMNF = alid052.NUMNF"
StrSql = StrSql & " where alid050.DTEMIS Between '" & Format(DataI.Text, "yyyy-mm-dd") & "' And '" & Format(DataF.Text, "yyyy-mm-dd") & "'"
StrSql = StrSql & " and Status like 'Autorizado%'"
StrSql = StrSql & " order by alid050.NUMNF"
Debug.Print StrSql
Set RsNota = AbreRecordset(StrSql, True)
Do Until RsNota.EOF
   '==> Procura o cliente
   StrSql = "Select * from alid001 where codigo='" & RsNota!Cliente & "'"
   Set RsFor = db.OpenRecordset(StrSql)
   If Not RsFor.EOF Then
      Estado = RsFor!Estado & ""
      NomeForn = Left(RsFor!RazaoSoc, 50) & ""
   Else
      Estado = ""
      NomeForn = ""
   End If
   CFOP_Saida = RsNota!CFOP & ""
      
   '==> Cria a string de inserção do rel contabil
   CFOP_Saida = Replace(CFOP_Saida, ".", "")
   CFOP_Saida = Replace(CFOP_Saida, ",", "")
   CFOP_Saida = Replace(CFOP_Saida, "-", "")
   CFOP_Saida = Replace(CFOP_Saida, " ", "")
   If CFOP_Saida = "5102" Or CFOP_Saida = "6102" Or CFOP_Saida = "5403" Or CFOP_Saida = "6403" Then
     If RsNota!icms = 0 Then
      If CFOP_Saida = "5102" Or CFOP_Saida = "6102" Or CFOP_Saida = "5403" Or CFOP_Saida = "6403" Then
          If UCase(Estado) = "MG" Then
              CFOP_Saida = "5403"
          Else
              CFOP_Saida = "6403"
          End If
       End If
     Else
       If CFOP_Saida = "5102" Or CFOP_Saida = "6102" Or CFOP_Saida = "5403" Or CFOP_Saida = "6403" Then
          If UCase(Estado) = "MG" Then
              CFOP_Saida = "5102"
          Else
              CFOP_Saida = "6102"
          End If
        End If
     End If
   End If
   Desconto = IIf(Not IsNull(RsNota!Desconto), RsNota!Desconto, 0)
   ValorProduto = CDbl(RsNota!VALUNIT) * CDbl(RsNota!Qtde)
   ValorTotal = CDbl(RsNota!total)
   ValorBase = CDbl(RsNota!BaseIcms)
   valorIcms = RsNota!valorIcms
   
   Pis = IIf(Not IsNull(RsNota!valorpis), RsNota!valorpis, 0)
   Cofins = IIf(Not IsNull(RsNota!valorcofins), RsNota!valorcofins, 0)
   st = IIf(Not IsNull(RsNota!vICMSSubstituto), RsNota!vICMSSubstituto, 0)
   '==> Cria a string de inserção do rel contabil
   StrSql = "Insert into RelsaidaContabil (NF,cfop,Codcliente,Nome," & _
            "ValorProduto,BaseIcms,Icms,TotalNota,Entrada,Desconto,PercIcms,Pis,Confins,ST) Values ('" & _
            RsNota!NumNf & "','" & _
            Replace(CFOP_Saida, ".", "") & "','" & _
            RsNota!Cliente & "','" & _
            IIf(Not RsFor.EOF, Replace(NomeForn, "'", "''"), "") & "'," & _
            Replace(ValorProduto, ",", ".") & "," & _
            Replace(IIf(valorIcms > 0, ValorBase, 0), ",", ".") & "," & _
            Replace(CDbl(IIf(Len(valorIcms) > 0, IIf(IsNumeric(valorIcms), valorIcms, 0), 0)), ",", ".") & "," & _
            Replace(CDbl(ValorTotal), ",", ".") & ",'" & _
            Format(RsNota!DTEMIS, "yyyy-mm-dd") & "'," & _
            Replace(CDbl(Desconto), ",", ".") & "," & Replace(CDbl(RsNota!icms), ",", ".") & "," & _
            Replace(CDbl(Pis), ",", ".") & "," & Replace(CDbl(Cofins), ",", ".") & "," & Replace(CDbl(st), ",", ".") & ")"
   'Debug.Print StrSql
   afetados = ExecutaSql(StrSql)
   'Debug.Print acessoado.DEscricaoErro
   Set RsFor = Nothing
   RsNota.MoveNext
Loop
Set RsNota = Nothing


Exit Sub
erroBusca:
MsgBox err.Description & err.Number
Resume Next
End Sub
Function BuscaIcmsProduto(codigoproduto As Long, Optional Nome As String = "") As Double

Dim RsProduto As ADODB.Recordset
Dim StrSql As String
Dim db As Database
Dim RsAnt As Recordset
Dim Achou As Boolean
StrSql = "Select * from produtos where codigo=" & codigoproduto

Set RsProduto = AbreRecordset(StrSql, True)

If Not RsProduto.EOF Then
    If Len(RsProduto!cst) > 0 Then

        If CDbl(RsProduto!cst) = 60 Or CDbl(RsProduto!cst) = 16 Or CDbl(RsProduto!cst) = 26 Then
           BuscaIcmsProduto = 0
        Else
           BuscaIcmsProduto = 18
        End If
    Else
       BuscaIcmsProduto = 18
    End If
        
Else
   Set db = OpenDatabase(GLBase)
   StrSql = "Select * from alid009ant where cod='" & Right("00000" & codigoproduto, 5) & "'"
   Set RsAnt = db.OpenRecordset(StrSql)
   If Not RsAnt.EOF Then
        If Len(RsAnt!cst) > 0 Then

            If CDbl(RsAnt!cst) = 60 Or CDbl(RsAnt!cst) = 16 Or CDbl(RsAnt!cst) = 26 Then
               BuscaIcmsProduto = 0
            Else
               BuscaIcmsProduto = 18
            End If
        Else
           BuscaIcmsProduto = 18
        End If
   Else
      ErroEncontrado = True
      a = UBound(MtErro) + 1
      Achou = False
      LcMsg = Chr(13) & " O Produto " & codigoproduto & " - " & Nome & " não encontrado."
      For b = 0 To UBound(MtErro)
         If MtErro(b) = LcMsg Then
            Achou = True
            Exit For
         End If
      Next
      If Not Achou Then
         ReDim Preserve MtErro(a)
         MtErro(a) = LcMsg
      End If
   End If
End If


End Function
Private Sub CmdGerar_Click()
On Error Resume Next
Dim StrSql          As String
LcCap = Me.Caption
Me.Caption = "Aguarde, processando os dados..."
Screen.MousePointer = 11
BuscaDados
Screen.MousePointer = 0
Me.Caption = LcCap
Screen.MousePointer = vbHourglass
StrSql = "SELECT RelSaidaContabil.Codigo, RelSaidaContabil.NF, RelSaidaContabil.CFOP, RelSaidaContabil.CodCliente, RelSaidaContabil.Nome, RelSaidaContabil.ValorProduto, RelSaidaContabil.BaseIcms, RelSaidaContabil.BaseSubs, RelSaidaContabil.ValorIcmsSubs, RelSaidaContabil.Icms, RelSaidaContabil.TotalNota, RelSaidaContabil.Entrada, RelSaidaContabil.QuantidadeSaida, RelSaidaContabil.Desconto, RelSaidaContabil.PercIcms"
StrSql = StrSql & ",RelSaidaContabil.Pis,RelSaidaContabil.Confins,RelSaidaContabil.ST"
StrSql = StrSql & " FROM RelSaidaContabil"
StrSql = StrSql & StrWhe & " order by  RelSaidaContabil.Entrada"
Set Rs = AbreRecordset(StrSql, True)
'MsgBox Rs.RecordCount
Load Relatorios
With Relatorios
     Rel.DiscardSavedData
     Rel.Database.SetDataSource Rs
     Rel.Subreport1.OpenSubreport.DiscardSavedData
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

If IsDate(Format(DataI.Text, "dd/mm/yy")) And IsDate(Format(DataF.Text, "dd/mm/yy")) Then
   lctitulo = "Relatorio de Notas de Saida: " & DataI.Text & " à " & DataF.Text
   Else
   lctitulo = "Relatorio de Notas de Saida"
End If
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


