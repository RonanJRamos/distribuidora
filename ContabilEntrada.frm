VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form ContabilEntrada 
   BackColor       =   &H00B3E9FD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Posição Contabil de Entradas"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3285
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   2880
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton CmdGerar 
      Caption         =   "GerarRel"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin MSMask.MaskEdBox Dataf 
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   600
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
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
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
      TabIndex        =   5
      Top             =   240
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   90
   End
End
Attribute VB_Name = "ContabilEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Rs As ADODB.Recordset
Private Rel As New CryRelEntrada
Private Sub CmdFechar_Click()
On Error Resume Next
Unload Me
End Sub
Sub BuscaDados()
On Error GoTo erroBusca
Dim RsFor As Recordset
Dim RsNota As ADODB.Recordset
Dim StrSql As String
Dim db As Database
Dim CFOP_Entrada As String
Set db = OpenDatabase(GLBase)

'==> Exclui dados antigo do rel
StrSql = "Delete from RelEntradaContabil"
afetados = ExecutaSql(StrSql)
'==> Abre a tb de notas
'StrSql = "Select * from entradanf where data Between '" & Format(Datai.Text, "yyyy-mm-dd") & "' And '" & Format(Dataf.Text, "yyyy-mm-dd") & "' order by data"
StrSql = "SELECT itensentradanf.data AS dataitem,entradanf.IcmsSubst,entradanf.BaseIcmsSubst,itensentradanf.ipi, entradanf.NF,entradanf.BaseIcms, entradanf.CLICRED, entradanf.DATA, itensentradanf.QTDE, itensentradanf.VALUNIT, itensentradanf.ValorTotal, itensentradanf.Icms, itensentradanf.cfop " & _
         "FROM entradanf INNER JOIN itensentradanf ON entradanf.NF = itensentradanf.NUMNF " & _
         "WHERE (((entradanf.DATA) Between '" & Format(DataI.Text, "yyyy-mm-dd") & "' And '" & Format(DataF.Text, "yyyy-mm-dd") & "'));"

Set RsNota = AbreRecordset(StrSql, True)
Debug.Print StrSql
Do Until RsNota.EOF
  If CDate(RsNota!dataitem) >= CDate(DataI.Text) And CDate(RsNota!dataitem) <= CDate(DataF.Text) Then

   '==> Procura o fornecedor
   StrSql = "Select * from alid002 where codigo='" & RsNota!clicred & "'"
   Set RsFor = db.OpenRecordset(StrSql)
   If Not RsFor.EOF Then Estado = RsFor!Estado & ""
   CFOP_Entrada = RsNota!CFOP
   '==> Cria a string de inserção do rel contabil
   If CFOP_Entrada = "1102" Or CFOP_Entrada = "2102" Or CFOP_Entrada = "1403" Or CFOP_Entrada = "2403" Then
     If RsNota!icms = 0 Then
      If CFOP_Entrada = "1102" Or CFOP_Entrada = "2102" Or CFOP_Entrada = "1403" Or CFOP_Entrada = "2403" Then
          If UCase(Estado) = "MG" Then
              CFOP_Entrada = "1403"
          Else
              CFOP_Entrada = "2403"
          End If
       End If
     Else
       If CFOP_Entrada = "1102" Or CFOP_Entrada = "2102" Or CFOP_Entrada = "1403" Or CFOP_Entrada = "2403" Then
          If UCase(Estado) = "MG" Then
              CFOP_Entrada = "1102"
          Else
              CFOP_Entrada = "2102"
          End If
        End If
     End If
   End If
             
            StrSql = "Insert into RelEntradaContabil (NF,cfop,CodFornecedor,NomeFornecedor," & _
                     "ValorProduto,BaseIcms,Icms,PerIcms,Ipi,TotalNota,Entrada,BaseSubs,ValorIcmsSubs) Values ('" & _
                     RsNota!NF & "','" & _
                     CFOP_Entrada & "','" & _
                     RsNota!clicred & "','" & _
                     IIf(Not RsFor.EOF, RsFor!RAZAOSOC, "") & "'," & _
                     Replace(CDbl(RsNota!ValorTotal), ",", ".") & "," & _
                     Replace(CDbl(IIf(RsNota!icms > 0, RsNota!ValorTotal, 0)), ",", ".") & "," & _
                     Replace(CDbl(RsNota!icms) * (CDbl(IIf(RsNota!icms > 0, RsNota!ValorTotal, 0)) / 100), ",", ".") & "," & _
                     Replace(CDbl(RsNota!icms), ",", ".") & "," & _
                     Replace(CDbl(RsNota!ipi), ",", ".") & "," & _
                     Replace(CDbl(RsNota!ValorTotal), ",", ".") & ",'" & _
                     Format(RsNota!Data, "yyyy-mm-dd") & "'," & _
                     Replace(CDbl(RsNota!BaseIcmsSubst), ",", ".") & "," & _
                     Replace(CDbl(RsNota!IcmsSubst), ",", ".") & ")"
            afetados = ExecutaSql(StrSql)
            Set RsFor = Nothing
  End If
   RsNota.MoveNext
Loop
Set RsNota = Nothing


Exit Sub
erroBusca:
MsgBox err.Description & err.Number
Resume 0
End Sub

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
StrSql = "SELECT RelEntradaContabil.Codigo, RelEntradaContabil.NF, RelEntradaContabil.CFOP, RelEntradaContabil.CodFornecedor, RelEntradaContabil.NomeFornecedor, RelEntradaContabil.ValorProduto, RelEntradaContabil.BaseIcms, RelEntradaContabil.Icms, RelEntradaContabil.BaseSubs, RelEntradaContabil.ValorIcmsSubs, RelEntradaContabil.Ipi, RelEntradaContabil.TotalNota, RelEntradaContabil.Entrada, RelEntradaContabil.QuantidadeSaida, RelEntradaContabil.PerICMS"
StrSql = StrSql & " FROM RelEntradaContabil"
StrSql = StrSql & StrWhe & " order by  RelEntradaContabil.Entrada"
Set Rs = AbreRecordset(StrSql, True)
Load Relatorios
With Relatorios
     Rel.DiscardSavedData
     Rel.Database.SetDataSource Rs
     .CRViewer1.ReportSource = Rel
     setaformula
      .CRViewer1.ViewReport
End With
Relatorios.Show
Screen.MousePointer = vbDefault
End Sub

Sub GeraRel()
On Error Resume Next

CryRelatorio.DataFiles(0) = GLBase
CryRelatorio.ReportFileName = App.Path & "\contabilentrada.rpt"

CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Relatório contabil de entrada."

CryRelatorio.Formulas(0) = "periodo=' de " & DataI.Text & " a " & DataF.Text & "'"
LcTipoSaida = 0
CryRelatorio.SortFields(0) = "+{RelEntradaContabil.nf}"
CryRelatorio.Destination = LcTipoSaida
CryRelatorio.PrintReport
If CryRelatorio.LastErrorNumber > 0 Then MsgBox CryRelatorio.LastErrorString

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
   lctitulo = "Relatorio de Notas de Entrada: " & DataI.Text & " à " & DataF.Text
   Else
   lctitulo = "Relatorio de Notas de Entrada"
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
