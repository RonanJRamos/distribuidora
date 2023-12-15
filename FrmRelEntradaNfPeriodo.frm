VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmRelEntradaNfPeriodo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Entrada de Estoque Período"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Pesquisa"
      Height          =   1335
      Left            =   2040
      TabIndex        =   16
      Top             =   2160
      Width           =   2175
      Begin VB.OptionButton Iniciado 
         Caption         =   "Iniciado por"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Qualquer 
         Caption         =   "Em Qualquer Parte"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   570
         Width           =   1695
      End
      Begin VB.OptionButton Igual 
         Caption         =   "Igual a"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4695
   End
   Begin VB.ComboBox Fornecedor 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4695
   End
   Begin VB.TextBox Nome 
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   4440
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   5160
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirmar F3"
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   480
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
      Left            =   5160
      TabIndex        =   7
      Text            =   "1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1815
      Begin VB.OptionButton impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Video 
         Caption         =   "Video"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   375
      Left            =   120
      TabIndex        =   0
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
      Left            =   1920
      TabIndex        =   1
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
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
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor"
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
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   5040
      X2              =   5040
      Y1              =   0
      Y2              =   3480
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
      TabIndex        =   12
      Top             =   120
      Width           =   1185
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
      Left            =   1920
      TabIndex        =   11
      Top             =   120
      Width           =   1080
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
      Left            =   5160
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "FrmRelEntradaNfPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type DadoFornecedor
        codigo As String
        Nome As String
End Type
Private MtFornecedor() As DadoFornecedor
Private LcTam, a As Long


Private Sub Cbo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub analitico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub
Function GeraDados()
On Error GoTo errGera
Dim RsNota As ADODB.Recordset
Dim RsNotaMdb As Recordset
Dim LcSql As String
Dim LcNome As String
Dim db As Database
Dim C As Integer
LcSql = "Select * from EntradaNf where data Between '" & Format(Datai.Text, "yyyy-mm-dd") & "' And '" & Format(Dataf.Text, "yyyy-mm-dd") & "'"
If fornecedor.ListIndex > -1 Then LcSql = LcSql & " and CLICRED='" & Nome.Text & "'"
'AbreBase
Set db = OpenDatabase(GLBase)
'==> Apagando Registros
db.Execute "Delete * from EntradaNf"

Set RsNota = AbreRecordset(LcSql, True)
Set RsNotaMdb = db.OpenRecordset("Select * from EntradaNf", dbOpenDynaset, dbSeeChanges, dbOptimistic)
RsNota.Requery
'===> Apagando Registros antigos
'Do Until RsNotaMdb.EOF
'    DoEvents
'    RsNotaMdb.Delete
'    RsNotaMdb.MoveNext
'Loop

Do Until RsNota.EOF
    RsNotaMdb.AddNew
    For C = 0 To RsNota.Fields.Count - 1
        LcNome = RsNota.Fields(C).Name
        If UCase(LcNome) = "CODIGO" Then
            RsNotaMdb(LcNome) = CDbl(RsNota.Fields(C))
        Else
            RsNotaMdb(LcNome) = RsNota.Fields(C)
        End If
        DoEvents
    Next
    RsNotaMdb.Update
    RsNota.MoveNext
    DoEvents
Loop
RsNota.Close
RsNotaMdb.Close
geradados1
Exit Function
errGera:
'MsgBox err.Description & err.Number
Resume Next
End Function
Function geradados1()
On Error GoTo errGera
Dim RsNota As ADODB.Recordset
Dim RsNotaMdb As Recordset
Dim Rs As Recordset
Dim LcSql As String
Dim LcNome As String
Dim db As Database

'AbreBase
Set db = OpenDatabase(GLBase)
'==> Apagando Registros
db.Execute "Delete * from ItensEntradaNf"
Set RsNotaMdb = db.OpenRecordset("Select * from ItensEntradaNf", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set Rs = db.OpenRecordset("Select * from EntradaNf", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcSql = "Select * from ItensEntradaNf where (ItensEntradaNf.data Between #" & Format(Datai.Text, "mm/dd/yy") & "# And #" & Format(Dataf.Text, "mm/dd/yy") & "#)  and descricao like '" & UCase(Text1.Text) & "%' and fornecedor like '" & fornecedor.Text & "%';"
'Debug.Print LcSql

 Set RsNota = AbreRecordset(LcSql, True)
   Do Until RsNota.EOF
        RsNotaMdb.AddNew
        For C = 0 To RsNota.Fields.Count - 1
            LcNome = RsNota.Fields(C).Name
            If UCase(LcNome) = "CODIGO" Then
                RsNotaMdb(LcNome) = CDbl(RsNota.Fields(C))
            Else
                RsNotaMdb(LcNome) = RsNota.Fields(C)
            End If
            DoEvents
        Next
        RsNotaMdb.Update
        RsNota.MoveNext
        DoEvents
    Loop
'RsNota.Close
RsNotaMdb.Close
Exit Function
errGera:
'MsgBox err.Description & err.Number
'Resume 0
Resume Next

End Function
Function carregaFornecedor()
Dim LcEmpresa, LcEndereco, LcFone, LcVer, LcCap, LcVer1 As String
Dim RsEmpresa As Recordset
AbreBase
LcTam = 0
Set RsEmpresa = Dbbase.OpenRecordset("select * from alid002 order by razaosoc") ', dbOpenDynaset, dbSeeChanges, dbOptimistic)
Do Until RsEmpresa.EOF
    ReDim Preserve MtFornecedor(LcTam)
    If Not IsNull(RsEmpresa!razaosoc) Then
        MtFornecedor(LcTam).codigo = RsEmpresa!codigo
        MtFornecedor(LcTam).Nome = RsEmpresa!razaosoc
        fornecedor.AddItem RsEmpresa!razaosoc
        LcTam = LcTam + 1
    End If
    RsEmpresa.MoveNext
Loop
If LcTam > 0 Then LcTam = LcTam - 1
RsEmpresa.Close
Dbbase.Close
Set RsEmpresa = Nothing
Set dbbasee = Nothing

End Function


Private Sub Command1_Click()
'On Error Resume Next
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, RsOpcao As Recordset
Dim LcEmpresa, LcEndereco, LcFone, LcCelular, Lccelular1, Lcemail, LcVer, LcVer1, LcCap As String
geradados1
AbreBase
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsOpcao = Dbbase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone
End If
'If Not RsOpcao.EOF Then
'   LcVer = RsOpcao!msg
'   LcVer1 = RsOpcao!Msg1
'End If
    If Iniciado Then LcFormula = "{ItensEntradaNf.descricao} like '" & UCase(Text1.Text) & "*'"
    If Qualquer Then LcFormula = "{ItensEntradaNf.descricao} like '*" & UCase(Text1.Text) & "*'"
    If Igual Then LcFormula = "{ItensEntradaNf.descricao}='" & UCase(Text1.Text) & "'"
    'If Len(Text1.Text) > 0 Then
    '   CryRelatorio.SortFields(0) = "+{ItensEntradaNf.descricao}"
    'Abertura do relatório de vendas
    '  End If
    'Cryrelatorio.Connect = "DSN=Relatorio"
    CryRelatorio.DataFiles(0) = GLBase
    'If analitico Then
       'lctitulo = "Relatório de Comissões << ANALÍTICO >>"
     If GlImprimeSemLinha Then
        CryRelatorio.ReportFileName = App.Path & "\notaentradaproduto.rpt"
     Else
        CryRelatorio.ReportFileName = App.Path & "\notaentradaproduto.rpt"
     End If
    'Else
       'lctitulo = "Relatório de Comissões << SINTÉTICO >>"
   ' End If
    CryRelatorio.SortFields(0) = "+{ItensEntradaNf.descricao}"
    
    CryRelatorio.CopiesToPrinter = Val(Copias.Text)
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
  LcFormula = LcFormula & "{ItensEntradaNf.DATA} >=" & LcChav1 & " And {ItensEntradaNf.data} <=" & LcChav2
  'LcFormula = LcFormula & " AND {ALID050.NATUREZA} <>'TR'"
  LcFormula = LcFormula & " and {ItensEntradaNf.fornecedor} like'" & fornecedor.Text & "*'"
'== fim filtro
'== fim filtro
CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Entrada de Estoque por Produto"

 CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
 CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
 CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'"
 CryRelatorio.Formulas(3) = "titulo='Relatório de Entrada de Estoque por Período'"
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
'MsgBox LcFormula
CryRelatorio.SelectionFormula = LcFormula

CryRelatorio.Destination = LcTipoSaida
CryRelatorio.PrintReport
Me.Caption = LcCap
'RsOpcao.Close
RsEmpresa.Close
Dbbase.Close
'Set RsOpcao = Nothing
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
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
Me.Height = 3900
Me.Width = 6540
carregaFornecedor
End Sub

Private Sub fornecedor_Click()
For a = 0 To LcTam
    If MtFornecedor(a).Nome = fornecedor.Text Then
       Nome.Text = MtFornecedor(a).codigo
       Exit For
    End If
Next

End Sub

Private Sub Fornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
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

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
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
