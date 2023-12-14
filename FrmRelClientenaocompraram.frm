VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmRelClientenaocompraram 
   BackColor       =   &H00DBE1B7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Clientes que Não Compraram - Período"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   4080
      Top             =   2040
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   3240
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.TextBox Codigo 
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox Vendedor 
      Height          =   315
      ItemData        =   "FrmRelClientenaocompraram.frx":0000
      Left            =   0
      List            =   "FrmRelClientenaocompraram.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1680
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   3960
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirmar F3"
      Height          =   615
      Left            =   3960
      TabIndex        =   6
      Top             =   600
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
      Left            =   2280
      TabIndex        =   5
      Text            =   "1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE1B7&
      Caption         =   "Saída"
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   1815
      Begin VB.OptionButton impressora 
         BackColor       =   &H00DBE1B7&
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Video 
         BackColor       =   &H00DBE1B7&
         Caption         =   "Video"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
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
      TabIndex        =   1
      Top             =   840
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Será considerado os clientes que foram cadastrados antes da data inicial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telemarketing"
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
      TabIndex        =   12
      Top             =   1320
      Width           =   1515
   End
   Begin VB.Line Line1 
      X1              =   3840
      X2              =   3840
      Y1              =   600
      Y2              =   3720
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
      TabIndex        =   11
      Top             =   600
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
      Left            =   2160
      TabIndex        =   10
      Top             =   600
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
      Left            =   2280
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "FrmRelClientenaocompraram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TipoVend
      Codigo As String
      Nome As String
End Type
Private LcTamanho, a As Integer
Private MtVendedor() As TipoVend
Private LcTempopassado As Boolean
Function CarregaTelemarketing()
On Error GoTo errc
Dim RsVendedor As Recordset
AbreBase
Set RsVendedor = Dbbase.OpenRecordset("ALID200", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcTamanho = 0
Do Until RsVendedor.EOF
   ReDim Preserve MtVendedor(LcTamanho)
   MtVendedor(LcTamanho).Codigo = RsVendedor!Codigo
   MtVendedor(LcTamanho).Nome = RsVendedor!Nome
   Vendedor.AddItem RsVendedor!Nome
   RsVendedor.MoveNext
   LcTamanho = LcTamanho + 1
   DoEvents
Loop
If LcTamanho > 0 Then LcTamanho = LcTamanho - 1
'Comissao.AddItem "TODOS"
'Comissao.Text = "TODOS"
RsVendedor.Close
Set RsVendedor = Nothing
Exit Function
errc:

Exit Function

End Function

Private Sub Cbo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub analitico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub
Function separanaocompra()
Dim RsCliente As Recordset, RsNota As Recordset, naocompraram As Recordset
AbreBase
Set RsCliente = Dbbase.OpenRecordset("alid001", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsNota = Dbbase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set naocompraram = Dbbase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)

Do Until RsCliente.EOF
  ' LcCriterio
Loop
End Function

Private Sub Command1_Click()
'On Error Resume Next
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, RsOpcao As Recordset, RsClientes As Recordset
Dim RSNotaSaida As ADODB.Recordset, RsCompra As Recordset, Rsorcamento As Recordset
Dim LcEmpresa, LcEndereco, LcFone, LcCelular, Lccelular1, Lcemail, LcVer, LcVer1, LCCap
Dim bb As Database
Dim LcAchouOrc, LcAchouNt As Boolean
Dim LcTotalCliente, LcClientesAtual As Long
Set bb = OpenDatabase(GLBase, False, False) ' "dBASE III;")
If Len(Vendedor.Text) > 0 Then
   LcCriterio = "select * from alid001 where TelemarketingAtende='" & Vendedor.Text & "' and dataCadastro<=#" & Format(Datai.Text, "mm/dd/yyyy") & "#"
Else
   LcCriterio = "select * from alid001 where dataCadastro<=#" & Format(Datai.Text, "mm/dd/yyyy") & "#"
End If
'abreconexao
Set RsEmpresa = bb.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsOpcao = Bb.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)

Set RsClientes = bb.OpenRecordset(LcCriterio, dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set Rsorcamento = bb.OpenRecordset("orcamento", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RSNotaSaida = bb.OpenRecordset("ALID050", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsCompra = bb.OpenRecordset("naocompraram", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcTempopassado = False
Timer1.Interval = 5000
Timer1.Enabled = True

Do While LcTempopassado = False
   DoEvents
Loop
Timer1.Enabled = False

'RSNotaSaida.MoveLast
'MsgBox RSNotaSaida.RecordCount
'RSNotaSaida.MoveFirst
LCCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone
End If
Do Until RsCompra.EOF
   RsCompra.Delete
   RsCompra.MoveNext
   DoEvents
Loop
LcAchouOrc = False
LcAchouNt = False
'If Not Rsorcamento.EOF Then LcAchouOrc = True
'If Not RSNotaSaida.EOF Then LcAchouNt = True
If Not RsClientes.EOF Then
   RsClientes.MoveLast
   LcTotalCliente = RsClientes.RecordCount
   RsClientes.MoveFirst
End If
LcClientesAtual = 1
lcEncontrados = 0
Do Until RsClientes.EOF
   '=== Criterio para Busca no Orcamento.
   DoEvents
   Me.Caption = " Verificando Cliente " & LcClientesAtual & " de " & LcTotalCliente & ". Encontrados " & lcEncontrados
     
   '=== Criterio para Busca no NotaFiscal.
   LcCriterio = "select * from alid050 where natureza<>'TR' and (DTemis Between '" & Format(Datai.Text, "yyyy-mm-dd") & "' And '" & Format(Dataf.Text, "yyyy-mm-dd") & "') AND (STATUS='EMITIDA' or STATUS='Autorizado o uso da NF-e') and (cliente='" & CLng(RsClientes!Codigo) & "' or cliente='" & Right("00000" & RsClientes!Codigo, 5) & "')  "
   Set RSNotaSaida = AbreRecordset(LcCriterio)
'MsgBox DEscricaoErro

   If RSNotaSaida.EOF Then
      Dim DataUltimaCompra As String
      Dim RsUltimaCompra As ADODB.Recordset
      Dim StrUltima As String
      StrUltima = "Select DTemis From alid050 where (cliente='" & CLng(RsClientes!Codigo) & "' or cliente='" & Right("00000" & RsClientes!Codigo, 5) & "') order by DTemis desc limit 1"
      Set RsUltima = AbreRecordset(StrUltima, True)
      If Not RsUltima.EOF Then
         If IsDate(RsUltima!DTemis) Then
             DataUltimaCompra = Format(RsUltima!DTemis, "dd/mm/yyyy")
         Else
            DataUltimaCompra = ""
         End If
      Else
          DataUltimaCompra = ""
      End If
      RsCompra.AddNew
      RsCompra("Nome") = RsClientes!RazaoSoc & ""
      RsCompra("Cgc") = RsClientes!CGC & ""
      RsCompra("codigo") = RsClientes!Codigo & ""
      RsCompra("tele") = RsClientes!TelemarketingAtende & ""
      RsCompra("fone") = RsClientes!Fone1 & ""
      RsCompra("Email") = RsClientes!Email & ""
      RsCompra("DataUltimaCompra") = DataUltimaCompra
      If Not IsNull(RsClientes!LimiteCredito) Then RsCompra("LimiteCredito") = RsClientes!LimiteCredito & ""
      
      RsCompra.Update
      lcEncontrados = lcEncontrados + 1
   End If
   RsClientes.MoveNext
   LcClientesAtual = LcClientesAtual + 1
   DoEvents
Loop
LcTempopassado = False
Timer1.Interval = 5000
Timer1.Enabled = True

Do While LcTempopassado = False
   DoEvents
Loop
Timer1.Enabled = False

CryRelatorio.DataFiles(0) = GLBase
If GlImprimeSemLinha Then
   CryRelatorio.ReportFileName = App.Path & "\NaoCompram.rpt"
Else
   CryRelatorio.ReportFileName = App.Path & "\NaoCompramsl.rpt"
End If
'CryRelatorio.CopiesToPrinter = Val(Txt1.Text)
CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = lctitulo

 CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
 CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
 CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'"
 CryRelatorio.Formulas(3) = "titulo='Relatório de Clientes Nao Compram'"
 'CryRelatorio.Formulas(4) = "Versiculo='" & LcVer & "'"
 'CryRelatorio.Formulas(5) = "Versiculo1='" & LcVer1 & "'"
 CryRelatorio.Formulas(3) = "Celular='" & LcCelular & "'"
 CryRelatorio.Formulas(4) = "Celular1='" & Lccelular1 & "'"
 CryRelatorio.Formulas(6) = "email='" & Lcemail & "'"

If impressora Then
   LcTipoSaida = 1
Else
   LcTipoSaida = 0
End If

CryRelatorio.Destination = LcTipoSaida
CryRelatorio.PrintReport
Me.Caption = LCCap
RsEmpresa.Close
'RsOpcao.Close
RsClientes.Close
Rsorcamento.Close
RSNotaSaida.Close
RsCompra.Close
bb.Close
Set RsEmpresa = Nothing
'Set RsOpcao = Nothing
Set RsClientes = Nothing
Set Rsorcamento = Nothing
Set RSNotaSaida = Nothing
Set RsCompra = Nothing
Set bb = Nothing
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
CarregaTelemarketing

DataS.Text = Format(GlDataSistema, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
Me.Height = 4095
Me.Width = 5370


End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
End Sub

Private Sub Impressora_Click()
copias.Visible = True
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

Private Sub Timer1_Timer()
LcTempopassado = True
End Sub

Private Sub Vendedor_Click()
Dim a As Integer
For a = 0 To LcTamanho
    If MtVendedor(a).Nome = Vendedor.Text Then
       Codigo.Text = MtVendedor(a).Codigo
       Exit For
    End If
Next
End Sub

Private Sub Vendedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Video_Click()
copias.Visible = False
Label3.Visible = False
End Sub

Private Sub Video_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub
