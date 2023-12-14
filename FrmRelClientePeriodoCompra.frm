VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmRelClienteComprasPeriodo 
   BackColor       =   &H00DBE1B7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Compras por Período - Clientes"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   3120
      Top             =   2040
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   4320
      Top             =   2760
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
   Begin VB.TextBox codigo 
      Height          =   285
      Left            =   3960
      TabIndex        =   13
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox Vendedor 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1440
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   3960
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F3"
      Height          =   615
      Left            =   3960
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   7
      Top             =   2280
      Width           =   1815
      Begin VB.OptionButton impressora 
         BackColor       =   &H00DBE1B7&
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Video 
         BackColor       =   &H00DBE1B7&
         Caption         =   "Video"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
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
      Top             =   480
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
      Caption         =   "Vendedor"
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
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   3840
      X2              =   3840
      Y1              =   0
      Y2              =   3600
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
      TabIndex        =   10
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
      Left            =   2160
      TabIndex        =   9
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
      Left            =   2280
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "FrmRelClienteComprasPeriodo"
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
Private Sub Cbo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Function GeraNota()
'On Error Resume Next
On Error GoTo errGera
Dim RsNota As ADODB.Recordset
Dim RsNotaMdb As Recordset
Dim LcSql As String
Dim LcNome As String
If Len(Vendedor.Text) = 0 Then
   LcSql = "Select * from alid050 where DTEMIS Between '" & Format(DataI.Text, "yyyy-mm-dd") & "' And '" & Format(DataF.Text, "yyyy-mm-dd") & "'"
Else
   LcSql = "Select * from alid050 where DTEMIS Between '" & Format(DataI.Text, "yyyy-mm-dd") & "' And '" & Format(DataF.Text, "yyyy-mm-dd") & "' and (status='EMITIDA' or Status='Autorizado o uso da NF-e') AND (VENDEDOR='" & CLng(Codigo.Text) & "' or VENDEDOR='" & Right("00000" & Codigo.Text, 5) & "')"
End If
AbreBase
'abreconexao
'MsgBox LcSql

Set RsNota = AbreRecordset(LcSql)
'MsgBox DEscricaoErro
Debug.Print LcSql
Set RsNotaMdb = Dbbase.OpenRecordset("Select * from alid050", dbOpenDynaset, dbSeeChanges, dbOptimistic)
RsNota.Requery
'===> Apagando Registros antigos
Do Until RsNotaMdb.EOF
    RsNotaMdb.Delete
    RsNotaMdb.MoveNext
Loop
'DBBASE.Execute "Delete
Do Until RsNota.EOF
    RsNotaMdb.AddNew
    For C = 0 To RsNota.Fields.Count - 1
        LcNome = RsNota.Fields(C).Name
        RsNotaMdb(LcNome) = RsNota.Fields(C)
        DoEvents
    Next
    RsNotaMdb.Update
    RsNota.MoveNext
    DoEvents
Loop
RsNota.Close
'FechaConexao
RsNotaMdb.Close
errGera:
'Resume Next
MsgBox err.Description & err.Number & " - " & LcNome
'Resume Next
End Function
Function RecuperaDataUltimaVendacliente(Cliente As String) As Date
Dim RsNota As ADODB.Recordset
Dim Data As Date
'== Verifica dados em Notas Fiscais
LcCriterio = "select * from alid050 where natureza<>'TR' and (DTemis Between '" & Format(DataI.Text, "yyyy-mm-dd") & "' And '" & Format(DataF.Text, "yyyy-mm-dd") & "') AND (STATUS='EMITIDA' or STATUS='Autorizado o uso da NF-e') and(Vendedor=" & Right("00000" & Codigo.Text, 5) & " OR Vendedor=" & CLng(Codigo.Text) & ") and(cliente='" & Cliente & "') order by codigo desc"

Set RsNota = AbreRecordset(LcCriterio, True)

If Not RsNota.EOF Then
   Data = RsNota!DTEMIS
Else
  Data = Null
End If

RecuperaDataUltimaVendacliente = Data
End Function
Private Sub Command1_Click()
On Error GoTo erroProcessaREl
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa   As Recordset
Dim RsOpcao     As Recordset
Dim RsCompra    As Recordset
Dim RsOrq       As Recordset
Dim RsNota      As Recordset
Dim RSNotaSaida As ADODB.Recordset
Dim LcEmpresa   As String
Dim LcEndereco  As String
Dim LcFone      As String
Dim LcCelular   As String
Dim Lccelular1  As String
Dim Lcemail     As String
Dim LcVer       As String
Dim LcVer1      As String
Dim LcCap       As String
Dim LcSql       As String
Dim LcSql1      As String
Dim LcPesq      As String
Dim LcAchou     As Boolean
Dim DataUl      As Date
'GeraNota

AbreBase
Dbbase.Execute "Delete from ClientesCompraramperiodo"

Set RsCompra = Dbbase.OpenRecordset("ClientesCompraramperiodo", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)

'Set RsOpcao = Dbbase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
'=== Limpa a base temporaria
err.Number = 0
Dbbase.Execute "Delete from ClientesCompraramperiodo"
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone & ""
End If

'=== Abertura do relatório de vendas
If Len(Codigo.Text) = 0 Then
   MsgBox "selecione o Vendedor para gerar o Relatorio.", 64, "Aviso"
   Vendedor.SetFocus
   Exit Sub
End If

'== Verifica dados em Orcamento
err.Number = 0
'== Verifica dados em Notas Fiscais
LcCriterio = "select * from alid050 where  (DTemis Between '" & Format(DataI.Text, "yyyy-mm-dd") & "' And '" & Format(DataF.Text, "yyyy-mm-dd") & "') AND (STATUS='EMITIDA' or STATUS='Autorizado o uso da NF-e') and(Vendedor=" & Right("00000" & Codigo.Text, 5) & " OR Vendedor=" & CLng(Codigo.Text) & ") order by codigo desc"

Set RSNotaSaida = AbreRecordset(LcCriterio, True)
'MsgBox DEscricaoErro
'LcCap = Me.Caption
If Not RSNotaSaida.EOF Then
   RSNotaSaida.MoveLast
   LcTotalReg = RSNotaSaida.RecordCount
   RSNotaSaida.MoveFirst
End If
a = 0
LcTempopassado = False
Timer1.Interval = 5000
Timer1.Enabled = True

Do While LcTempopassado = False
   DoEvents
Loop
Timer1.Enabled = False

Do Until RSNotaSaida.EOF
   a = a + 1
   Me.Caption = "Processando reg.:" & a & " de " & LcTotalReg
   DoEvents
   LcAchou = True
   LcPesq = "cliente='" & Right("00000" & RSNotaSaida!Cliente, 5) & "'"
   'If err.Number <> 3021 And err.Number > 0 Then
   '   Exit Do
   'Else
   '   err.Number = 0
   'End If
   If Not RsCompra.BOF Then RsCompra.MoveFirst
     DataUl = RecuperaDataUltimaVendacliente(Right("00000" & RSNotaSaida!Cliente, 5))
     If Not RsCompra.EOF Then
         RsCompra.FindFirst LcPesq
        
         If RsCompra.NoMatch Then
            RsCompra.AddNew
            RsCompra!Cliente = Right("00000" & RSNotaSaida!Cliente, 5)
            RsCompra!Data = DataUl
           ' RsCompra("Email") = RsClientes!Email & ""
            RsCompra.Update
         End If
    Else
            RsCompra.AddNew
            RsCompra!Cliente = Right("00000" & RSNotaSaida!Cliente, 5)
            RsCompra!Data = DataUl
           ' RsCompra("Email") = RsClientes!Email & ""
            RsCompra.Update
   End If
   RSNotaSaida.MoveNext
Loop
For a = 0 To 10000
   DoEvents
Next
LcTempopassado = False
Timer1.Interval = 5000
Timer1.Enabled = True

Do While LcTempopassado = False
   DoEvents
Loop
Timer1.Enabled = False

Me.Caption = LcCap
'== fim filtro
If Not LcAchou Then
   MsgBox "Não Exite Vendas Neste Periodo.", 64, "Aviso"
   Exit Sub
End If

CryRelatorio.DataFiles(0) = GLBase
If GlImprimeSemLinha Then
   CryRelatorio.ReportFileName = App.Path & "\RelClienteCompra.rpt"
Else
   CryRelatorio.ReportFileName = App.Path & "\RelClienteComprasl.rpt"
End If
CryRelatorio.CopiesToPrinter = Val(copias.Text)
CryRelatorio.SortFields(0) = "+{ALID001.razaosoc}"

CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowState = crptMaximized
CryRelatorio.WindowTitle = lctitulo

 CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
 CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
 CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'"
 CryRelatorio.Formulas(3) = "titulo='Relatório de Compras do Cliente'"
 'CryRelatorio.Formulas(4) = "Versiculo='" & LcVer & "'"
 'CryRelatorio.Formulas(5) = "Versiculo1='" & LcVer1 & "'"
 CryRelatorio.Formulas(5) = "Celular='" & LcCelular & "'"
 CryRelatorio.Formulas(4) = "Celular1='" & Lccelular1 & "'"
 CryRelatorio.Formulas(6) = "email='" & Lcemail & "'"

If impressora Then
   LcTipoSaida = 1
Else
   LcTipoSaida = 0
End If
'CryRelatorio.SelectionFormula = LcFormula

CryRelatorio.Destination = LcTipoSaida
CryRelatorio.DiscardSavedData = True
CryRelatorio.PrintReport
Me.Caption = LcCap
'RsOpcao.Close
RsEmpresa.Close
Dbbase.Close
'Set RsOpcao = Nothing
Set RsEmpresa = Nothing
Set Dbbase = Nothing
If CryRelatorio.LastErrorNumber > 0 Then MsgBox CryRelatorio.LastErrorString
Exit Sub
erroProcessaREl:
MsgBox "Ocorreu o seguinte erro Gerando o Relatorio:" & err.Number & " - " & err.Description


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
If Not IsDate(DataF.Text) And DataF.Text <> "  /  /  " Then
      MsgBox "A data digitada não é Válida...", 48, "Aviso"
      DataF.SetFocus
End If
End Sub

Private Sub Datai_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"

If KeyCode = 13 Then SendKeys "{TAB}"
   
End Sub

Private Sub Datai_LostFocus()
If Not IsDate(DataI.Text) And DataI.Text <> "  /  /  " Then
      MsgBox "A data digitada não é Válida...", 48, "Aviso"
      DataI.SetFocus
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
DataS.Text = Format(GlDataSistema, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
Me.Height = 3930
Me.Width = 5370
CarregaTelemarketing

End Sub
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
For a = 0 To LcTamanho
    If MtVendedor(a).Nome = Vendedor.Text Then
       Codigo.Text = MtVendedor(a).Codigo
       Exit For
    End If
Next

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
