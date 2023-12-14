VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Transferencia 
   Caption         =   "Transferencia entre Galpões."
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   11040
      TabIndex        =   21
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Quantidade 
      Height          =   285
      Left            =   8880
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox Galpao 
      Height          =   315
      ItemData        =   "Transferencia.frx":0000
      Left            =   7320
      List            =   "Transferencia.frx":000A
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton CmdFiltrar 
      Caption         =   "Filtrar"
      Height          =   375
      Left            =   11040
      TabIndex        =   19
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Fl_Produto 
      Height          =   285
      Left            =   3120
      TabIndex        =   18
      Top             =   240
      Width           =   7815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5055
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   8916
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "Data"
         Caption         =   "Data"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "CodigoProduto"
         Caption         =   "Codigo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Nome"
         Caption         =   "Nome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Unidade"
         Caption         =   "Un"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "QTEntradaSanta"
         Caption         =   "Santa - E"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "QTSaidaSanta1"
         Caption         =   "Santa 1 - S"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "QTSaidaCalifornia"
         Caption         =   "California - S"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "SaldoGeral"
         Caption         =   "Saldo Geral"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Codigo"
         Caption         =   "Codigo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4649,953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   555,024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column08 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdLancar 
      Caption         =   "Lançar"
      Height          =   375
      Left            =   11040
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Unidade 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Valor 
      Height          =   285
      Left            =   9960
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox Nome 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   960
      Width           =   5415
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fl_Data 
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade"
      Height          =   195
      Index           =   2
      Left            =   8880
      TabIndex        =   20
      Top             =   720
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
      Height          =   195
      Index           =   6
      Left            =   3120
      TabIndex        =   17
      Top             =   0
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      Height          =   195
      Index           =   5
      Left            =   1920
      TabIndex        =   16
      Top             =   0
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filtro"
      Height          =   195
      Left            =   1320
      TabIndex        =   14
      Top             =   120
      Width           =   330
   End
   Begin VB.Line Line2 
      X1              =   1200
      X2              =   1200
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   11880
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unidade"
      Height          =   195
      Index           =   4
      Left            =   6600
      TabIndex        =   12
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Unitario"
      Height          =   195
      Index           =   3
      Left            =   9960
      TabIndex        =   10
      Top             =   720
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Galpao de Origem"
      Height          =   195
      Index           =   1
      Left            =   7320
      TabIndex        =   9
      Top             =   720
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   345
   End
End
Attribute VB_Name = "Transferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private GlBase      As String
Private StrGeral    As String

Private Sub CmdExcluir_Click()
On Error Resume Next
Dim StrSql As String
Dim CodigoExcluir As Long
DataGrid1.Col = 8

CodigoExcluir = CLng(DataGrid1.Text)
If CodigoExcluir = 0 Then Exit Sub

If MsgBox("Confirma a Exclusão do Registro?", vbYesNo, "Confirmação") = vbNo Then
   Codigo.SetFocus
   Exit Sub
End If

StrSql = "Delete from Estoquefiscal where codigo=" & CodigoExcluir

conexaoAdo.Execute StrSql, afetados
AtualizaGrid
Codigo.SetFocus

End Sub

Private Sub CmdFiltrar_Click()
On Error GoTo errfiltro
Dim RsBalanco As ADODB.Recordset
Dim StrSql As String
Dim StrWhere As String
LcCap = Me.Caption
Me.Caption = "Aguarde, efetuando filtro..."
Screen.MousePointer = 11
StrGeral = "Select * from estoquefiscal "
If IsDate(Fl_Data.Text) Then
   StrWhere = "where data='" & Format(Fl_Data.Text, "yyyy-mm-dd") & "'"
End If
If Len(Fl_Produto.Text) > 0 Then
   If Len(StrWhere) > 0 Then
      StrWhere = StrWhere & " and nome like '" & UCase(Fl_Produto.Text) & "%'"
   Else
      StrWhere = "where nome like '" & UCase(Fl_Produto.Text) & "%'"
   End If
End If
StrGeral = StrGeral & StrWhere & " Order by nome"
Set RsBalanco = AbreRecordset(StrGeral)
Set DataGrid1.DataSource = RsBalanco

Me.Caption = LcCap
Screen.MousePointer = 0

Exit Sub
errfiltro:
Me.Caption = LcCap
Screen.MousePointer = 0
MsgBox Err.Description & Err.Number
End Sub

Private Sub CmdLancar_Click()
On Error GoTo ErroLanca
Dim RsBalanco   As ADODB.Recordset
Dim StrSql      As String
Dim Saldo       As Double
Dim TemRegistro As Boolean
Dim QSanta      As Double
Dim QSanta1     As Double
Dim QCalifornia As Double
Dim SaldoT      As Double
Dim CustoT      As Double

'==> Zera as quantidades
If Len(Quantidade.Text) = 0 Then Quantidade.Text = 0
If Len(Valor.Text) = 0 Then Valor.Text = 0
'==> As validações
'==> Primeiro a data
If Not IsDate(Data.Text) Then
   MsgBox "É nescessário informar a data de lançamento.", 64, "Aviso"
   Data.SetFocus
   Exit Sub
End If
'==> O Produto
If Len(Codigo.Text) = 0 Then
   MsgBox "É nescessário informar o produto.", 64, "Aviso"
   Codigo.SetFocus
   Exit Sub
End If
'==> Verifica a escolha do Galpão
If Len(Galpao.Text) = 0 Then
   MsgBox "É nescessário informar o galpão.", 64, "Aviso"
   Galpao.SetFocus
   Exit Sub
End If
If UCase(Galpao.Text) <> "SANTA MARIA 1" And UCase(Galpao.Text) <> "CALIFORNIA" Then
   MsgBox "O Galpão informado é inválido.", 64, "Aviso"
   Galpao.SetFocus
   Exit Sub

End If
'==> agora as quantidades
If CDbl(Quantidade.Text) = 0 Then
   MsgBox "É nescessário informar a quantidade.", 64, "Aviso"
   Santa.SetFocus
   Exit Sub
End If
'==> agora o Valor
If CDbl(Valor.Text) = 0 Then
   MsgBox "É nescessário informar o valor.", 64, "Aviso"
   Valor.SetFocus
   Exit Sub
End If

TemRegistro = False


'SaldoT = CDbl(Santa.Text) + CDbl(Santa1.Text) + CDbl(california.Text)

CustoT = CDbl(Valor.Text) * CDbl(Quantidade.Text)
'==> Cria a Sql
SaldoT = 0
 '==> Vamos Inserir a saida do galpao
 StrSql = "Insert into estoqueFiscal(data,CodigoProduto,Unidade,quantidadeEntradageral" & _
        ",valorcustomediounitario,vcustototal,saldoGeral,QuantidadeSaidaGeral,QTEntradaSanta" & _
        ",QTEntradaSanta1,QTEntradaCalifornia,QTSaidaSanta,QTSaidaSanta1,QTSaidaCalifornia" & _
        ",Saldosanta,SaldoSanta1,SaldoCalifornia,Nome) Values ('" & _
        Format(Data.Text, "yyyy-mm-dd") & "'," & _
        Codigo.Text & ",'" & _
        Unidade.Text & "'," & _
        Replace(SaldoT, ",", ".") & "," & _
        Replace(CDbl(Valor.Text), ",", ".") & "," & _
        Replace(CustoT, ",", ".") & "," & _
        "0," & _
        "0," & _
        Replace(CDbl(Quantidade.Text), ",", ".") & "," & _
        "0," & _
        "0," & _
        "0," & _
        IIf(UCase(Galpao.Text) = "SANTA MARIA 1", Replace(Quantidade.Text, ",", "."), 0) & "," & _
        IIf(UCase(Galpao.Text) = "CALIFORNIA", Replace(Quantidade.Text, ",", "."), 0) & "," & _
        "0," & _
        "0," & _
        "0,'" & _
        Replace(Replace(Replace(Nome.Text, ",", ""), "'", ""), Chr(34), "") & "')"

'==Executa a alteração na tabela
conexaoAdo.Execute StrSql, afetados

 '==> Vamos Inserir a Entrada no Santa Maria
' StrSql = "Insert into estoqueFiscal(data,CodigoProduto,Unidade,quantidadeEntradageral" & _
        ",valorcustomediounitario,vcustototal,saldoGeral,QuantidadeSaidaGeral,QTEntradaSanta" & _
        ",QTEntradaSanta1,QTEntradaCalifornia,QTSaidaSanta,QTSaidaSanta1,QTSaidaCalifornia" & _
        ",Saldosanta,SaldoSanta1,SaldoCalifornia,Nome) Values ('" & _
        Format(Data.Text, "yyyy-mm-dd") & "'," & _
        Codigo.Text & ",'" & _
        Unidade.Text & "'," & _
        Replace(SaldoT, ",", ".") & "," & _
        Replace(CDbl(Valor.Text), ",", ".") & "," & _
        Replace(CustoT, ",", ".") & "," & _
        "0," & _
        "0," & _
        Replace(CDbl(Quantidade.Text), ",", ".") & "," & _
        "0," & _
        "0," & _
        "0," & _
        "0," & _
        "0," & _
        "0," & _
        "0," & _
        "0,'" & _
        Replace(Replace(Replace(Nome.Text, ",", ""), "'", ""), Chr(34), "") & "')"

'==Executa a alteração na tabela
'conexaoAdo.Execute StrSql, afetados

'==> Limpa os campos
Codigo.Text = ""
Nome.Text = ""
Unidade.Text = ""
Quantidade.Text = ""
Valor.Text = ""
Codigo.SetFocus
AtualizaGrid

Exit Sub
ErroLanca:
MsgBox "Ocorreu o Seguinte erro lançando o item:" & Chr(13) & Err.Description & " Nº:" & Err.Number, 64, "Aviso"
'Resume 0

End Sub

Private Sub Codigo_GotFocus()
On Error Resume Next
CmdExcluir.Enabled = False

End Sub

Private Sub Codigo_LostFocus()
BuscaNome
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
CmdExcluir.Enabled = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44

End Sub

Private Sub Form_Load()
CarregaCombo
GeraGrid
End Sub
Sub CarregaCombo()
On Error Resume Next
Dim StrSql As String
Dim RsProduto As ADODB.Recordset
Dim LcCap As String
Dim LcTotalreg As Long
Dim a As Long

LcCap = Me.Caption

Screen.MousePointer = 11

Me.Caption = "Aguarde, carregando a lista de produtos..."

StrSql = "Select * from Produtos order by nome"
Set RsProduto = AbreRecordset(StrSql, True)
Nome.Clear
If Not RsProduto.EOF Then
   RsProduto.MoveLast
   LcTotalreg = RsProduto.RecordCount
   RsProduto.MoveFirst
End If
a = 0

Do Until RsProduto.EOF
  a = a + 1
  Me.Caption = "Carregando produto " & RsProduto!Nome & " Reg.:" & a & " de " & LcTotalreg
  DoEvents
  If Err.Number <> 0 Then Exit Do
  Nome.AddItem RsProduto!Nome & ""
  RsProduto.MoveNext
Loop

Set RsProduto = Nothing

Me.Caption = LcCap
Screen.MousePointer = 0

End Sub
Sub GeraGrid()
On Error GoTo errcarregando
Dim RsBalanco As ADODB.Recordset
Dim StrSql As String
LcCap = Me.Caption

Screen.MousePointer = 11

Me.Caption = "Aguarde, carregando o Grid ..."
DoEvents
StrGeral = "Select * from estoquefiscal order by nome"
Set RsBalanco = AbreRecordset(StrGeral)
Set DataGrid1.DataSource = RsBalanco

saida:
Me.Caption = LcCap
Screen.MousePointer = 0

Exit Sub
errcarregando:
GoTo saida
MsgBox Err.Description & Err.Number

End Sub
Sub BuscaNome()
On Error Resume Next
Dim RsProduto As ADODB.Recordset
Dim StrSql As String
Dim Db As Database
Dim RsUnidade As Recordset
If Len(Codigo.Text) = 0 Then Exit Sub
StrSql = "Select * from Produtos where Codigo=" & Codigo.Text

Set RsProduto = AbreRecordset(StrSql, True)

If Not RsProduto.EOF Then
   Nome.Text = RsProduto!Nome & ""
   'Custo.Text = RsProduto!Custo & ""
   BuscaGlbase
   Set Db = OpenDatabase(GlBase)
   Set RsUnidade = Db.OpenRecordset("Select * from alid004 where cod='" & RsProduto!UnidMedida & "'")
   If Not RsUnidade.EOF Then
      Unidade.Text = RsUnidade!Simbolo
   End If
   Set RsUnidade = Nothing
   Galpao.SetFocus
Else
   Nome.Text = ""
   Unidade.Text = ""
   'Custo.Text = ""
   Codigo.Text = ""
   'Santa.Text = ""
   'Santa1.Text = ""
   'california.Text = ""

   Codigo.SetFocus
   MsgBox "Codigo não encontrado.", 64, "Aviso"
End If
End Sub
Sub BuscaCodigo()
'On Error Resume Next
Dim RsProduto As ADODB.Recordset
Dim StrSql As String
Dim Db As Database
Dim RsUnidade As Recordset
If Len(Nome.Text) = 0 Then Exit Sub
StrSql = "Select * from Produtos where nome='" & Nome.Text & "'"

Set RsProduto = AbreRecordset(StrSql, True)

If Not RsProduto.EOF Then
   Codigo.Text = RsProduto!Codigo
   'Custo.Text = RsProduto!Custo & ""
   BuscaGlbase
   Set Db = OpenDatabase(GlBase)
   codU = Right("00" & CStr(RsProduto!UnidMedida), 2)
   Set RsUnidade = Db.OpenRecordset("Select * from alid004 where cod='" & Right("00" & CStr(RsProduto!UnidMedida), 2) & "'")
   If Not RsUnidade.EOF Then
      Unidade.Text = RsUnidade!Simbolo
   End If
   Set RsUnidade = Nothing
   Galpao.SetFocus
Else
   Nome.Text = ""
   Unidade.Text = ""
  ' Custo.Text = ""
   Codigo.Text = ""
  ' Santa.Text = ""
   'Santa1.Text = ""
   'california.Text = ""
   Codigo.SetFocus
   MsgBox "Nome não encontrado.", 64, "Aviso"
   
End If

End Sub

Sub BuscaGlbase()
Dim NumeroDoArquivo As Integer
NumeroDoArquivo = FreeFile

Open App.Path & "\BaseDados.txt" For Input As #NumeroDoArquivo
        Line Input #NumeroDoArquivo, GlBase
        
Close #NumeroDoArquivo
End Sub

Private Sub Nome_GotFocus()
On Error Resume Next
CmdExcluir.Enabled = False

End Sub

Private Sub Nome_LostFocus()
BuscaCodigo
End Sub
Sub AtualizaGrid()
Dim RsBalanco As ADODB.Recordset
Set RsBalanco = AbreRecordset(StrGeral)
Set DataGrid1.DataSource = RsBalanco


End Sub
