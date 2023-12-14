VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form BalancoGalpao 
   Caption         =   "Entrada Balanço Galpão"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   11040
      TabIndex        =   22
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Custo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   3120
      TabIndex        =   21
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton CmdFiltrar 
      Caption         =   "Filtrar"
      Height          =   375
      Left            =   11040
      TabIndex        =   20
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Fl_Produto 
      Height          =   285
      Left            =   3120
      TabIndex        =   19
      Top             =   240
      Width           =   7815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4335
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   7646
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
         Caption         =   "Santa"
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
         DataField       =   "QTEntradaSanta1"
         Caption         =   "Santa 1"
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
         DataField       =   "QTEntradaCalifornia"
         Caption         =   "California"
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
            ColumnWidth     =   3809,764
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
      Top             =   870
      Width           =   735
   End
   Begin VB.TextBox Unidade 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox california 
      Height          =   285
      Left            =   9840
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Santa1 
      Height          =   285
      Left            =   8760
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Santa 
      Height          =   285
      Left            =   7680
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
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
      Width           =   5775
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
      TabIndex        =   16
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
      Caption         =   "Produto"
      Height          =   195
      Index           =   6
      Left            =   3120
      TabIndex        =   18
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
      TabIndex        =   17
      Top             =   0
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filtro"
      Height          =   195
      Left            =   1320
      TabIndex        =   15
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
      Left            =   6960
      TabIndex        =   13
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "California"
      Height          =   195
      Index           =   3
      Left            =   9840
      TabIndex        =   11
      Top             =   720
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Santa Maria 1"
      Height          =   195
      Index           =   2
      Left            =   8760
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Santa Maria"
      Height          =   195
      Index           =   1
      Left            =   7680
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   855
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
Attribute VB_Name = "BalancoGalpao"
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
If Len(Santa.Text) = 0 Then Santa.Text = 0
If Len(Santa1.Text) = 0 Then Santa1.Text = 0
If Len(california.Text) = 0 Then california.Text = 0

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
'==> agora as quantidades
If CDbl(Santa.Text) + CDbl(Santa1.Text) + CDbl(california.Text) = 0 Then
   MsgBox "É nescessário informar a quantidade em pelo menos um galpão.", 64, "Aviso"
   Santa.SetFocus
   Exit Sub
End If

'==> Vamos ver se já existe um registro para este produto nesta data.
TemRegistro = False
StrSql = "Select * from estoquefiscal where codigoproduto=" & Codigo.Text & " and data='" & Format(Data.Text, "yyyy-mm-dd") & "'"

Set RsBalanco = AbreRecordset(StrSql, True)

If Not RsBalanco.EOF Then
   TemRegistro = True
   QSanta = RsBalanco!saldoSanta
   QSanta1 = RsBalanco!saldoSanta1
   QCalifornia = RsBalanco!saldocalifornia
   SaldoT = (QSanta + QSanta1 + QCalifornia) - (RsBalanco!QTSaidaSanta + RsBalanco!QTSaidaSanta1 + RsBalanco!QTSaidaCalifornia) + (CDbl(Santa.Text) + CDbl(Santa1.Text) + CDbl(california.Text))
Else
   TemRegistro = False
   QSanta = 0
   QSanta1 = 0
   QCalifornia = 0
   SaldoT = CDbl(Santa.Text) + CDbl(Santa1.Text) + CDbl(california.Text)

End If

Set RsBalanco = Nothing
CustoT = CDbl(Custo.Text) * SaldoT
'==> Cria a Sql
If TemRegistro Then
   '==> Vamos Atualizar
   StrSql = "Update estoqueFiscal Set " & _
            "QTEntradaSanta=" & Replace(CDbl(Santa.Text) + QSanta, ",", ".") & _
            ",QTEntradaSanta1=" & Replace(CDbl(Santa1.Text) + QSanta1, ",", ".") & _
            ",QTEntradacalifornia=" & Replace(CDbl(california.Text) + QCalifornia, ",", ".") & _
            ",quantidadeEntradageral=" & Replace(SaldoT, ",", ".") & _
            ",valorcustomediounitario=" & Replace(CDbl(Custo.Text), ",", ".") & _
            ",vcustototal=" & Replace(CustoT, ",", ".") & _
            ",saldoGeral=" & Replace(SaldoT, ",", ".") & _
            ",QuantidadeSaidaGeral=0" & _
            ",QTSaidaSanta=0" & _
            ",QTSaidaSanta1=0" & _
            ",QTSaidaCalifornia=0" & _
            ",Saldosanta=" & Replace(CDbl(Santa.Text) + QSanta, ",", ".") & _
            ",Saldosanta1=" & Replace(CDbl(Santa1.Text) + QSanta1, ",", ".") & _
            ",SaldoCalifornia=" & Replace(CDbl(california.Text) + QCalifornia, ",", ".") & _
            " Where Codigoproduto=" & Codigo.Text & " and Data='" & Format(Data.Text, "yyyy-mm-dd") & "'"

Else
 '==> Vamos Inserir
 StrSql = "Insert into estoqueFiscal(data,CodigoProduto,Unidade,quantidadeEntradageral" & _
        ",valorcustomediounitario,vcustototal,saldoGeral,QuantidadeSaidaGeral,QTEntradaSanta" & _
        ",QTEntradaSanta1,QTEntradaCalifornia,QTSaidaSanta,QTSaidaSanta1,QTSaidaCalifornia" & _
        ",Saldosanta,SaldoSanta1,SaldoCalifornia,Nome) Values ('" & _
        Format(Data.Text, "yyyy-mm-dd") & "'," & _
        Codigo.Text & ",'" & _
        Unidade.Text & "'," & _
        Replace(SaldoT, ",", ".") & "," & _
        Replace(CDbl(Custo.Text), ",", ".") & "," & _
        Replace(CustoT, ",", ".") & "," & _
        Replace(SaldoT, ",", ".") & "," & _
        "0," & _
        Replace(CDbl(Santa.Text), ",", ".") & "," & _
        Replace(CDbl(Santa1.Text), ",", ".") & "," & _
        Replace(CDbl(california.Text), ",", ".") & "," & _
        "0," & _
        "0," & _
        "0," & _
        Replace(CDbl(Santa.Text), ",", ".") & "," & _
        Replace(CDbl(Santa1.Text), ",", ".") & "," & _
        Replace(CDbl(california.Text), ",", ".") & ",'" & _
        Replace(Replace(Replace(Nome.Text, ",", ""), "'", ""), Chr(34), "") & "')"
End If

'==Executa a alteração na tabela
conexaoAdo.Execute StrSql, afetados

'==> Limpa os campos
Codigo.Text = ""
Nome.Text = ""
Unidade.Text = ""
Santa.Text = ""
Santa1.Text = ""
california.Text = ""
Custo.Text = ""
Codigo.SetFocus
AtualizaGrid

Exit Sub
ErroLanca:
MsgBox "Ocorreu o Seguinte erro lançando o item:" & Chr(13) & Err.Description & " Nº:" & Err.Number, 64, "Aviso"


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
If KeyCode = 13 Then
   SendKeys "{tab}"
   SendKeys "{home}+{end}"
End If
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
Dim RsEstoque As ADODB.Recordset
Dim StrSql As String
Dim Db As Database
Dim RsUnidade As Recordset
If Len(Codigo.Text) = 0 Then Exit Sub
StrSql = "Select * from Produtos where Codigo=" & Codigo.Text

Set RsProduto = AbreRecordset(StrSql, True)

If Not RsProduto.EOF Then
   StrSql = "Select * from estoquefiscal where Codigoproduto=" & Codigo.Text & " and data='" & Format(Data.Text, "yyyy-mm-dd") & "'"
   Debug.Print StrSql
   Set RsEstoque = AbreRecordset(StrSql, True)
   If Not RsEstoque.EOF Then
      Santa.Text = RsEstoque!saldoSanta
      Santa1.Text = RsEstoque!saldoSanta1
      california.Text = RsEstoque!saldocalifornia
   Else
      Santa.Text = 0
      Santa1.Text = 0
      california.Text = 0
   End If
   
   Nome.Text = RsProduto!Nome & ""
   Custo.Text = RsProduto!Custo & ""
   BuscaGlbase
   Set Db = OpenDatabase(GlBase)
   Set RsUnidade = Db.OpenRecordset("Select * from alid004 where cod='" & RsProduto!UnidMedida & "'")
   If Not RsUnidade.EOF Then
      Unidade.Text = RsUnidade!Simbolo
   End If
   Set RsUnidade = Nothing
   Santa.SetFocus
Else
   Nome.Text = ""
   Unidade.Text = ""
   Custo.Text = ""
   Codigo.Text = ""
   Santa.Text = ""
   Santa1.Text = ""
   california.Text = ""

   Codigo.SetFocus
   MsgBox "Codigo não encontrado.", 64, "Aviso"
End If
End Sub
Sub BuscaCodigo()
On Error Resume Next
Dim RsProduto As ADODB.Recordset
Dim StrSql As String
Dim Db As Database
Dim RsUnidade As Recordset
If Len(Nome.Text) = 0 Then Exit Sub
StrSql = "Select * from Produtos where nome='" & Nome.Text & "'"

Set RsProduto = AbreRecordset(StrSql, True)

If Not RsProduto.EOF Then
   Codigo.Text = RsProduto!Codigo
   Custo.Text = RsProduto!Custo & ""
   BuscaGlbase
   Set Db = OpenDatabase(GlBase)
   codU = Right("00" & CStr(RsProduto!UnidMedida), 2)
   Set RsUnidade = Db.OpenRecordset("Select * from alid004 where cod='" & Right("00" & CStr(RsProduto!UnidMedida), 2) & "'")
   If Not RsUnidade.EOF Then
      Unidade.Text = RsUnidade!Simbolo
   End If
   Set RsUnidade = Nothing
   Santa.SetFocus
Else
   Nome.Text = ""
   Unidade.Text = ""
   Custo.Text = ""
   Codigo.Text = ""
   Santa.Text = ""
   Santa1.Text = ""
   california.Text = ""
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
