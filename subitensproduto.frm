VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form subitensproduto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sub Itens do Produto"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Codigo 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exclui Item F4"
      Height          =   495
      Left            =   8520
      TabIndex        =   11
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   8520
      TabIndex        =   9
      Top             =   3080
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Salvar F2"
      Height          =   495
      Left            =   8520
      TabIndex        =   8
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F5"
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Quantidade 
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox produto 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   4455
   End
   Begin MSFlexGridLib.MSFlexGrid item 
      Height          =   4215
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Para Listar os Itens, Pressione F5 "
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1560
      TabIndex        =   13
      Top             =   600
      Width           =   2400
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Codigo"
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
      TabIndex        =   12
      Top             =   840
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade"
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
      Left            =   6120
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Item"
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
      Left            =   1560
      TabIndex        =   6
      Top             =   840
      Width           =   450
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   10320
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Principal 
      BackStyle       =   0  'Transparent
      Caption         =   "principal"
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
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto Principal"
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
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "subitensproduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type LcExItem
    codigo As String
    produto As String
    Quantidade As String
End Type

Private Sub Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
KeyCode = 0
End Sub

Private Sub Codigo_LostFocus()
Dim Rs     As Recordset
Dim LcResposta  As Integer
LcSql = "Select * from alid009"
If Len(codigo.Text) = 0 Then Exit Sub
codigo.Text = Right("00000" & codigo.Text, 5)
LcAlterou = True
AbreBase
Set Rs = Dbbase.OpenRecordset(LcSql, dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcPes = "Cod='" & codigo.Text & "'"
Rs.FindFirst LcPes
If Not Rs.NoMatch Then
   produto.Text = Rs!nome & ""
Else
   codigo.Text = ""
End If
Rs.Close
Dbbase.Close
Set Rs = Nothing
Set Dbbase = Nothing

End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim LcValor As Double
Dim a As Integer
Item.Rows = Item.Rows + 1
a = Item.Rows - 1
Item.TextMatrix(a, 0) = Right("00" & a, 2)
Item.TextMatrix(a, 1) = codigo.Text
Item.TextMatrix(a, 2) = produto.Text
Item.TextMatrix(a, 3) = Quantidade.Text

codigo.Text = ""
produto.Text = ""
Quantidade.Text = ""
'==> Totaliza Recibo
'For a = 1 To item.Rows - 1
    'LcValor = LcValor + CDbl(item.TextMatrix(a, 2))
'Next
'txt(0).Text = CStr(CalculanumeroHono)
'codigo.SetFocus
Command2.Enabled = True
Command3.Enabled = True
codigo.SetFocus
'Command1.Enabled = False
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim LcResposta As String
Dim a          As Long
Dim b          As Long
Dim MtTemp()   As LcExItem
Dim LcAchou    As Boolean

LcResposta = InputBox("Entre com o Número do Item a Excluir.")
If Len(LcResposta) = 0 Then
   MsgBox "Não foi Selecionada nenhum Item para a exclusão.", 64, "Aviso"
   Exit Sub
Else
   LcResposta = Right("00" & LcResposta, 2)
End If
b = 0
For a = 1 To Item.Rows - 1
    If LcResposta <> Item.TextMatrix(a, 0) Then
       ReDim Preserve MtTemp(b)
       MtTemp(b).codigo = Item.TextMatrix(a, 1)
       MtTemp(b).produto = Item.TextMatrix(a, 2)
       MtTemp(b).Quantidade = Item.TextMatrix(a, 3)
       b = b + 1
    Else
       LcAchou = True
    End If
Next

If LcAchou Then
   Item.Rows = 1
   b = b - 1
   For a = 0 To b
        Item.Rows = Item.Rows + 1
        c = Item.Rows - 1
        Item.TextMatrix(c, 0) = Right("00" & a + 1, 2)
        Item.TextMatrix(c, 1) = MtTemp(a).codigo & ""
        Item.TextMatrix(c, 2) = MtTemp(a).produto & ""
        Item.TextMatrix(c, 3) = MtTemp(a).Quantidade & ""
   Next
Else
   MsgBox "O Item " & LcResposta & " Não foi Lançado", 64, "Aviso"
End If

End Sub

Private Sub Command3_Click()
'On Error Resume Next
Dim RsDados     As Recordset
Dim LcResposta  As Integer
Dim a As Integer
LcSql = "Select * from DadosProduto where produtoprincipal='" & FrmProduto.Txt(0).Text & "'"

LcAlterou = True
AbreBase
Set RsDados = Dbbase.OpenRecordset(LcSql, dbOpenDynaset, dbSeeChanges, dbOptimistic)

'===> Exclui Os já Cadastrados
Do Until RsDados.EOF
   RsDados.Delete
   RsDados.MoveNext
Loop
'====> Grava os Itens
For a = 1 To Item.Rows - 1
   RsDados.AddNew
   RsDados!Item = Item.TextMatrix(a, 0) & ""
   RsDados!produtoprincipal = FrmProduto.Txt(0).Text & ""
   RsDados!produto = Item.TextMatrix(a, 1) & ""
   RsDados!descricaoprincipal = FrmProduto.Txt(1).Text & ""
   RsDados!Descricao = Item.TextMatrix(a, 2) & "" & ""
   RsDados!Quantidade = CDbl(Item.TextMatrix(a, 3))
   RsDados.Update
Next


RsDados.Close
Dbbase.Close


Set RsDados = Nothing
Set Dbbase = Nothing
'limpa
Unload Me

End Sub

Private Sub Command4_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Activate()
Set GlFormA = Me
Principal.Caption = FrmProduto.Txt(1).Text
End Sub
Function Buscadados()
Dim Rs As Recordset
Dim a As Integer
Dim LcSql As String
LcSql = "select * from dadosproduto where produtoprincipal='" & FrmProduto.Txt(0).Text & "' order by item"
AbreBase
Item.Rows = 1
Set Rs = Dbbase.OpenRecordset(LcSql, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Do Until Rs.EOF
    Item.Rows = Item.Rows + 1
    a = Item.Rows - 1
    Item.TextMatrix(a, 0) = Rs!Item
    Item.TextMatrix(a, 1) = Rs!produto
    Item.TextMatrix(a, 2) = Rs!Descricao
    Item.TextMatrix(a, 3) = Rs!Quantidade
    Rs.MoveNext
Loop
Rs.Close
Dbbase.Close
Set Rs = Nothing
Set dbbae = Nothing
End Function
Private Sub Form_Load()
GeraGrid
Buscadados
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmProduto.SetFocus
End Sub
Function GeraGrid()
On Error Resume Next
Item.ColAlignment(0) = 7
Item.ColAlignment(1) = 3
Item.ColAlignment(2) = 1
Item.ColAlignment(2) = 1


Item.ColWidth(0) = 500
Item.ColWidth(1) = 1000
Item.ColWidth(2) = 4600
Item.ColWidth(3) = 1000

Item.TextMatrix(0, 0) = "Item"
Item.TextMatrix(0, 1) = "Codigo"
Item.TextMatrix(0, 2) = "Produto"
Item.TextMatrix(0, 3) = "Quantidade"

End Function

Private Sub produto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 116 Then
   GlCriterioSql = "Select * from alid009 where nome like '" & produto.Text & "*' order by NOME"
   KeyCode = 0
   Load FrmBuscaProduto
   produto.Text = ""
   FrmBuscaProduto.Tag = produto.Text
   FrmBuscaProduto.Show , Me
End If
End Sub

Private Sub produto_LostFocus()
Dim Rs     As Recordset
Dim LcResposta  As Integer
LcSql = "Select * from alid009"
'If Len(produto.Text) = 0 Then Exit Sub
LcAlterou = True
AbreBase
Set Rs = Dbbase.OpenRecordset(LcSql, dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcPes = "nome='" & produto.Text & "'"
Rs.FindFirst LcPes
If Not Rs.NoMatch Then
   produto.Text = Rs!nome & ""
   codigo.Text = Rs!cod & ""
Else
   GlCriterioSql = "Select * from alid009 where nome like '" & produto.Text & "*' order by NOME"
   Load FrmBuscaProduto
   FrmBuscaProduto.Tag = produto.Text
   FrmBuscaProduto.Show , Me
End If
Rs.Close
Dbbase.Close
Set Rs = Nothing

End Sub

Private Sub Quantidade_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub
