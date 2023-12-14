VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form comissao 
   BackColor       =   &H00B3E9FD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comissão do Vendedor"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Excluir F3"
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Fechar 
      Caption         =   "&Fechar  F10"
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton Confirma 
      Caption         =   "&Confirma   F2"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Item 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5106
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
   End
   Begin MSMask.MaskEdBox valor 
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "comissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type MsComissao
        Descricao As String
        valor As Currency
End Type
Private LcMat() As MsComissao
Private LcTam As Long
Private LcAtivoComissao, a As Integer


Private Sub Command1_Click()
On Error Resume Next
FrmExcluiItem.Show , Me
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 114 Then SendKeys "%+{E}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Confirma_Click()
On Error Resume Next
Dim RsComissao As Recordset
Dim b As Integer
AbreBase
LcSql = "Select * From comissaovendedor where vendedor='" & FrmFuncionario.Txt(0).Text & "'"
Set RsComissao = Dbbase.OpenRecordset(LcSql)

Do Until RsComissao.EOF
   RsComissao.Delete
   RsComissao.MoveNext
Loop
For a = 1 To Item.Rows - 1
 
    RsComissao.AddNew
    RsComissao!Vendedor = FrmFuncionario.Txt(0).Text
    RsComissao!Comissao = CDbl(Item.TextMatrix(a, 0))
    RsComissao!Descricao = Item.TextMatrix(a, 1)
    RsComissao.Update

  If err.Number = 9 Then Exit For
Next
RsComissao.Close
Dbbase.Close
Set RsComissao = Nothing
Set Dbbase = Nothing
Unload Me
End Sub

Private Sub Confirma_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 114 Then SendKeys "%+{E}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Fechar_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Fechar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 114 Then SendKeys "%+{E}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Activate()
If Not LcAtivoComissao Then
   Inicializacom
   LcAtivoComissao = True
End If
Set GlFormA = Me
End Sub

Private Sub Form_Load()
GeraGrid
LcTam = 0
End Sub
Function ExcluiItem(LcNItem As Currency)
Dim a, b, LcAchou As Integer
On Error Resume Next


For a = Item.Rows - 1 To 1 Step -1
    If Item.TextMatrix(a, 0) = LcNItem Then
        LcAchou = True
        If Item.Rows = 2 Then
           Item.Rows = 1
        Else
          Item.RemoveItem (a)
        End If
       Exit For
    End If
    If err.Number = 9 Then Exit For
Next
If Not LcAchou Then MsgBox "Item Não Encontrado...", vbInformation, "Aviso"
Txt(0).Text = ""
valor.Text = ""


Txt(0).SetFocus

End Function
Function GeraGrid()
Item.ColAlignment(0) = 7
Item.ColAlignment(1) = 3

Item.ColWidth(0) = 500
Item.ColWidth(1) = 3000


Item.TextMatrix(0, 0) = "Valor"
Item.TextMatrix(0, 1) = "Descrição"


LcTamanhoGrid = 1
End Function
Function montagrid()
Dim LcAchou As Integer
On Error Resume Next
'==== Verifica se Foi digitados todos os campos
If Len(Trim(Txt(0).Text)) = 0 Then
   MsgBox "Necessário Informar a Descrição.", 48, "Aviso"
   Txt(0).SetFocus
   Exit Function
End If
If Len(valor.Text) = 0 Then
   MsgBox "Necessário Informar o Valor da Comissão.", 48, "Aviso"
   valor.SetFocus
   Exit Function
End If
Dim b As Integer

If LcTam = 0 Then LcTam = Item.Rows

If LcTam = 2 Then
   b = 1
   LcTam = LcTam + 1
Else
   b = Item.Rows
   LcTam = LcTam + 1
End If


    Item.Rows = b + 1
    Item.TextMatrix(b, 0) = valor.Text
    Item.TextMatrix(b, 1) = Txt(0).Text
    
    

LcTam = LcTam + 1
Txt(0).Text = ""
valor.Text = ""


Txt(0).SetFocus
End Function
Function Inicializacom()
On Error Resume Next
Dim RsComissao As Recordset
Dim b As Integer
AbreBase
LcSql = "Select * From comissaovendedor where vendedor='" & FrmFuncionario.Txt(0).Text & "'"
Set RsComissao = Dbbase.OpenRecordset(LcSql)
b = 1
Do Until RsComissao.EOF
  
  Item.Rows = b + 1
  Item.TextMatrix(b, 0) = RsComissao!Comissao
  Item.TextMatrix(b, 1) = RsComissao!Descricao
  b = b + 1
  RsComissao.MoveNext
 
Loop

Txt(0).Text = ""
valor.Text = ""

RsComissao.Close
Dbbase.Close
Set RsComissao = Nothing
Set Dbbase = Nothing

Txt(0).SetFocus
End Function

Private Sub Form_Unload(Cancel As Integer)
LcAtivoComissao = False
End Sub

Private Sub item_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 114 Then SendKeys "%+{E}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 114 Then SendKeys "%+{E}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub valor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Txt(0).SetFocus
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 114 Then SendKeys "%+{E}"
End Sub

Private Sub valor_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub valor_LostFocus()
If Len(valor.Text) > 0 Then
   If Not IsNumeric(valor.Text) Then
      MsgBox "O Valor da Comissão deve ser Numérico.", vbInformation, "Aviso"
      valor.Text = ""
      valor.SetFocus
      Exit Sub
   End If
End If

montagrid
End Sub
