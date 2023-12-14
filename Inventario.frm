VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Inventario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario de estoque"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   720
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton CmdGerar 
      Caption         =   "Gerar"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Dataf 
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
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
      TabIndex        =   2
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "Inventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFechar_Click()
On Error Resume Next
Unload Me

End Sub
Sub GeraRel()
On Error GoTo ErrGera
Dim StrSql As String
Dim Rs As ADODB.Recordset
Dim db As Database
Dim Saldo As Double
Dim SaldoAnterior As Double
Dim CodProduto As Long


Set db = OpenDatabase(GLBase)
db.Execute "Delete from RelIventario"
StrSql = "Select * from estoquefiscal "
If IsDate(DataI.Text) And IsDate(DataF.Text) Then
    StrSql = StrSql & " where (data Between #" & Format(CDate(DataI.Text), "mm/dd/yy") & "# And #" & Format(CDate(DataF.Text), "mm/dd/yy") & "#)"
End If
StrSql = StrSql & " order by codigoproduto"
Set Rs = AbreRecordset(StrSql, True)
CodProduto = 0
If Not Rs.EOF Then
   Rs.MoveLast
   TotalReg = Rs.RecordCount
   Rs.MoveFirst
End If
x = 0
Do Until Rs.EOF
   x = x + 1
   Me.Caption = "Reg " & x & " de " & TotalReg
   DoEvents
   '==> Busca o saldo asnterior
   If CodProduto <> CLng(Rs!codigoproduto) Then
      CodProduto = Rs!codigoproduto
      SaldoAnterior = BuscarSaldoAnterior(CodProduto)
      Saldo = BuscarSaldoUltimo(CodProduto)
   End If
   StrSql = "insert into RelIventario (CodigoProduto,nome,quantidade,valorcustomediounitario,vcustototal,saldo,quantidadeSaida,SaldoAnterior) values (" & _
         Rs!codigoproduto & ",'" & _
         Rs!Nome & "'," & _
         Replace(Rs!Quantidade, ",", ".") & "," & _
         Replace(Rs!valorcustomediounitario, ",", ".") & "," & _
         Replace(Rs!vcustototal, ",", ".") & "," & _
         Replace(Saldo, ",", ".") & "," & _
         Replace(Rs!quantidadeSaida, ",", ".") & "," & _
         Replace(SaldoAnterior, ",", ".") & ")"
   db.Execute StrSql
   Rs.MoveNext
Loop

Exit Sub
ErrGera:
MsgBox err.Description & err.Number
'Resume 0
End Sub

Private Sub CmdGerar_Click()
On Error Resume Next
LcCap = Me.Caption
Me.Caption = "Aguarde..."
Screen.MousePointer = 11
GeraRel
Cryrelatorio.DataFiles(0) = GLBase
Cryrelatorio.ReportFileName = App.Path & "\iventario.rpt"

Cryrelatorio.DiscardSavedData = True
Cryrelatorio.WindowTop = 50
Cryrelatorio.WindowWidth = 700
Cryrelatorio.WindowLeft = 50
Cryrelatorio.WindowHeight = 500
Cryrelatorio.WindowTitle = "Iventario de estoque."

Cryrelatorio.Formulas(0) = "Titulo=' Inventario de estoque Periodo de " & DataI.Text & " a " & DataF.Text & "'"
LcTipoSaida = 0
'CryRelatorio.SortFields(0) = "+{RelsaidaContabil.nf}"
Cryrelatorio.Destination = LcTipoSaida
Cryrelatorio.PrintReport
Me.Caption = LcCap
Screen.MousePointer = 0

If Cryrelatorio.LastErrorNumber > 0 Then MsgBox Cryrelatorio.LastErrorString
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"

End Sub

