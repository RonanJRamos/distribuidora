VERSION 5.00
Begin VB.Form FrmAcertaLimiteCliente 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Acerta Limite de Credito dos Clientes"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAcerta 
      Caption         =   "Acertar Credito"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4455
   End
End
Attribute VB_Name = "FrmAcertaLimiteCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAcerta_Click()
Dim rsCliente As Recordset
Dim RsReceita As ADODB.Recordset
Dim Total As Long
Dim i As Long
Dim StrSql As String
AbreBase
Set rsCliente = Dbbase.OpenRecordset("ALID001", dbOpenDynaset, dbSeeChanges, dbOptimistic)
If Not rsCliente.EOF Then
   rsCliente.MoveLast
   Total = rsCliente.RecordCount
   rsCliente.MoveFirst
End If
Do Until rsCliente.EOF
    i = i + 1
    Label1.Caption = "Acertando CLiente:" & rsCliente!RazaoSoc & " - " & i & " de " & Total
    DoEvents
    Dim LcValor As Currency
    StrSql = "select sum(valor) as Valor_Total from alid015 where VALPAGO =0 and cliente='" & rsCliente!Codigo & "'"
    Set RsReceita = AbreRecordset(StrSql, True)
    If Not RsReceita.EOF Then
       If IsNumeric(RsReceita!Valor_Total) Then
          LcValor = RsReceita!Valor_Total
       Else
          LcValor = 0
       End If
       StrSql = "Update alid001 set CreditoUtilizado=" & Replace(CStr(LcValor), ",", ".") & " where codigo='" & rsCliente!Codigo & "'"
       Dbbase.Execute StrSql
       
    End If
    rsCliente.MoveNext
Loop
MsgBox "Alteração Finalizada", 64, "Aviso"
End Sub
