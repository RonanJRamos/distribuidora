VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form DetalhaporTipo 
   BackColor       =   &H00B3E9FD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalhes do Tipo"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11340
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Fechar  F10"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   6120
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid item 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9975
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   12713983
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "DetalhaporTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private a As Integer
Private Sub Command1_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim db As Database, RsTipo As Recordset, RsConta As Recordset, RsCli As Recordset
Dim RsRe As Recordset
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

GeraGrid

LcCrit = "Select * from alid008 where XTPMONET='" & DetalhaCaixaTm.Tag & "'"
Set db = OpenDatabase(GLBase, False, False) ' "dBASE III;")
Set RsTipo = db.OpenRecordset(LcCrit, dbOpenDynaset, dbSeeChanges, dbOptimistic) '"Alid008", dbOpenDynaset, dbSeeChanges, dbOptimistic)
If GlRec = "R" Then
   Set RsCli = db.OpenRecordset("Alid001", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   Set RsRe = db.OpenRecordset("Alid015", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   LcCriNota = "Select * from MovimentacaoCaixa where DataLancamento=#" & Format(Caixa.Data.Text, "mm/dd/yy") & "# and tipomonetario='" & RsTipo!TPMONET & "' and Rec_pag='R'"
Else
   Set RsRe = db.OpenRecordset("Alid014", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   Set RsCli = db.OpenRecordset("Alid002", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   LcCriNota = "Select * from MovimentacaoCaixa where DataLancamento=#" & Format(Caixa.Data.Text, "mm/dd/yy") & "# and tipomonetario='" & RsTipo!TPMONET & "' and Rec_pag='D'"
End If
'MsgBox LcCriNota
Set RsConta = db.OpenRecordset(LcCriNota, dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset, dbSeeChanges, dbOptimistic)

If RsConta.EOF Then
   MsgBox "Não Foi Encontrado Lançamnetos com Este Tipo Monetario.", 64, "Aviso"
   Exit Sub
End If
a = 1
Do Until RsConta.EOF
   Item.Rows = a + 1
   If GlRec = "R" Then
      LcC = "Nf='" & RsConta!NF & "'"
      RsRe.FindFirst LcC
      If Not RsRe.NoMatch Then
        LCCriCli = "codigo='" & RsRe!Cliente & "'"
      End If
   Else
      LcC = "Nf='" & RsConta!NF & "'"
      RsRe.FindFirst LcC
      If Not RsRe.NoMatch Then
         LCCriCli = "codigo='" & RsConta!credor & "'"
      End If
   End If
   RsCli.FindFirst LCCriCli
   If Not RsCli.NoMatch Then
      LcNome = RsCli!RAZAOSOC & ""
   End If
   Item.TextMatrix(a, 0) = RsConta!NF & ""
   Item.TextMatrix(a, 1) = LcNome & ""
   Item.TextMatrix(a, 2) = RsConta!DataLancamento
   Item.TextMatrix(a, 3) = AcertaNumero(CStr(RsConta!valor), 2)
   RsConta.MoveNext
   a = a + 1
Loop
RsConta.Close
RsTipo.Close
RsCli.Close
db.Close

End Sub
Function GeraGrid()
On Error Resume Next
Item.ColAlignment(0) = 1
Item.ColAlignment(1) = 1
Item.ColAlignment(2) = 4
Item.ColAlignment(3) = 7

Item.ColWidth(0) = 1500
Item.ColWidth(1) = 5500
Item.ColWidth(2) = 1500
Item.ColWidth(3) = 1500


Item.TextMatrix(0, 0) = "Documento"
Item.TextMatrix(0, 1) = "Cliente"
Item.TextMatrix(0, 2) = "Data Pag."
Item.TextMatrix(0, 3) = "Valor Pag."

End Function

Private Sub Form_Unload(Cancel As Integer)
DetalhaCaixaTm.SetFocus
End Sub

Private Sub item_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then SendKeys "%{F}"
End Sub
