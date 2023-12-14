VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form DetalhaCaixaTm 
   BackColor       =   &H00CAE1A2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalha Caixa"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   Icon            =   "DetalhaCaixaTm.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Fechar F10"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5760
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid item 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9340
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   14286847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para Ver Detalhes, de um Duplo Clique em cima do Item "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "DetalhaCaixaTm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcCrit As String

Private a, b As Integer

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim db As Database, RsMov As Recordset, RsTipo As Recordset
Dim LcTipo As String
Dim LcAchou As Boolean
LcCrit = "Select * from MovimentacaoCaixa where dataLancamento=#" & Format(Caixa.Data.Text, "mm/dd/yy") & "# and Rec_Pag='" & Caixa.Tag & "' order by tipomonetario"
GeraGrid
Set db = OpenDatabase(GLBase, False, False) ' "dBASE III;")
Set RsMov = db.OpenRecordset(LcCrit, dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsTipo = db.OpenRecordset("Alid008", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

If RsMov.EOF Then
   MsgBox "Não existe Lançamentos para esta Opção", 64, "Aviso"
   'Unload Me
   Exit Sub
End If
a = 1

Do Until RsMov.EOF
   LcAchou = False
   Item.Rows = a + 1
   LcTipo = "TPMONET='" & RsMov!TipoMonetario & "'"
   RsTipo.FindFirst LcTipo
   If Not RsTipo.NoMatch Then
      LcMone = RsTipo!XTPMONET & ""
   Else
      LcMone = ""
   End If
   ' ===Procura Para Ver se Já existe o tipo na lista
   For x = 1 To Item.Rows - 1
       If Item.TextMatrix(x, 0) = LcMone Then
          b = x
          LcAchou = True
          Exit For
       End If
   Next
   If Not LcAchou Then
      Item.TextMatrix(a, 0) = LcMone
      Item.TextMatrix(a, 1) = RsMov!DataLancamento & ""
      Item.TextMatrix(a, 2) = AcertaNumero(CStr(RsMov!valor), 2) & ""
      a = a + 1
   Else
      Item.TextMatrix(b, 2) = AcertaNumero(CStr(RsMov!valor + CDbl(Item.TextMatrix(b, 2))), 2) & ""
   End If
   RsMov.MoveNext
   
 Loop
 RsMov.Close
 RsTipo.Close
 db.Close
 
End Sub
Function GeraGrid()
On Error Resume Next
Item.ColAlignment(0) = 1
Item.ColAlignment(1) = 4
Item.ColAlignment(2) = 7
'item.ColAlignment(3) = 1

Item.ColWidth(0) = 3000
Item.ColWidth(1) = 2500
Item.ColWidth(2) = 2000
'item.ColWidth(3) = 3800


Item.TextMatrix(0, 0) = "Tipo Monetário"
Item.TextMatrix(0, 1) = "Data"
Item.TextMatrix(0, 2) = "Valor"
'item.TextMatrix(0, 3) = "Endereço"

End Function

Private Sub Form_Unload(Cancel As Integer)
Caixa.SetFocus
End Sub

Private Sub Item_DblClick()
On Error Resume Next
Me.Tag = Item.TextMatrix(Item.Row, 0)
DetalhaporTipo.Show , Me
End Sub

Private Sub item_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 121 Then SendKeys "%+{F}"
End Sub
