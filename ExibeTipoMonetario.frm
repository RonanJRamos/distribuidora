VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ExibeMonetario 
   Caption         =   "Exibe Tipo Monetário"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar  F10"
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
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid MostraCliente 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5106
      _Version        =   393216
      FixedCols       =   0
   End
End
Attribute VB_Name = "ExibeMonetario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then Command2_Click

End Sub

Private Sub Form_Load()
montagrid
EscreveGrid
End Sub
Function montagrid()
MostraCliente.ColAlignment(0) = 7
MostraCliente.ColAlignment(1) = 1

MostraCliente.ColWidth(0) = 700
MostraCliente.ColWidth(1) = 3400

MostraCliente.TextMatrix(0, 0) = "Código"
MostraCliente.TextMatrix(0, 1) = "Tipo"
End Function
Function EscreveGrid()
On Error GoTo errList
Dim RsCidade As Recordset
AbreBase
Set RsCidade = DbBase.OpenRecordset("select * from alid008 order by XTPMONET")
a = 2
LcCap = Me.Caption
Me.Caption = "Aguarde, Criando lista de cidades..."
MostraCliente.Rows = 1

Do Until RsCidade.EOF
   MostraCliente.Rows = a
   MostraCliente.TextMatrix(a - 1, 0) = RsCidade!TPMONET
   MostraCliente.TextMatrix(a - 1, 1) = RsCidade!XTPMONET
   a = a + 1
   RsCidade.MoveNext
Loop
Me.Caption = LcCap
RsCidade.Close
Set RsCidade = Nothing
Exit Function
errList:
MsgBox Err.Description & " Nº: " & Err
Exit Function
End Function

Private Sub MostraCliente_DblClick()
a = MostraCliente.Row
LcRegAtual = False
GlFormA.SetFocus
Select Case GlFormA.Name
   Case Is = "Receitas"
       Receitas.txt(5).Text = MostraCliente.TextMatrix(a, 0)
       Receitas.txt(6).Text = MostraCliente.TextMatrix(a, 1)
End Select
Unload Me
End Sub

Private Sub MostraCliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   LcRegAtual = False
   a = MostraCliente.Row
   GlFormA.SetFocus
   Select Case GlFormA.Name

      Case Is = "Receitas"
       Receitas.txt(5).Text = MostraCliente.TextMatrix(a, 0)
       Receitas.txt(6).Text = MostraCliente.TextMatrix(a, 1)
End Select
   Unload Me
End If
If KeyCode = 121 Then Command2_Click
End Sub
