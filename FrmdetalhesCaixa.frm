VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmDetalhesCaixa 
   Caption         =   "Detalhes do Caixa"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10995
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok F2"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   6960
      TabIndex        =   34
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   15
      Left            =   3030
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   14
      Left            =   4485
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   13
      Left            =   5940
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   12
      Left            =   7395
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   11
      Left            =   8850
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   10
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   9
      Left            =   7395
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   8
      Left            =   8850
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   7
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   6
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   5
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   4
      Left            =   3030
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   3
      Left            =   4485
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   2
      Left            =   5940
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   6120
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Itens 
      Height          =   4455
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   7858
      _Version        =   393216
      Cols            =   19
   End
   Begin VB.CommandButton CmdFecha 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   8520
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vendas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   15
      Left            =   1560
      TabIndex        =   33
      Top             =   5760
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Saídas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   14
      Left            =   3120
      TabIndex        =   32
      Top             =   5760
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Acréscimos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   13
      Left            =   4440
      TabIndex        =   31
      Top             =   5760
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descontos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   12
      Left            =   6000
      TabIndex        =   30
      Top             =   5760
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "S. a Dev"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   11
      Left            =   7440
      TabIndex        =   29
      Top             =   5760
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "S. Devolver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   8880
      TabIndex        =   28
      Top             =   5760
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "S. Empr."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   120
      TabIndex        =   27
      Top             =   6600
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F. à Dev"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   1560
      TabIndex        =   26
      Top             =   6600
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F. Dev."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   3120
      TabIndex        =   25
      Top             =   6600
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F. Loc."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   4560
      TabIndex        =   24
      Top             =   6600
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cli. Novos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   6000
      TabIndex        =   23
      Top             =   6600
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cli. Mov"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   7440
      TabIndex        =   22
      Top             =   6600
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Rebob."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   8880
      TabIndex        =   21
      Top             =   6600
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Entradas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   120
      TabIndex        =   20
      Top             =   5760
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Final"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "FrmDetalhesCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcTamanhoGrid, a As Long
Private Sub CmdFecha_Click()
On Error Resume Next
Unload Me

End Sub

Private Sub CmdOk_Click()
MontaPesq
End Sub

Private Sub Form_Load()
On Error Resume Next
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

GeraGrid
End Sub
Function GeraGrid()
Itens.ColAlignment(0) = 7
Itens.ColAlignment(1) = 1
Itens.ColAlignment(2) = 1
Itens.ColAlignment(3) = 1
Itens.ColAlignment(4) = 1
Itens.ColAlignment(5) = 1
Itens.ColAlignment(6) = 1
Itens.ColAlignment(7) = 1
Itens.ColAlignment(8) = 1
Itens.ColAlignment(9) = 1
Itens.ColAlignment(10) = 1
Itens.ColAlignment(11) = 1
Itens.ColAlignment(12) = 1
Itens.ColAlignment(13) = 1
Itens.ColAlignment(14) = 1
Itens.ColAlignment(15) = 1
Itens.ColAlignment(16) = 1
Itens.ColAlignment(17) = 1
Itens.ColAlignment(18) = 1



For a = 0 To 18
   Itens.ColWidth(a) = 1000
Next

Itens.TextMatrix(0, 0) = "Numlancto"
Itens.TextMatrix(0, 1) = "Nf"
Itens.TextMatrix(0, 2) = "recdesp"
Itens.TextMatrix(0, 3) = "cricred"
Itens.TextMatrix(0, 4) = "Saldo"
Itens.TextMatrix(0, 5) = "tpmonet"
Itens.TextMatrix(0, 6) = "valor"
Itens.TextMatrix(0, 7) = "data"
Itens.TextMatrix(0, 8) = "Contr"
Itens.TextMatrix(0, 9) = "tipord"
Itens.TextMatrix(0, 10) = "tipomonete"
Itens.TextMatrix(0, 11) = "codigo"
'Itens.TextMatrix(0, 12) = "S. Emp."
'Itens.TextMatrix(0, 13) = "Brindes"
'Itens.TextMatrix(0, 14) = "Reb."
'Itens.TextMatrix(0, 15) = "Acres."
'Itens.TextMatrix(0, 16) = "Desc."
'Itens.TextMatrix(0, 17) = "V. A Rec"
'Itens.TextMatrix(0, 18) = "Dev. Pend"

LcTamanhoGrid = 1
End Function
Function MontaPesq()
On Error GoTo errormGrid
Dim RsCaixa As Recordset
Dim LcCriSql As String
Dim LcAchou As Integer
Dim a, TSDev, TSaDev, TSEmp, TFLoca, TFaDev, TFDev, TCliNo, TCliMov, TReb As Long
Dim TBrindes As Long
Dim TEntradas, TSaidas, tVendas, TAcre, TDesc As Currency

LcCriSql = "select * From ALID016 where Data Between #" & Format(Txt(0).Text, "mm/dd/yyyy") & "# And #" _
& Format(Txt(1).Text, "mm/dd/yyyy") & "# order by ALID016.Data"

Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsCaixa = Dbbase.OpenRecordset(LcCriSql) ', dbOpenDynaset)
LcAchou = False

a = 1
Do Until RsCaixa.EOF
   
      
      Itens.Rows = a + 1
      Itens.TextMatrix(a, 0) = RsCaixa!numlancto
      Itens.TextMatrix(a, 1) = Format(RsCaixa!nf, "Currency")
      Itens.TextMatrix(a, 2) = Format(RsCaixa!recdesp, "Currency")
      Itens.TextMatrix(a, 3) = Format(RsCaixa!clicred, "Currency")
      Itens.TextMatrix(a, 4) = Format(RsCaixa!saldo, "Currency")
      Itens.TextMatrix(a, 5) = RsCaixa!tpmonet
      Itens.TextMatrix(a, 6) = RsCaixa!valor
      Itens.TextMatrix(a, 7) = RsCaixa!Data
      Itens.TextMatrix(a, 8) = RsCaixa!contr
      Itens.TextMatrix(a, 9) = RsCaixa!tipord
      Itens.TextMatrix(a, 10) = RsCaixa!tpmonete
      Itens.TextMatrix(a, 11) = RsCaixa!Codigo
      'Itens.TextMatrix(a, 12) = RsCaixa!SacolasEmpresdas
      'Itens.TextMatrix(a, 13) = RsCaixa!Brindes
      'Itens.TextMatrix(a, 14) = RsCaixa!Rebobinacoes
      'Itens.TextMatrix(a, 15) = Format(RsCaixa!Acrescimos, "Currency")
      'Itens.TextMatrix(a, 16) = Format(RsCaixa!Descontos, "Currency")
      'Itens.TextMatrix(a, 17) = RsCaixa!SacolasEmpresdas
      'Itens.TextMatrix(a, 18) = RsCaixa!DevolucoesPendentes
      
      TEntradas = RsCaixa!entrada + TEntradas
      TSaidas = RsCaixa!Saida + TSaidas
      tVendas = RsCaixa!EntradaVenda + tVendas
      TAcre = RsCaixa!Acrescimos + TAcre
      TDesc = RsCaixa!Descontos + TDesc
      TSDev = RsCaixa!SacolasDev + TSDev
      TSaDev = RsCaixa!SacolasMov + TSaDev
      TSEmp = RsCaixa!SacolasEmpresdas + TSEmp
      TFLoca = RsCaixa!FitasLocadas + TFLoca
      TFaDev = RsCaixa!FitasaDev + TFaDev
      TFDev = RsCaixa!FitasaDev + TFDev
      TCliNo = RsCaixa!CLientesNovos + TCliNo
      TCliMov = RsCaixa!ClientesMov + TCliMov
      TReb = RsCaixa!Rebobinacoes + TReb
      RsCaixa.MoveNext
      a = a + 1
 Loop

 RsCaixa.Close
Txt(6).Text = Format(TEntradas, "Currency")
Txt(4).Text = Format(TSaidas, "Currency")
Txt(5).Text = Format(tVendas, "Currency")
Txt(3).Text = Format(TAcre, "Currency")
Txt(2).Text = Format(TDesc, "Currency")
Txt(8).Text = TSDev
Txt(9).Text = TSaDev
Txt(7).Text = TSEmp
Txt(14).Text = TFLoca
Txt(10).Text = TFaDev
Txt(15).Text = TFDev
Txt(13).Text = TCliNo
Txt(12).Text = TCliMov
Txt(11).Text = TReb
Exit Function
errormGrid:
Exit Function
Select Case ErrosSistema
       Case Is = 0
          Resume Next
       Case Is = 6
          Resume 0
       Case Is = 7
          End
       Case Else
        Exit Function
        MsgBox Err.Description & Err
End Select
End Function

Private Sub Itens_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "%+{TAB}"
End Sub

Private Sub txt_Change(Index As Integer)
If Len(Trim(Txt(1).Text)) > 0 Then
   CmdOk.Enabled = True
Else
   CmdOk.Enabled = False
End If
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then If Index = 0 Then txt(1).SetFocus
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "%+{TAB}"
End Sub

