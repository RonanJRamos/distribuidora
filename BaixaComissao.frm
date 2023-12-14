VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form BaixaComissao 
   BackColor       =   &H00E6E7C2&
   Caption         =   "Fechamento de Comissões"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10200
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   7455
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox PerComissao 
      Height          =   375
      Left            =   4800
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox Lucro 
      Height          =   375
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox Custo 
      Height          =   375
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox TSelecionada 
      Height          =   375
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1935
   End
   Begin VB.TextBox tcomissao 
      Height          =   375
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1695
   End
   Begin VB.TextBox tVendas 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00D3D597&
      Caption         =   "&Fechar F10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E6E7C2&
      Caption         =   "Exibição"
      Height          =   1095
      Left            =   3960
      TabIndex        =   12
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E6E7C2&
         Caption         =   "Sintético"
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
         TabIndex        =   14
         Top             =   720
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E6E7C2&
         Caption         =   "Analítico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E6E7C2&
      Caption         =   "Status"
      Height          =   1095
      Left            =   6240
      TabIndex        =   8
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E6E7C2&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E6E7C2&
         Caption         =   "Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   450
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E6E7C2&
         Caption         =   "Não Pago"
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
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Comissao 
      Height          =   4575
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   -2147483624
      BackColorFixed  =   13882775
      BackColorBkg    =   15987682
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00D3D597&
      Caption         =   "&Executa F2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
   Begin MSMask.MaskEdBox datai 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.ComboBox Vendedor 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin MSMask.MaskEdBox Dataf 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00D3D597&
      Caption         =   "&Imprimir F6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00D3D597&
      Caption         =   "Marcar &Todos Como Pago  F5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6840
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00D3D597&
      Caption         =   "&Marcar Todos como Não Pago  F4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6840
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00D3D597&
      Caption         =   "&Confirma Lançamento  F3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6840
      Width           =   1695
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E6E7C2&
      Caption         =   "Comissão %"
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
      Left            =   4800
      TabIndex        =   30
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E6E7C2&
      Caption         =   "Lucro %"
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
      Left            =   3480
      TabIndex        =   28
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E6E7C2&
      Caption         =   "Total Custo"
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
      Left            =   1800
      TabIndex        =   26
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E6E7C2&
      Caption         =   "Total Sel.  Pagar"
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
      Left            =   8160
      TabIndex        =   24
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E6E7C2&
      Caption         =   "Total Comissão"
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
      Left            =   6360
      TabIndex        =   22
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E6E7C2&
      Caption         =   "Total Vendas"
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
      TabIndex        =   20
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E6E7C2&
      Caption         =   "Data Final"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E6E7C2&
      Caption         =   "Data Inicial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1065
   End
End
Attribute VB_Name = "BaixaComissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Vendedor
        codigo As String
        Nome As String
End Type
Private LcTamanhoGrid As Long
Private MtVendedor() As Vendedor
Private LcTGrid, a As Long
Private LcCusto As Double
Private RsComissao As Recordset, rsCliente As Recordset, RsSintetico As Recordset

Private Sub Comissao_DblClick()
On Error Resume Next
Dim LcLinha As Integer
LcLinha = Comissao.Row
If Comissao.TextMatrix(LcLinha, 5) = "Sim" Then
   Comissao.TextMatrix(LcLinha, 5) = "Não"
   If Len(TSelecionada.Text) > 0 Then
     TSelecionada.Text = Format(CDbl(TSelecionada.Text) - CDbl(Comissao.TextMatrix(LcLinha, 4)), "Currency")
   End If
Else
  Comissao.TextMatrix(LcLinha, 5) = "Sim"
  
  If Len(TSelecionada.Text) > 0 Then
    TSelecionada.Text = Format(CDbl(TSelecionada.Text) + CDbl(Comissao.TextMatrix(LcLinha, 4)), "Currency")
  Else
    TSelecionada.Text = Format(CDbl(Comissao.TextMatrix(LcLinha, 4)), "Currency")
  End If
End If
End Sub

Private Sub Comissao_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then
If Comissao.TextMatrix(LcLinha, 5) = "Sim" Then
   Comissao.TextMatrix(LcLinha, 5) = "Não"
   If Len(TSelecionada.Text) > 0 Then
     TSelecionada.Text = Format(CDbl(TSelecionada.Text) - CDbl(Comissao.TextMatrix(LcLinha, 4)), "Currency")
   End If
Else
  Comissao.TextMatrix(LcLinha, 5) = "Sim"
  
  If Len(TSelecionada.Text) > 0 Then
    TSelecionada.Text = Format(CDbl(TSelecionada.Text) + CDbl(Comissao.TextMatrix(LcLinha, 4)), "Currency")
  Else
    TSelecionada.Text = Format(CDbl(Comissao.TextMatrix(LcLinha, 4)), "Currency")
  End If
End If
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim LcCap As String
LcCusto = 0
Custo.Text = ""
PerComissao.Text = ""
Lucro.Text = ""
If Len(Vendedor.Text) = 0 Then
   MsgBox "Escolha o Vendedor Para Listar as Comissões...", 64, "Aviso"
   Vendedor.SetFocus
   Exit Sub
End If

If Datai.Text = "  /  /  " Then
   MsgBox "Escolha a Data Inicial do Periodo ...", 64, "Aviso"
   Vendedor.SetFocus
   Exit Sub
End If
If Dataf.Text = "  /  /  " Then
   Dataf.Text = Date
   'Exit Sub
End If
LcCap = Me.Caption
Me.Caption = "Aguarde, Filtrando Registros..."
Comissao.Rows = 1
If Option4 Then montagrid
If Option5 Then GeraSintetico

If GlComissaoBelclean Then
    For a = 1 To Comissao.Rows - 1
        BuscaCustoComissao (Comissao.TextMatrix(a, 0))
    Next
    Custo.Text = Format(CCur(LcCusto), "Currency")
    Lucro.Text = AcertaNumero(((CCur(tVendas.Text) * 100) / CCur(Custo.Text)) - 100, 3)
End If
Me.Caption = LcCap
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command2_Click()
On Error Resume Next
For a = 1 To Comissao.Rows - 1
    Comissao.TextMatrix(a, 5) = "Não"
Next
TSelecionada.Text = ""

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim LcTotal As Double
For a = 1 To Comissao.Rows - 1
    Comissao.TextMatrix(a, 5) = "Sim"
    LcTotal = LcTotal + CDbl(Comissao.TextMatrix(a, 4))
Next
TSelecionada.Text = Format(LcTotal, "Currency")
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command4_Click()
' On Error Resume Next
If Option4 Then BaixaComissaoAnalitico
If Option5 Then Baixasintetico
Comissao.Rows = 1
Vendedor.Text = ""
Datai.Text = "  /  /  "
Dataf.Text = "  /  /  "
tVendas.Text = ""
tcomissao.Text = ""
Custo.Text = ""
PerComissao.Text = ""
Lucro.Text = ""

TSelecionada.Text = ""
End Sub
Function BaixaComissaoAnalitico()
On Error Resume Next
Dim RsComissao As Recordset
AbreBase
Set RsComissao = Dbbase.OpenRecordset("alid201", dbOpenDynaset, dbSeeChanges, dbOptimistic)
For a = 1 To Comissao.Rows - 1
    If Comissao.TextMatrix(a, 5) = "Sim" Then
       LcCri = "Codigo=" & Val(Comissao.TextMatrix(a, 6))
       RsComissao.FindFirst LcCri
       If Not RsComissao.NoMatch Then
          RsComissao.Edit
          RsComissao("pago") = True
          RsComissao.Update
       End If
    Else
       LcCri = "Codigo=" & Val(Comissao.TextMatrix(a, 6))
       RsComissao.FindFirst LcCri
       If Not RsComissao.NoMatch Then
          RsComissao.Edit
          RsComissao("pago") = False
          RsComissao.Update
       End If
    End If
Next
RsComissao.Close

End Function
Function Baixasintetico()
On Error Resume Next
Dim RsComissao As Recordset
AbreBase

For a = 1 To Comissao.Rows - 1
   If Comissao.TextMatrix(a, 5) = "Sim" Then
      LcSql = "select * from alid201 where nf='" & Comissao.TextMatrix(a, 0) & "'"
      Set RsComissao = Dbbase.OpenRecordset(LcSql) '"alid201", dbOpenDynaset, dbSeeChanges, dbOptimistic)
      Do Until RsComissao.EOF
         RsComissao.Edit
         RsComissao("pago") = True
         RsComissao.Update
         RsComissao.MoveNext
      Loop
   Else
     LcSql = "select * from alid201 where nf='" & Comissao.TextMatrix(a, 0) & "'"
      Set RsComissao = Dbbase.OpenRecordset(LcSql) '"alid201", dbOpenDynaset, dbSeeChanges, dbOptimistic)
      Do Until RsComissao.EOF
         RsComissao.Edit
         RsComissao("pago") = False
         RsComissao.Update
         RsComissao.MoveNext
      Loop
   End If
Next
'RsComissao.Close

End Function

Private Sub Command4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command5_Click()
On Error Resume Next
Unload Me

End Sub

Private Sub Command5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command6_Click()
On Error Resume Next
On Error Resume Next
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, RsOpcao As Recordset
Dim LcEmpresa, LcEndereco, LcFone, Lccelular, Lccelular1, Lcemail, LcVer, LcCap, LcVer1 As String
AbreBase
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone
End If
 '== Inicio Filtro
  strData = CDate(Format(Datai.Text, "dd/mm/yyyy"))
  LcAno = Year(strData)
  LcMes = Month(strData)
  LcDia = Day(strData)
  LcDataInicio = LcAno & "," & LcMes & "," & LcDia
  LcChav1 = " date(" & LcDataInicio & ")"
         
  strData = CDate(Format(Dataf.Text, "dd/mm/yyyy"))
  LcAno = Year(strData)
  LcMes = Month(strData)
  LcDia = Day(strData)
  LcDataInicio = LcAno & "," & LcMes & "," & LcDia
  LcChav2 = " date(" & LcDataInicio & ")"
Dim a As Long
Dim LcCodigoVendedor As String
For a = 0 To Vendedor.ListCount - 1
    If MtVendedor(a).Nome = Vendedor.Text Then
       LcCodigoVendedor = MtVendedor(a).codigo
       Exit For
    End If
Next
CryRelatorio.DataFiles(0) = GLBase
CryRelatorio.ReportFileName = App.Path & "\RelComissaoD.rpt"
LcFormula = "{ALID201.VENDEDOR}='" & LcCodigoVendedor & "'"
If Len(LcFormula) > 0 Then LcFormula = LcFormula & " And "
LcFormula = LcFormula & "({ALID201.DATAVENDA} >=" & LcChav1 & " And {ALID201.DATAVENDA} <=" & LcChav2 & ")"
'MsgBox LcFormula
CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Comissões"
CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'"
CryRelatorio.Formulas(5) = "titulo='Comissões'"
CryRelatorio.Formulas(3) = "Celular='" & Lccelular & "'"
CryRelatorio.Formulas(4) = "Celular1='" & Lccelular1 & "'"
CryRelatorio.Formulas(7) = "Lucro='" & Lucro.Text & "'"
CryRelatorio.Formulas(8) = "Percentual='" & PerComissao.Text & "'"
CryRelatorio.Formulas(9) = "Custo='" & Custo.Text & "'"
CryRelatorio.Formulas(10) = "Comissao='" & tcomissao.Text & "'"
CryRelatorio.Formulas(11) = "Vendedor='" & Vendedor.Text & "'"
LcTipoSaida = 0
CryRelatorio.SelectionFormula = LcFormula
CryRelatorio.Destination = LcTipoSaida
CryRelatorio.PrintReport
Me.Caption = LcCap
RsEmpresa.Close
Dbbase.Close
Set RsEmpresa = Nothing
Set Dbbase = Nothing
If CryRelatorio.LastErrorNumber > 0 Then MsgBox CryRelatorio.LastErrorString
End Sub

Private Sub Dataf_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Dataf_LostFocus()
On Error Resume Next
If Dataf.Text = "  /  /  " Then Exit Sub
If Not IsDate(Dataf.Text) Then
   MsgBox "Data Inválida", 64, "Aviso"
   Dataf.SetFocus
End If
End Sub

Private Sub Datai_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Datai_LostFocus()
On Error Resume Next
If Datai.Text = "  /  /  " Then Exit Sub
If Not IsDate(Datai.Text) Then
   MsgBox "Data Inválida", 64, "Aviso"
   Datai.SetFocus
End If
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
Label7.Visible = GlComissaoBelclean
Label8.Visible = GlComissaoBelclean
Label9.Visible = GlComissaoBelclean
Custo.Visible = GlComissaoBelclean
Lucro.Visible = GlComissaoBelclean
Command6.Visible = GlComissaoBelclean
PerComissao.Visible = GlComissaoBelclean
GeraVendedor
GeraGrid
End Sub

Function GeraVendedor()
On Error Resume Next
Dim RsVendedor As Recordset
LcTGrid = 0
LcCriSql = "VENDEDOR='"
AbreBase
Set RsVendedor = Dbbase.OpenRecordset("select * from alid200 order by nome") ', dbOpenDynaset)
Do Until RsVendedor.EOF
   If Not IsNull(RsVendedor!Nome) Then
      If err.Number > 0 Then Exit Do
      ReDim Preserve MtVendedor(LcTGrid)
      MtVendedor(LcTGrid).codigo = RsVendedor!codigo
      MtVendedor(LcTGrid).Nome = RsVendedor!Nome
      Vendedor.AddItem RsVendedor!Nome
      LcTGrid = LcTGrid + 1
    End If
    RsVendedor.MoveNext
Loop
LcTGrid = LcTGrid - 1
RsVendedor.Close
Set RsVendedor = Nothing

End Function
Function GeraGrid()
On Error Resume Next
Comissao.ColAlignment(0) = 1
Comissao.ColAlignment(1) = 1
Comissao.ColAlignment(2) = 1
Comissao.ColAlignment(3) = 7
Comissao.ColAlignment(4) = 7
Comissao.ColAlignment(5) = 1

Comissao.ColWidth(0) = 950
Comissao.ColWidth(1) = 1100
Comissao.ColWidth(2) = 5000
Comissao.ColWidth(3) = 1000
Comissao.ColWidth(4) = 1000
Comissao.ColWidth(5) = 700
Comissao.ColWidth(6) = 0
Comissao.TextMatrix(0, 0) = "Documento"
Comissao.TextMatrix(0, 1) = "Pag.Comissão"
Comissao.TextMatrix(0, 2) = "Cliente"
Comissao.TextMatrix(0, 3) = "V.Venda"
Comissao.TextMatrix(0, 4) = "Comissão"
Comissao.TextMatrix(0, 5) = "Pago"

LcTamanhoGrid = 1
End Function
Function montagrid()
On Error Resume Next
Dim bb As Database
Dim RsProduto As Recordset
Dim LcCodigoVendedor As String
Dim LcTotalVendas, LcTotalComissao, LcTotalSelec As Double
Dim LcPago As Integer
If Option1 Then LcPago = False
If Option2 Then LcPago = True
'=== Busca Codigo Vendedor
For a = 0 To LcTGrid
    If MtVendedor(a).Nome = Vendedor.Text Then
       LcCodigoVendedor = MtVendedor(a).codigo
       Exit For
    End If
Next
LcCriSql = "select * from alid201 where VENDEDOR='" & LcCodigoVendedor & "' And DATAVENDA >= #" & Format(Datai.Text, "mm/dd/yyyy") & "# and datavenda <= #" & Format(Dataf.Text, "mm/dd/yyyy") & "#"
If Not Option3 Then
   LcCriSql = LcCriSql & " And pago=" & LcPago
End If
LcCriSql = LcCriSql & " order by nf"
Set bb = OpenDatabase(GLBase, False, False) ' "dBASE III;")
Set RsComissao = bb.OpenRecordset(LcCriSql) ', dbOpenDynaset)
Set rsCliente = bb.OpenRecordset("alid001", dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)
Set RsProduto = bb.OpenRecordset("alid009", dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)

LcTamanho = Comissao.Rows
a = 2
Me.Caption = msg
Comissao.Rows = 1
LcAchou = False
Comissao.TextMatrix(0, 2) = "Produto"
Do Until RsComissao.EOF
  LcAchou = True
  If Len(Trim(RsComissao!NF)) > 0 Then
   If Not IsNull(RsComissao!NF) Then
       
     Comissao.Rows = a
     Comissao.TextMatrix(a - 1, 0) = RsComissao!NF & ""
     Comissao.TextMatrix(a - 1, 1) = RsComissao!DATAVENDA & ""
     LcPesq = "cod='" & RsComissao!Produto & "'"
     RsProduto.FindFirst LcPesq
     Comissao.Rows = a
     If Not RsProduto.NoMatch Then
       Comissao.TextMatrix(a - 1, 2) = RsProduto!Nome & ""
     End If
    ' Comissao.TextMatrix(a - 1, 2) = RsComissao!Cliente & ""
     Comissao.TextMatrix(a - 1, 3) = Format(RsComissao!ValorTotal, "currency")
     Comissao.TextMatrix(a - 1, 4) = Format(RsComissao!Comissao, "currency")
     LcTotalVendas = LcTotalVendas + CDbl(Format(RsComissao!ValorTotal, "currency"))
     LcTotalComissao = LcTotalComissao + CDbl(Format(RsComissao!Comissao, "Currency"))
     
     
     If RsComissao!pago Then
        Comissao.TextMatrix(a - 1, 5) = "Sim"
        LcTotalSelec = LcTotalSelec + CDbl(Format(RsComissao!Comissao, "Currency"))
     Else
        Comissao.TextMatrix(a - 1, 5) = "Não"
     End If
     Comissao.TextMatrix(a - 1, 6) = RsComissao!codigo
     a = a + 1
    End If
   End If
   RsComissao.MoveNext
Loop
tVendas.Text = Format(LcTotalVendas, "Currency")
tcomissao.Text = Format(LcTotalComissao, "currency")
TSelecionada.Text = Format(LcTotalSelec, "currency")
End Function
Function GeraSintetico()
'On Error Resume Next
Dim LcComissao, LcTotal As Currency
Dim LcMuda, LcGrava As Integer
Dim LcTotalSelec As Currency
Dim LcPago As Integer
Dim rsCliente As Recordset

For a = 0 To LcTGrid
    If MtVendedor(a).Nome = Vendedor.Text Then
       LcCodigoVendedor = MtVendedor(a).codigo
       Exit For
    End If
Next

AbreBase
LcCriterio1 = "Select * from alid201 where VENDEDOR='" & LcCodigoVendedor & "' and "
LcCriterio1 = LcCriterio1 & " DATAVENDA>=#" & Format(Datai.Text, "mm/dd/yyyy") & "# and DATAVENDA <=#" & Format(Dataf.Text, "mm/dd/yyyy") & "#"
'LcCriterio1 = LcCriterio1 & " Order by Nf"
If Option1 Then LcPago = 0
If Option2 Then LcPago = -1
If Not Option3 Then
   LcCriterio1 = LcCriterio1 & " And pago=" & LcPago
End If
LcCriterio1 = LcCriterio1 & " order by nf"

'MsgBox LcCriterio1
Set rsCliente = Dbbase.OpenRecordset("alid001") ', dbOpenDynaset)

Comissao.TextMatrix(0, 2) = "Cliente"

Debug.Print LcCriterio1
Set RsComissao = Dbbase.OpenRecordset(LcCriterio1)
Set RsSintetico = Dbbase.OpenRecordset("sintetico", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcTotalSelec = 0
Do Until RsSintetico.EOF
   RsSintetico.Delete
   RsSintetico.MoveNext
Loop
LcNota = RsComissao!NF
Do Until RsComissao.EOF
   If LcMuda Then
      LcNota = RsComissao!NF
      LcMuda = False
   End If
   If LcNota = RsComissao!NF Then
      LcComissao = LcComissao + RsComissao!Comissao
      LcTotal = LcTotal + RsComissao!ValorTotal
      LcGrava = True
   Else
      RsComissao.MovePrevious
      Call GravaSintetico(LcComissao, LcTotal)
      LcComissao = 0
      LcTotal = 0
      LcMuda = True
      LcGrava = False
   End If
   RsComissao.MoveNext
   
Loop
If LcGrava Then
   RsComissao.MovePrevious
   Call GravaSintetico(LcComissao, LcTotal)
   LcGrava = False
End If
RsSintetico.MoveFirst
LcTamanho = Comissao.Rows
a = 2
Do Until RsSintetico.EOF
  LcAchou = True
  If Len(Trim(RsSintetico!NF)) > 0 Then
   If Not IsNull(RsSintetico!NF) Then
     LcPesq = "codigo='" & RsSintetico!Cliente & "'"
    ' RsCliente.FindFirst LcPesq
     Comissao.Rows = a
     Comissao.TextMatrix(a - 1, 0) = RsSintetico!NF & ""
     Comissao.TextMatrix(a - 1, 1) = RsSintetico!DATAVENDA & ""
     'If Not RsCliente.NoMatch Then
       Comissao.TextMatrix(a - 1, 2) = RsSintetico!Cliente & ""
     'End If
     LcTotalVendas = LcTotalVendas + CDbl(Format(RsSintetico!ValorTotal, "currency"))
     LcTotalComissao = LcTotalComissao + CDbl(Format(RsSintetico!Comissao, "Currency"))
     Comissao.TextMatrix(a - 1, 3) = Format(RsSintetico!ValorTotal, "currency")
     Comissao.TextMatrix(a - 1, 4) = Format(RsSintetico!Comissao, "currency")

     If RsSintetico!pago Then
        Comissao.TextMatrix(a - 1, 5) = "Sim"
        LcTotalSelec = LcTotalSelec + CCur(RsComissao!Comissao)
     Else
        Comissao.TextMatrix(a - 1, 5) = "Não"
     End If
     a = a + 1
    End If
   End If
   RsSintetico.MoveNext
Loop
tVendas.Text = Format(LcTotalVendas, "Currency")
tcomissao.Text = Format(LcTotalComissao, "currency")
TSelecionada.Text = Format(LcTotalSelec, "currency")
End Function
Function GravaSintetico(LcComissao, LcTotal As Currency)
Dim rsCliente As Recordset
'On Error Resume Next
LcCriterio22 = "Select * from alid001 where codigo='" & RsComissao!Cliente & "'"
Set rsCliente = Dbbase.OpenRecordset(LcCriterio22)

RsSintetico.AddNew
   RsSintetico!Vendedor = RsComissao!Vendedor
   RsSintetico!NF = RsComissao!NF
   RsSintetico!Comissao = LcComissao
   RsSintetico!ValorTotal = LcTotal
   RsSintetico!ItemBaixo = RsComissao!ItemBaixo
   RsSintetico!DATAVENDA = RsComissao!DATAVENDA
    RsSintetico!pago = RsComissao!pago
   RsSintetico!Cliente = rsCliente!razaosoc
RsSintetico.Update
rsCliente.Close
Set rsCliente = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Option3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Option4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Option5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub PerComissao_Change()
If IsNumeric(PerComissao.Text) Then
    tcomissao.Text = Format((CCur(tVendas.Text) * CCur(PerComissao.Text)) / 100, "currency")
    TSelecionada.Text = tcomissao.Text
End If
End Sub

Private Sub tcomissao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub TSelecionada_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub tVendas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Vendedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub
Function BuscaCustoComissao(LcNf As String)
On Error Resume Next
Dim RsL As ADODB.Recordset
Dim LcSql As String
Dim LcValor As Double
LcSql = "SELECT alid052.NUMNF, alid052.codProd, alid052.QTDE, alid052.VALUNIT, alid052.QTDUM, produtos.Preco, produtos.custoTotal, produtos.QtdMedida"
LcSql = LcSql & " FROM alid052 INNER JOIN produtos ON alid052.codProd = produtos.codigo"
LcSql = LcSql & " Where ALID052.NUMNF='" & LcNf & "'"
Set RsL = AbreRecordset(LcSql, True)
Do Until RsL.EOF
    LcValor = 0
    If Not IsNull(RsL!CustoTotal) Then
        LcValor = CCur(RsL!CustoTotal) / CCur(RsL!QtdMedida)
        LcValor = LcValor * (CCur(RsL!QTDUM) * CCur(RsL!Qtde))
        LcCusto = LcCusto + LcValor
    End If
    RsL.MoveNext
Loop
RsL.Close
Set RsL = Nothing
End Function
