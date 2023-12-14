VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form FrmRelTipoMonet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Despesas por Tipo Monetário"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   3840
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo "
      Height          =   1215
      Left            =   2040
      TabIndex        =   16
      Top             =   1920
      Visible         =   0   'False
      Width           =   2175
      Begin VB.OptionButton sintetico 
         Caption         =   "Sintético"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton analitico 
         Caption         =   "Analítico"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.TextBox codigo 
      Height          =   405
      Left            =   2640
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Tipo 
      Height          =   315
      ItemData        =   "FrmRelDespTipoMonet.frx":0000
      Left            =   120
      List            =   "FrmRelDespTipoMonet.frx":0002
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar"
      Height          =   615
      Left            =   4680
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F3"
      Height          =   615
      Left            =   4680
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox copias 
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
      Left            =   4680
      TabIndex        =   7
      Text            =   "1"
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
      Begin VB.OptionButton impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Video 
         Caption         =   "Video"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Dataf 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial"
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
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Final"
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
      Left            =   2040
      TabIndex        =   13
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Monetário"
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
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1590
   End
   Begin VB.Line Line1 
      X1              =   4440
      X2              =   4440
      Y1              =   0
      Y2              =   3720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copias"
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
      Left            =   4680
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "FrmRelTipoMonet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TipoMonet
      codigo As String
      nome As String
End Type
Private LcTamanho, a As Integer
Private MtTipoMonet() As TipoMonet
Private RsTipoMonet As Recordset, RsSintetico As Recordset
Private Sub Cbo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Function CarregaTipoMonet()
On Error GoTo errc
Dim RsTipoMonet As Recordset
AbreBase
Set RsTipoMonet = Dbbase.OpenRecordset("ALID008", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcTamanho = 0
Do Until RsTipoMonet.EOF
   ReDim Preserve MtTipoMonet(LcTamanho)
   MtTipoMonet(LcTamanho).codigo = RsTipoMonet!TPMONET
   MtTipoMonet(LcTamanho).nome = RsTipoMonet!XTPMONET
   tipo.AddItem RsTipoMonet!XTPMONET
   RsTipoMonet.MoveNext
   LcTamanho = LcTamanho + 1
Loop
If LcTamanho > 0 Then LcTamanho = LcTamanho - 1
'Comissao.AddItem "TODOS"
'Comissao.Text = "TODOS"
RsTipoMonet.Close
Set RsTipoMonet = Nothing
Exit Function
errc:

Exit Function
MsgBox err.Description & err

End Function

Private Sub analitico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub TipoMonet_Click()
Dim a As Integer
For a = 0 To LcTamanho
    If MtTipoMonet(a).nome = tipo.Text Then
       codigo.Text = MtTipoMonet(a).codigo
       Exit For
    End If
Next
End Sub

Private Sub Comissao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim LcFormula, LcCriterio As String, LcTipoSaida, a As Integer
Dim RsEmpresa As Recordset, RsOpcao As Recordset
Dim LcEmpresa, LcEndereco, LcFone, Lccelular, Lccelular1, Lcemail, LcVer, LcVer1, LcCap As String
AbreBase
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsOpcao = Dbbase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LcEndereco = RsEmpresa!endereco & " " & RsEmpresa!bairro
   LcFone = RsEmpresa!fone
End If
If Not RsOpcao.EOF Then
   LcVer = RsOpcao!msg
   LcVer1 = RsOpcao!Msg1
End If

    'Abertura do relatório de vendas
        
    CryRelatorio.DataFiles(0) = GLBase
    
    lctitulo = "Relatório de Despesas  Por Tipo Monetário"
    If GlImprimeSemLinha Then
       CryRelatorio.ReportFileName = App.Path & "\despesa.rpt"
    Else
       CryRelatorio.ReportFileName = App.Path & "\despesasl.rpt"
    End If
    'CryRelatorio.SortFields(0) = "+{ALID201.VENDEDORr}"
    
    CryRelatorio.CopiesToPrinter = Val(Txt1.Text)
    

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
  For a = 0 To LcTamanho - 1
     If tipo.Text = MtTipoMonet(a).nome Then
        LcTipo = MtTipoMonet(a).codigo
        Exit For
     End If
  Next
 ' LcFormula = "{ALID014.TPMONET}='" & MtTipoMonet(tipo.ListIndex).codigo & "'"
  LcFormula = "{ALID014.DTVENC} >=" & LcChav1 & " And {ALID014.DTVENC} <=" & LcChav2 & " and {alid014.TPMONET}='" & MtTipoMonet(tipo.ListIndex).codigo & "'"


'== fim filtro
'== fim filtro
CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Despesa por Tipo Monetário"

 CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
 CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
 CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'"
 CryRelatorio.Formulas(3) = "titulo='Relatório de Despesa por Tipo Monetário'"
 'CryRelatorio.Formulas(4) = "Versiculo='" & LcVer & "'"
 'CryRelatorio.Formulas(5) = "Versiculo1='" & LcVer1 & "'"
 CryRelatorio.Formulas(5) = "Celular='" & Lccelular & "'"
 CryRelatorio.Formulas(4) = "Celular1='" & Lccelular1 & "'"
 CryRelatorio.Formulas(6) = "email='" & Lcemail & "'"

If impressora Then
   LcTipoSaida = 1
Else
   LcTipoSaida = 0
End If
CryRelatorio.SelectionFormula = LcFormula

CryRelatorio.Destination = LcTipoSaida
CryRelatorio.PrintReport
Me.Caption = LcCap
RsOpcao.Close
RsEmpresa.Close
Dbbase.Close
Set RsOpcao = Nothing
Set RsEmpresa = Nothing
Set Dbbase = Nothing
If CryRelatorio.LastErrorNumber > 0 Then MsgBox CryRelatorio.LastErrorString


End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Copias_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Dataf_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Dataf_LostFocus()
If Not IsDate(Dataf.Text) And Dataf.Text <> "  /  /  " Then
      MsgBox "A data digitada não é Válida...", 48, "Aviso"
      Dataf.SetFocus
End If
End Sub

Private Sub Datai_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
   
End Sub

Private Sub Datai_LostFocus()
If Not IsDate(Datai.Text) And Datai.Text <> "  /  /  " Then
      MsgBox "A data digitada não é Válida...", 48, "Aviso"
      Datai.SetFocus
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
DataS.Text = Format(GlDataSistema, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
Me.Height = 3675
Me.Width = 6300
CarregaTipoMonet

End Sub

Private Sub Impressora_Click()
copias.Visible = True
Label3.Visible = True
End Sub
Function GeraSintetico()
Dim LcComissao, LcTotal As Currency
Dim LcMuda As Integer
AbreBase
LcCriterio1 = "Select * from alid201 where VENDEDOR='" & codigo.Text & "' order by nf"
Set RsComissao = Dbbase.OpenRecordset(LcCriterio1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsSintetico = Dbbase.OpenRecordset("sintetico", dbOpenDynaset, dbSeeChanges, dbOptimistic)
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
      LcTotal = LcTotal + RsComissao!VALORTOTAL
   Else
      RsComissao.MovePrevious
      Call GravaSintetico(LcComissao, LcTotal)
      LcComissao = 0
      LcTotal = 0
      LcMuda = True
   End If
   RsComissao.MoveNext
   
Loop
End Function
Function GravaSintetico(LcComissao, LcTotal As Currency)
RsSintetico.AddNew
   RsSintetico!Vendedor = RsComissao!Vendedor
   RsSintetico!NF = RsComissao!NF
   RsSintetico!Comissao = LcComissao
   RsSintetico!VALORTOTAL = LcTotal
   RsSintetico!ITEMBAIXO = RsComissao!ITEMBAIXO
   RsSintetico!DATAVENDA = RsComissao!DATAVENDA
   RsSintetico!CLIENTE = RsComissao!CLIENTE
   RsSintetico.Update
End Function
Private Sub Impressora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub sintetico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Tipo_Click()
Dim a As Integer
For a = 0 To LcTamanho
    If MtTipoMonet(a).nome = tipo.Text Then
       codigo.Text = MtTipoMonet(a).codigo
       Exit For
    End If
Next

End Sub

Private Sub Tipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Video_Click()
copias.Visible = False
Label3.Visible = False
End Sub

Private Sub Video_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub
