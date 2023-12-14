VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form relentradaProdutoNf 
   Caption         =   "Relatório de Entrada de Produtos - NF"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton fechar 
      Caption         =   "Fechar"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   3495
   End
   Begin VB.CommandButton imprimir 
      Caption         =   "Imprimir"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3495
   End
   Begin Crystal.CrystalReport Cryrelatorio 
      Left            =   360
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSMask.MaskEdBox datai 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
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
      Left            =   1920
      TabIndex        =   1
      Top             =   840
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
      Caption         =   "Emissão entre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "relentradaProdutoNf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Fechar_Click()
Unload Me
End Sub

Private Sub imprimir_Click()
'Efetua a Selecao Campo
Dim LcFormula As String
'On Error Resume Next
Dim RsEmpresa As Recordset
Dim a, item, LcResposta As Long
Dim LcCriterio, LcEmpresa, LcEndereco, LcFone As String


Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset)

LcEmpresa = RsEmpresa!Razao
LcEndereco = RTrim(RsEmpresa!Endereco) & " Bairro: " & RsEmpresa!Bairro & "  Cidade: " & RsEmpresa!Cidade
LcFone = "Fone: " & RsEmpresa!Fone
If Not IsNull(RsEmpresa!Fax) Then
   LcFone = LcFone & " Fax: " & RsEmpresa!Fax
End If
If err <> 0 Then
 LcEmpresa = ""
 LcFone = ""
 LcEndereco = ""
End If
RsEmpresa.Close
If LcCampo = "datacompra" Then
         strData = CDate(Format(datai.Text, "dd/mm/yyyy"))
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
         LcFormula = "{entradanf.data} >=" & LcChav1 & " And {Entradanf.data} <=" & LcChav2
        lctitulo = "Entrada de Produtos de NF Período " & Format(datai.Text, "dd/mm/yyyy") & " Até " & Format(Dataf.Text, "dd/mm/yyyy")

End If
' MsgBox LcFormula
Cryrelatorio.DataFiles(0) = GLBase
Cryrelatorio.ReportFileName = App.Path & "\EntradaProdutonf.rpt"
Cryrelatorio.SelectionFormula = LcFormula

Cryrelatorio.SortFields(0) = "+{Entradanf.data" & LcCampo & "}"
'Cryrelatorio.CopiesToPrinter = Val(Copias.Text)


Cryrelatorio.WindowTop = 50
Cryrelatorio.WindowWidth = 700
Cryrelatorio.WindowLeft = 50
Cryrelatorio.WindowHeight = 500
Cryrelatorio.WindowTitle = "Entrada de Produtos - NF "

'Cryrelatorio.Formulas(0) = "titulo='" & lctitulo & "'"
'Cryrelatorio.Formulas(1) = "Empresa='" & LcEmpresa & "'"
'Cryrelatorio.Formulas(2) = "Endereco='" & LcEndereco & "'"
'Cryrelatorio.Formulas(3) = "Fone='" & LcFone & "'"


   LcTipoSaida = 0
   
'Relatorio.SelectionFormula = LcFormula
Cryrelatorio.Destination = LcTipoSaida
Cryrelatorio.PrintReport

If Cryrelatorio.LastErrorNumber > 0 Then
   If Cryrelatorio.LastErrorString <> "No Error" Then
     If Len(Trim(Cryrelatorio.LastErrorString)) <> 0 Then
        MsgBox Cryrelatorio.LastErrorString
     End If
   End If
End If

End Sub
