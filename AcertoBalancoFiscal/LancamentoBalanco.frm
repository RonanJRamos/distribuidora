VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form LancamentoBalanco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento do Balanço anterior"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdFiltrar 
      Caption         =   "Filtrar"
      Height          =   375
      Left            =   9720
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Produto 
      Height          =   285
      Left            =   7200
      TabIndex        =   6
      Top             =   720
      Width           =   4575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Mostrar Somente Não lançados"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   3495
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7215
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12726
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Data"
         Caption         =   "Data"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "codigoProduto"
         Caption         =   "Cod Produto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Nome"
         Caption         =   "Nome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Quantidade"
         Caption         =   "Quantidade"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "ValorCustoMedioUnitario"
         Caption         =   "Custo Medio Unitario"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "VCustoTotal"
         Caption         =   "Valor custo Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Saldo"
         Caption         =   "Saldo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   3225,26
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1560,189
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1349,858
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
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
      Left            =   7200
      TabIndex        =   7
      Top             =   480
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data "
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
      Left            =   7200
      TabIndex        =   5
      Top             =   120
      Width           =   570
   End
   Begin VB.Line Line1 
      X1              =   7080
      X2              =   7080
      Y1              =   0
      Y2              =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data do Balanço"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1755
   End
End
Attribute VB_Name = "LancamentoBalanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
On Error GoTo errcarregando
Dim RsBalanco As ADODB.Recordset
Dim StrSql As String
If Check1.Value = 1 Then
   StrSql = "Select * from estoquefiscal where quantidade=0 order by nome"
Else
   StrSql = "Select * from estoquefiscal order by nome"
End If

Set RsBalanco = AbreRecordset(StrSql)
Set DataGrid1.DataSource = RsBalanco


Exit Sub
errcarregando:
MsgBox Err.Description & Err.Number
End Sub

Private Sub CmdFiltrar_Click()
On Error GoTo errfiltro
Dim RsBalanco As ADODB.Recordset
Dim StrSql As String
Dim StrWhere As String
LcCap = Me.Caption
Me.Caption = "Aguarde, efetuando filtro..."
Screen.MousePointer = 11
StrSql = "Select * from estoquefiscal "
If IsDate(MaskEdBox1.Text) Then
   StrWhere = "where data='" & Format(MaskEdBox1.Text, "yyyy-mm-dd") & "'"
End If
If Len(Produto.Text) > 0 Then
   If Len(StrWhere) > 0 Then
      StrWhere = StrWhere & " and nome like '" & UCase(Produto.Text) & "%'"
   Else
      StrWhere = "where nome like '" & UCase(Produto.Text) & "%'"
   End If
End If
StrSql = StrSql & StrWhere & " Order by nome"
Set RsBalanco = AbreRecordset(StrSql)
Set DataGrid1.DataSource = RsBalanco

Me.Caption = LcCap
Screen.MousePointer = 0

Exit Sub
errfiltro:
Me.Caption = LcCap
Screen.MousePointer = 0
MsgBox Err.Description & Err.Number

End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
Dim Quantidade As Double
Dim ValorUnit As Double
Dim Valor As Double
Dim Col As Long
Col = DataGrid1.Col
If DataGrid1.Col = 4 Then
    DataGrid1.Col = 3
    Quantidade = IIf(Len(DataGrid1.Text) > 0, CDbl(DataGrid1.Text), 0)
    DataGrid1.Col = 4
    ValorUnit = IIf(Len(DataGrid1.Text) > 0, CDbl(DataGrid1.Text), 0)
    Valor = Quantidade * ValorUnit
    DataGrid1.Col = 5
    DataGrid1.Text = Valor
    DataGrid1.Col = 6
    DataGrid1.Text = Quantidade
    DataGrid1.AllowAddNew = True
    DataGrid1.Row = DataGrid1.Row + 1
    DataGrid1.Col = 3
End If


End Sub

Private Sub DataGrid1_GotFocus()
If Not IsDate(Data.Text) Then
   MsgBox "Informe a data do balanço", 64, "Aviso"
   Data.SetFocus
   Exit Sub
End If
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
   Load Pesquisa
   Pesquisa.Tag = DataGrid1.Row
   Pesquisa.Show , Me
End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If DataGrid1.Col > 2 Then
   If KeyAscii = 44 Then KeyAscii = 46
End If
End Sub

Private Sub Form_Load()
On Error GoTo errcarregando
Dim RsBalanco As ADODB.Recordset
Dim StrSql As String

StrSql = "Select * from estoquefiscal order by nome"
Set RsBalanco = AbreRecordset(StrSql)
Set DataGrid1.DataSource = RsBalanco


Exit Sub
errcarregando:
MsgBox Err.Description & Err.Number
End Sub


