VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form MostraVales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vales de Clientes"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid data2 
      Height          =   3135
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16445144
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "item"
         Caption         =   "Item"
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
         DataField       =   "Codprod"
         Caption         =   "Código"
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
         DataField       =   "descricao"
         Caption         =   "Descrição"
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
         DataField       =   "Qtde"
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
         DataField       =   "Valunit"
         Caption         =   "Valor Unitário"
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
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4350,047
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1005,165
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid Data1 
      Height          =   2655
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   15979936
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "marca"
         Caption         =   "Lanc"
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
         DataField       =   "numnf"
         Caption         =   "Nº Vale"
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
         DataField       =   "dtemis"
         Caption         =   "Emissão"
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
         DataField       =   "cliente"
         Caption         =   "Cliente"
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
         DataField       =   "ValorNota"
         Caption         =   "Valor do Vale"
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
            Alignment       =   2
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   5160,189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1349,858
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produtos do Vale"
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
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vales"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "MostraVales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Rs As ADODB.Recordset
Function SetaVales(LcCliente As String)

'abreconexao

LcSql = "Select * from Vales where cliente='" & LcCliente & "' and baixado=0" ' order by numnf"

Set Rs = AbreRecordset(LcSql)
Set Data1.DataSource = Rs
If Not Rs.EOF Then SetaPr
End Function
Function SetaPr()
Dim rs1 As ADODB.Recordset
Dim LcSql1 As String
'abreconexao
LcSql1 = "Select * from ValesProdutos where numnf='" & Data1.Columns(1) & "' order by item"
Set rs1 = AbreRecordset(LcSql1)
Set data2.DataSource = rs1

End Function

Private Sub Command1_Click()
On Error Resume Next
Dim LcV As String
Rs.MoveFirst
Do Until Rs.EOF
    If Rs!marca = "X" Then
       If Len(LcV) > 0 Then LcV = LcV & "|"
       LcV = LcV & Rs!numnf
    End If
    Rs.MoveNext
Loop
FrmSaidaProduto.CodVales.Text = LcV
FrmSaidaProduto.LancaVales
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Data1_Click()
On Error Resume Next
SetaPr
End Sub

Private Sub Data1_DblClick()
On Error Resume Next
Dim LcM As String
LcM = Data1.Columns(1)
Rs.Find "numnf='" & LcM & "'"
If Not Rs.EOF Then
    If Len(Rs!marca) > 0 Then
       Rs!marca = ""
    Else
       Rs!marca = "X"
    End If
    Rs.Update
End If
End Sub

