VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form receitasnaoquitadas 
   BackColor       =   &H00E4E3D6&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Receitas Não Quitadas"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E4E3D6&
      Caption         =   "Ordenar Por"
      Height          =   735
      Left            =   3840
      TabIndex        =   9
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton OptDoc 
         BackColor       =   &H00E4E3D6&
         Caption         =   "Documento"
         Height          =   195
         Left            =   1680
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptVencimento 
         BackColor       =   &H00E4E3D6&
         Caption         =   "vencimento"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin MSMask.MaskEdBox DataI 
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton CmdFiltrar 
      Caption         =   "Filtrar"
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Doc 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5055
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8916
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   14342603
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Codigo"
         Caption         =   "Codigo"
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
         DataField       =   "NF"
         Caption         =   "Doc"
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
         DataField       =   "data"
         Caption         =   "Lançamento"
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
         DataField       =   "valor"
         Caption         =   "Valor"
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
         DataField       =   "Dtvenc"
         Caption         =   "Vencimento"
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
         DataField       =   "nomedesp"
         Caption         =   "Receita"
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
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   3330,142
         EndProperty
      EndProperty
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F2"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   6240
      Width           =   2535
   End
   Begin MSMask.MaskEdBox DataF 
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E4E3D6&
      Caption         =   "Periodo Vencimento"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E4E3D6&
      Caption         =   "Documento"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "receitasnaoquitadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFiltrar_Click()
On Error Resume Next
Dim Rsa As ADODB.Recordset
Dim LcCriterio As String

LcCriterio = "Select * from alid015 where nf like '" & Doc.Text & "%'"
If IsDate(DataI.Text) And IsDate(DataF.Text) Then
   LcCriterio = LcCriterio & " and DTVENC between #" & Format(DataI.Text, "mm/dd/yy") & "# and #" & Format(DataF.Text, "mm/dd/yy") & "#"
Else
  If IsDate(DataI.Text) Then
    LcCriterio = LcCriterio & " and DTVENC = #" & Format(DataI.Text, "mm/dd/yy") & "#"
  End If
  
End If
If OptVencimento.Value Then
    LcCriterio = LcCriterio & " order by DTVENC"
End If
If OptDoc.Value Then
    LcCriterio = LcCriterio & " order by nf"
End If

Set Rsa = AbreRecordset(LcCriterio)

Set DataGrid1.DataSource = Rsa
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim a As Long

a = 0 'DataGrid1.Col
FrmBaixaReceita.Codigo.Text = DataGrid1.Columns(0)
FrmBaixaReceita.Txt(0).Text = DataGrid1.Columns(1)
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Desp_DblClick()
SendKeys "%{C}"
End Sub

Private Sub Desp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "%{C}"
End If
If KeyCode = 121 Then Unload Me
If KeyCode = 113 Then SendKeys "%{C}"
End Sub

Private Sub DataGrid1_DblClick()
On Error Resume Next
SendKeys "%{C}"

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
SendKeys "%{C}"
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim Rsa As ADODB.Recordset
Dim LcCriterio As String

LcCriterio = "Select * from alid015 where VALPAGO=0 order by DTVENC"
'abreconexao
Set Rsa = AbreRecordset(LcCriterio)
Set DataGrid1.DataSource = Rsa

End Sub
