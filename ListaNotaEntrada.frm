VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ListaNotaEntrada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista nota de Entrada"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdFiltrar 
      Caption         =   "Filtrar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Nota 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "NF"
         Caption         =   "NF"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Data"
         Caption         =   "Data"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "cfop"
         Caption         =   "CFOP"
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
         DataField       =   "clicred"
         Caption         =   "Fornec."
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
         DataField       =   "ValorProduto"
         Caption         =   "Valor Produto"
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
         DataField       =   "ipi"
         Caption         =   "Ipi"
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
         DataField       =   "valor"
         Caption         =   "Total"
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
      BeginProperty Column07 
         DataField       =   "BaseIcms"
         Caption         =   "Base Icms"
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
      BeginProperty Column08 
         DataField       =   "Icms"
         Caption         =   "Icms"
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
            ColumnWidth     =   1110,047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   689,953
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column08 
            Locked          =   -1  'True
            ColumnWidth     =   1080
         EndProperty
      EndProperty
   End
   Begin VB.Label Fornecedor 
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "Nº Nota"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Data da Nota"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "ListaNotaEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFiltrar_Click()
On Error Resume Next
Dim StrSql As String
Dim Rs As ADODB.Recordset

If Len(Nota.Text) > 0 Then
   StrSql = "Select * from entradanf where nf like '%" & Nota.Text & "%' order by nf"
Else
   If IsDate(Data.Text) Then
       StrSql = "Select * from entradanf where data='" & Format(Data.Text, "yyyy-mm-dd") & "' order by data"
   Else
       StrSql = "Select * from entradanf order by nf"
   End If
End If
Set Rs = AbreRecordset(StrSql)
Set DataGrid1.DataSource = Rs

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 44 Then KeyAscii = 46
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
Dim db As Database
Dim RsF As Recordset
Dim Colant As Integer
Colant = DataGrid1.Col

DataGrid1.Col = 3
Set db = OpenDatabase(GLBase)
Set RsF = db.OpenRecordset("Select * from alid002 where codigo='" & DataGrid1.Text & "'")
DataGrid1.Col = Colant
If Not RsF.EOF Then
   Fornecedor.Caption = RsF!razaosoc & ""
Else
   Fornecedor.Caption = ""
End If
Set RsF = Nothing

End Sub

