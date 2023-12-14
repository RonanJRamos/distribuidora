VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmRelatorioCTE 
   BackColor       =   &H00B3E9FD&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relatório CTe"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00B3E9FD&
      Caption         =   "Saída"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
      Begin VB.OptionButton Video 
         BackColor       =   &H00B3E9FD&
         Caption         =   "Video"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton impressora 
         BackColor       =   &H00B3E9FD&
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
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
      Left            =   2160
      TabIndex        =   2
      Text            =   "1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F3"
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   3840
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   3960
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
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
      TabIndex        =   7
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
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
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   2160
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   750
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
      TabIndex        =   9
      Top             =   120
      Width           =   1080
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
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1185
   End
   Begin VB.Line Line1 
      X1              =   3720
      X2              =   3720
      Y1              =   0
      Y2              =   2520
   End
End
Attribute VB_Name = "FrmRelatorioCTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Erro GoTo errosaida

Dim Rel As New CrysRelCRT
Dim Rs As ADODB.Recordset
Dim StrSql As String
Dim StrWhere As String

If IsDate(Datai.Text) Then
  If IsDate(Dataf.Text) Then
     StrWhere = "Emissao between '" & Format(Datai.Text, "yyyy-mm-dd") & "' and '" & Format(Dataf.Text, "yyyy-mm-dd") & "'"
  Else
    StrWhere = "Emissao ='" & Format(Datai.Text, "yyyy-mm-dd") & "'"
  End If

End If
StrSql = "Select * from nfentrada_cte "
If Len(StrWhere) > 0 Then
   StrSql = StrSql & " where " & StrWhere
End If
LcCap = Me.Caption
Screen.MousePointer = vbHourGlassThis
Me.Caption = "Aguarde, Gerando o Relatório..."

Set Rs = AbreRecordset(StrSql, True)
Load Relatorios
    With Relatorios
         Rel.DiscardSavedData
         Rel.Database.SetDataSource Rs
         .CRViewer1.ReportSource = Rel
    End With
 'setaformula
Relatorios.CRViewer1.ViewReport
    Relatorios.Show
errosaida:
Screen.MousePointer = vbDefault
Me.Caption = LcCap
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
