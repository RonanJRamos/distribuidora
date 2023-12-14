VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form configuracoes 
   BackColor       =   &H00DDF2FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurações Para Acesso ao Banco de Dados"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox senhasql 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   3120
      Width           =   2895
   End
   Begin VB.TextBox usuariosql 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox bancosql 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   2535
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6000
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar botoes 
      Height          =   630
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1111
      ButtonWidth     =   1931
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar F2"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Fechar F12"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox ServidorMySql 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox BaseAcess 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   6135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDF2FF&
      Caption         =   "Tipo de Banco de Dados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      Begin VB.OptionButton Option3 
         BackColor       =   &H00DDF2FF&
         Caption         =   "Sql Server"
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
         Left            =   3240
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00DDF2FF&
         Caption         =   "MySql"
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
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00DDF2FF&
         Caption         =   "Ms Access"
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
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha de Acesso ao Sql"
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
      Index           =   3
      Left            =   3480
      TabIndex        =   14
      Top             =   2880
      Width           =   2100
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario de Acesso ao Sql"
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
      Index           =   2
      Left            =   3480
      TabIndex        =   13
      Top             =   2280
      Width           =   2205
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Banco de Dados"
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
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   2235
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Servidor"
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
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Base de Dados"
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
      Top             =   1320
      Width           =   1305
   End
End
Attribute VB_Name = "configuracoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub botoes_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case UCase(Trim(Button))
    Case Is = "&FECHAR F12"
        Unload Me
    Case Is = "&SALVAR F2"
       Salvaconfig
       'FechaConexao
       'abreconexao
       Unload configuracoes
End Select

End Sub
Function Salvaconfig()
Dim strValue As String
Dim Arqini   As String
Arqini = BuscaDirWin
Arqini = Arqini & LcPAth & "\" & App.EXEName & ".ini"


If Option1.Value Then strValue = "Access"
If Option2.Value Then strValue = "MySql"
If Option3.Value Then strValue = "SqlServer"

GravaIni "Base de Dados", "tipo de banco", strValue, Arqini
GravaIni "Base de Dados", "Servidor Mysql", CStr(ServidorMySql.Text), Arqini
GravaIni "Base de Dados", "BaseAcess", CStr(BaseAcess.Text), Arqini

GravaIni "Base de Dados", "nomebancosql", bancosql.Text, Arqini
GravaIni "Base de Dados", "usuariosql", usuariosql.Text, Arqini
GravaIni "Base de Dados", "senhasql", senhasql.Text, Arqini

End Function

Private Sub Command1_Click()
On Error Resume Next
Dim LcObj As CommonDialog
Set LcObj = objetos.dialogo
With LcObj
     .InitDir = App.Path
     .FileName = "*.mdb"
     .ShowOpen
     BaseAcess.Text = .FileName & ""
End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then Call botoes_ButtonClick(botoes.Buttons(1))
If KeyCode = 123 Then Call botoes_ButtonClick(botoes.Buttons(2))
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Dim Arqini As String
Dim Base As String
Arqini = BuscaDirWin
Arqini = Arqini & LcPAth & "\" & App.EXEName & ".ini"
If Dir(Arqini, vbArchive) <> "" Then
   'Option1.Enabled = False
   'Option2.Enabled = False
   'Frame1.Enabled = False
   Base = LeIni("Base de Dados", "tipo de banco", Arqini)
   If UCase(Base) = "ACCESS" Then Option1.Value = True
   If UCase(Base) = "MYSQL" Then Option2.Value = True
   If UCase(Base) = "SQLSERVER" Then Option3.Value = True
   
   ServidorMySql.Text = LeIni("Base de Dados", "Servidor MySql", Arqini)
   BaseAcess.Text = LeIni("base de dados", "BaseAcess", Arqini)
   bancosql.Text = LeIni("base de dados", "nomebancosql", Arqini)
   usuariosql.Text = LeIni("base de dados", "usuariosql", Arqini)
   senhasql.Text = LeIni("base de dados", "senhasql", Arqini)


End If


setacamposuteis

End Sub
Sub setacamposuteis()
On Error Resume Next
If Option1.Value Then
   Label1.Visible = True
   BaseAcess.Visible = True
   Label2(0).Visible = False
   Label2(1).Visible = False
   Label2(2).Visible = False
   Label2(3).Visible = False

   ServidorMySql.Visible = False
   Command1.Visible = True
Else
   Label1.Visible = False
   BaseAcess.Visible = False
   Label2(0).Visible = True
   Label2(1).Visible = True
   Label2(2).Visible = True
   Label2(3).Visible = True

   ServidorMySql.Visible = True
   Command1.Visible = False
End If
End Sub

Private Sub Option1_Click()
setacamposuteis
End Sub

Private Sub Option2_Click()
setacamposuteis
End Sub

Private Sub Option3_Click()
setacamposuteis

End Sub
