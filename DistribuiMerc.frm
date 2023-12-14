VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form DistribuiMerc 
   BackColor       =   &H00C5FEE1&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distribui Mercadorias"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Fechar 
      Caption         =   "&Fechar F10"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Confirma 
      Caption         =   "&Confirma F2"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin MSMask.MaskEdBox quantidade 
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox quantidade 
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   1
      Top             =   1845
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox quantidade 
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   2
      Top             =   2355
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox quantidade 
      Height          =   375
      Index           =   3
      Left            =   3960
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox quantidade 
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   11
      Top             =   3405
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox quantidade 
      Height          =   375
      Index           =   5
      Left            =   3960
      TabIndex        =   13
      Top             =   3915
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox quantidade 
      Height          =   375
      Index           =   6
      Left            =   3960
      TabIndex        =   15
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox quantidade 
      Height          =   375
      Index           =   7
      Left            =   3960
      TabIndex        =   17
      Top             =   4965
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox quantidade 
      Height          =   375
      Index           =   8
      Left            =   3960
      TabIndex        =   19
      Top             =   5475
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox quantidade 
      Height          =   375
      Index           =   9
      Left            =   3960
      TabIndex        =   21
      Top             =   6000
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Produto 
      BackStyle       =   0  'Transparent
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
      Height          =   735
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Produto 
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   3855
   End
   Begin VB.Line Line1 
      X1              =   5640
      X2              =   5640
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Label galpao 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   20
      Top             =   6000
      Width           =   3615
   End
   Begin VB.Label galpao 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   18
      Top             =   5475
      Width           =   3615
   End
   Begin VB.Label galpao 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   16
      Top             =   4965
      Width           =   3615
   End
   Begin VB.Label galpao 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   4440
      Width           =   3615
   End
   Begin VB.Label galpao 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   3915
      Width           =   3615
   End
   Begin VB.Label galpao 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   3405
      Width           =   3615
   End
   Begin VB.Label galpao 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   3615
   End
   Begin VB.Label galpao 
      BackStyle       =   0  'Transparent
      Caption         =   "Santa Maria 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   2355
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label galpao 
      BackStyle       =   0  'Transparent
      Caption         =   "Santa Maria "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1845
      Width           =   3615
   End
   Begin VB.Label galpao 
      BackStyle       =   0  'Transparent
      Caption         =   "California"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   3615
   End
End
Attribute VB_Name = "DistribuiMerc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type Lcdadosgalpao
        Codigo As String
        Nome As String
End Type
Private mtgalpao() As Lcdadosgalpao
Private Lct, a As Long
Sub BuscaNomeGalpao()
Dim db As Database
Dim Rs As Recordset
Dim a As Integer

Set db = OpenDatabase(GLBase)
Set Rs = db.OpenRecordset("Select * from alid012 order by codigo")
Do Until Rs.EOF
  If a = 3 Then Exit Do
  If a = 0 Then
     Galpao(0).Caption = Rs!Nome
  End If
  If a = 1 Then
     Galpao(1).Caption = Rs!Nome
  End If
  If a = 2 Then
     Galpao(2).Caption = Rs!Nome
  End If
  a = a + 1
  Rs.MoveNext
Loop
Set db = Nothing
Set Rs = Nothing

End Sub
Private Sub Confirma_Click()
Dim Lcq     As Double
Dim LcAnt   As Double
On Error Resume Next
If Len(Quantidade(0).Text) = 0 Then Quantidade(0).Text = 0
If Len(Quantidade(1).Text) = 0 Then Quantidade(1).Text = 0
If Len(Quantidade(2).Text) = 0 Then Quantidade(2).Text = 0

If Len(Quantidade(2).Text) > 0 Then
   FrmEntradaProduto.california.Text = Quantidade(0).Text
Else
   Quantidade(0).Text = "0"
End If
If Len(Quantidade(0).Text) > 0 Then
   FrmEntradaProduto.santamaria.Text = Quantidade(1).Text
Else
   Quantidade(1).Text = "0"
End If
If Len(Quantidade(1).Text) > 0 Then
   FrmEntradaProduto.santamaria1.Text = Quantidade(2).Text
Else
   Quantidade(2).Text = "0"
End If
LcAnt = CDbl(FrmEntradaProduto.Txt(3).Text)
Lcq = CDbl(Quantidade(0).Text) + CDbl(CDbl(Quantidade(1).Text)) + CDbl(CDbl(Quantidade(2).Text))
If Lcq < LcAnt Then
   MsgBox "A Quantidade Distribuida é Inferior a Quantidade Comprada", 64, "Aviso"
   Quantidade(0).SetFocus
   Exit Sub
End If
If Lcq > LcAnt Then
   MsgBox "A Quantidade Distribuida é Superior a Quantidade Comprada.", 64, "Aviso"
   Quantidade(0).SetFocus
   Exit Sub
End If
   

Unload Me
Exit Sub


End Sub

Private Sub Confirma_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
'Quantidade(0).SetFocus
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Fechar_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Fechar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Quantidade(0).SetFocus
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{Home}+{End}"
End If
End Sub

Private Sub Form_Load()
On Error Resume Next

Dim RsGalpao As Recordset
Dim LcPrimeiro As String
BuscaNomeGalpao
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
Produto(0).Caption = "Disponível : " & FrmEntradaProduto.Txt(3).Text
Produto(1).Caption = "de " & FrmEntradaProduto.Txt(2).Text
Exit Sub
AbreBase
Set RsGalpao = Dbbase.OpenRecordset("select * From alid012")
If Not RsGalpao.EOF Then
   For a = 0 To 9
      ReDim Preserve mtgalpao(a)
      If RsGalpao.EOF Then Exit For
      Galpao(a).Caption = RsGalpao!Nome
      Galpao(a).Visible = True
      Quantidade(a).Visible = True
      mtgalpao(a).Codigo = RsGalpao!Codigo
      mtgalpao(a).Nome = RsGalpao!Nome
      RsGalpao.MoveNext
   Next
End If

If a > 0 Then Lct = a - 1
RsGalpao.Close
Set RsGalpao = Nothing
For a = 0 To 9
    Quantidade(a).Text = 0
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
GlLibera = True
FrmEntradaProduto.SetFocus
End Sub

Private Sub Quantidade_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub
