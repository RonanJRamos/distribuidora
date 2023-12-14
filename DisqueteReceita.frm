VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form DisqueteReceita 
   BackColor       =   &H00DDF2FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gera Disquete Para Receita Federal"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Sair"
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Gerar Disquete"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDF2FF&
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   2415
      Begin MSMask.MaskEdBox Dataf 
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99/99/99"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox datai 
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99/99/99"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
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
         Left            =   720
         TabIndex        =   7
         Top             =   960
         Width           =   150
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDF2FF&
      Caption         =   "Tipo de Geração"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      Begin VB.OptionButton Option3 
         BackColor       =   &H00DDF2FF&
         Caption         =   "Nota de Serviços"
         Enabled         =   0   'False
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
         TabIndex        =   3
         Top             =   1320
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00DDF2FF&
         Caption         =   "Substituição Tributária"
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
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00DDF2FF&
         Caption         =   "Venda Consumidor Final"
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
         Width           =   2655
      End
   End
End
Attribute VB_Name = "DisqueteReceita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If datai.Text = "  /  /  " Then datai.Text = Format(Date, "dd/mm/yy")
If Dataf.Text = "  /  /  " Then Dataf.Text = Format(Date, "dd/mm/yy")
Call VerificaDisquete("a:")
Call GeraDisquete("a:")
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

