VERSION 5.00
Begin VB.Form FrmFuncionarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Funcionárioss"
   ClientHeight    =   7440
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10605
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      Height          =   1605
      Index           =   22
      Left            =   960
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Tag             =   "S/T/N/22/N/OBSERVACAO"
      Top             =   5040
      Width           =   9495
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   21
      Left            =   6240
      MaxLength       =   10
      TabIndex        =   21
      Tag             =   "S/D/N/21/N/DEMISSAO"
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   20
      Left            =   960
      MaxLength       =   10
      TabIndex        =   20
      Tag             =   "S/D/N/20/N/ADMISSAO"
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   19
      Left            =   6240
      MaxLength       =   10
      TabIndex        =   19
      Tag             =   "S/M/N/19/N/COMISSAO"
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   18
      Left            =   960
      TabIndex        =   18
      Tag             =   "S/M/N/18/N/SALARIO"
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   17
      Left            =   7800
      MaxLength       =   50
      TabIndex        =   17
      Tag             =   "S/T/N/17/N/HORARIO"
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   16
      Left            =   960
      MaxLength       =   100
      TabIndex        =   16
      Tag             =   "S/T/N/16/N/FUNCAO"
      Top             =   3600
      Width           =   5895
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   15
      Left            =   6480
      MaxLength       =   50
      TabIndex        =   15
      Tag             =   "S/T/N/15/N/MAE"
      Top             =   3120
      Width           =   3975
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   14
      Left            =   720
      MaxLength       =   50
      TabIndex        =   14
      Tag             =   "S/T/N/14/N/PAI"
      Top             =   3120
      Width           =   5055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   13
      Left            =   7800
      MaxLength       =   40
      TabIndex        =   13
      Tag             =   "S/T/N/13/N/CARTEIRATRABALHO"
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   12
      Left            =   4200
      MaxLength       =   30
      TabIndex        =   12
      Tag             =   "S/T/N/12/N/RG"
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   11
      Left            =   960
      MaxLength       =   24
      TabIndex        =   11
      Tag             =   "S/T/N/11/N/CPF"
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   10
      Left            =   5640
      MaxLength       =   24
      TabIndex        =   10
      Tag             =   "S/T/N/10/N/FONE"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   9
      Left            =   2640
      MaxLength       =   15
      TabIndex        =   9
      Tag             =   "S/T/N/09/N/CEP"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   8
      Left            =   840
      MaxLength       =   3
      TabIndex        =   8
      Tag             =   "S/T/N/08/N/ESTADO"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   7
      Left            =   8520
      MaxLength       =   150
      TabIndex        =   7
      Tag             =   "S/T/N/07/N/CIDADE"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   6
      Left            =   4560
      MaxLength       =   150
      TabIndex        =   6
      Tag             =   "S/T/N/06/N/BAIRRO"
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   5
      Left            =   840
      MaxLength       =   150
      TabIndex        =   5
      Tag             =   "S/T/N/05/N/COMPLEMENTO"
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   4
      Left            =   8280
      MaxLength       =   50
      TabIndex        =   4
      Tag             =   "S/T/N/04/N/NUMERO"
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   3
      Left            =   840
      MaxLength       =   150
      TabIndex        =   3
      Tag             =   "S/T/N/03/N/RUA"
      Top             =   1440
      Width           =   6615
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   8280
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "S/D/N/02/N/NASCIMENTO"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   840
      MaxLength       =   80
      TabIndex        =   1
      Tag             =   "S/T/S/01/N/NOME"
      Top             =   1080
      Width           =   6615
   End
   Begin VB.TextBox txt 
      Height          =   405
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "S/N/S/00/N/CODIGO"
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton CmdUltimo 
      Caption         =   "&Ultimo"
      Height          =   375
      Left            =   8970
      TabIndex        =   38
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3720
      Top             =   360
   End
   Begin VB.TextBox DataS 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   35
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox HoraS 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   34
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CmdSeguinte 
      Caption         =   "&Seguinte"
      Height          =   375
      Left            =   7515
      TabIndex        =   33
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "&Anterior"
      Height          =   375
      Left            =   6060
      TabIndex        =   32
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir"
      Height          =   375
      Left            =   3150
      TabIndex        =   31
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton CmdPrimeiro 
      Caption         =   "&Primeiro"
      Height          =   375
      Left            =   4605
      TabIndex        =   30
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1695
      TabIndex        =   29
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   8640
      Width           =   2895
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000005&
      X1              =   4680
      X2              =   4680
      Y1              =   3960
      Y2              =   4920
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   10560
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000005&
      X1              =   10560
      X2              =   0
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   10560
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Carteira Tr."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6600
      TabIndex        =   58
      Top             =   2640
      Width           =   1080
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fone"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5040
      TabIndex        =   36
      Top             =   2160
      Width           =   480
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   11520
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Obs.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   57
      Top             =   5040
      Width           =   480
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Demissão"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5280
      TabIndex        =   54
      Top             =   4560
      Width           =   915
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comissão"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5280
      TabIndex        =   56
      Top             =   4080
      Width           =   915
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salário"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   55
      Top             =   4080
      Width           =   690
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admissão"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   -240
      TabIndex        =   53
      Top             =   4320
      Width           =   135
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aniv."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7560
      TabIndex        =   27
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Horário"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6960
      TabIndex        =   51
      Top             =   3600
      Width           =   705
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admissão"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   52
      Top             =   4560
      Width           =   915
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   11400
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mãe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5880
      TabIndex        =   26
      Top             =   3120
      Width           =   405
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3720
      TabIndex        =   25
      Top             =   8640
      Width           =   675
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Função"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   50
      Top             =   3600
      Width           =   705
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   11400
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   11400
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Uf"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   43
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.E.P.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1920
      TabIndex        =   45
      Top             =   2160
      Width           =   630
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.P.F.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   44
      Top             =   2640
      Width           =   630
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3720
      TabIndex        =   23
      Top             =   2640
      Width           =   285
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pai"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   49
      Top             =   3120
      Width           =   315
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   41
      Top             =   1440
      Width           =   420
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7560
      TabIndex        =   47
      Top             =   1440
      Width           =   210
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compl"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   42
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      TabIndex        =   46
      Top             =   1800
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7560
      TabIndex        =   48
      Top             =   1800
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   40
      Top             =   1080
      Width           =   675
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   11400
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   39
      Top             =   480
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   " Controle de Funcionários"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   0
      TabIndex        =   37
      Top             =   120
      Width           =   10875
   End
   Begin VB.Menu MnArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu MnSair 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu MnRegistro 
      Caption         =   "&Registro"
      Begin VB.Menu MnSalvar 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnPesquisar 
         Caption         =   "&Pesquisar"
      End
      Begin VB.Menu MnOrdenar 
         Caption         =   "&Ordenar"
      End
      Begin VB.Menu MnExcluir 
         Caption         =   "&Excluir"
      End
   End
   Begin VB.Menu MnMovimento 
      Caption         =   "&Movimentar"
      Begin VB.Menu MnPrimeiro 
         Caption         =   "&Primeiro"
      End
      Begin VB.Menu MnAnterior 
         Caption         =   "&Anterior"
      End
      Begin VB.Menu MSeguinte 
         Caption         =   "&Seguinte"
      End
      Begin VB.Menu MnUltimo 
         Caption         =   "&Último"
      End
   End
   Begin VB.Menu MnPop 
      Caption         =   "&Pop"
      Visible         =   0   'False
      Begin VB.Menu PopPesquisar 
         Caption         =   "&Pesquisar"
      End
      Begin VB.Menu PopOrdenar 
         Caption         =   "&Ordenar"
      End
   End
End
Attribute VB_Name = "FrmFuncionarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcCarregado As Integer
Private Function DesabilitaTodos()
Dim a As Integer
For a = 0 To 30
    txt(a).Enabled = False
Next
End Function
Private Sub CmdAnterior_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enAnterior, Funcionario) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub CmdExcluir_Click()
On Error Resume Next
If Exclui(Funcionario) = 1 Then
      VinculaDados
End If
End Sub

Private Sub CmdFechar_Click()
Unload Me
End Sub

Private Sub CmdPrimeiro_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enPrimeiro, Funcionario) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub CmdSalvar_Click()
On Error Resume Next
Call SalvaRegistro(Funcionario)
LcRegAtual = True
VinculaDados
LcRegAtual = False

End Sub

Private Sub CmdSeguinte_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, Funcionario) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub CmdUltimo_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enultimo, Funcionario) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub



Private Sub Form_Activate()
On Error Resume Next
If LcCarregado Then Exit Sub
Select Case LcTipoDados
   Case Is = 1
        DesabilitaCtr
   Case Is = 2
      Call AbreBanco(Funcionario)
      VinculaDados
   Case Is = 3
      'DesabilitaTodos
      MnSalvar.Enabled = False
      MnExcluir.Enabled = False
      Call AbreBanco(Funcionario)
      CmdExcluir.Enabled = False
      VinculaDados
 End Select
'CriaMascara
LcRegAtual = False
FrmPrincipal.Visible = False
CarreGamatriz
LcCarregado = True

End Sub
Function CarreGamatriz()
Dim a As Integer, LcNome As String, LcTipo As String
GlFormAtual = Tabela.Funcionario

Set GlFormA = Me
For a = 0 To 30
    LcNome = Mid$(txt(a).Tag, 12)
    LcTipo = Mid$(txt(a).Tag, 3, 1)
    If Err <> 0 Then Exit For
    MtPesquisa(a).Indice = LcNome
    MtPesquisa(a).tipo = LcTipo
    
    Select Case LcNome
           Case Is = "CARTEIRATRABALHO"
                MtPesquisa(a).Campo = "CARTEIRA DE TRABALHO"
           Case Is = "FUNCAO"
                MtPesquisa(a).Campo = "FUNÇÃO"
           Case Is = "COMISSAO"
                MtPesquisa(a).Campo = "COMISSÃO"
           Case Is = "ADMISSAO"
                MtPesquisa(a).Campo = "ADMISSÃO"
           Case Is = "DEMISSAO"
                MtPesquisa(a).Campo = "DEMISSÃO"
           Case Is = "OBSERVACAO"
                MtPesquisa(a).Campo = "OBSERVAÇÃO"
           Case Else
                MtPesquisa(a).Campo = LcNome
      End Select
 Next
 
End Function
Private Sub Form_Load()
On Error Resume Next
Me.Height = 8565
Me.Width = 10695
DataS.Text = Format(GlDataSistema, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = 0
Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
 
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

If (LcTipoDados = 1) And (CmdSalvar.Enabled = True) Then
   GlPergunta = True
   SalvaRegistro (Cliente)
End If
If (LcTipoDados = 2) And LcAlterado Then SalvaRegistro (Cliente)
FechaBanco
FrmPrincipal.Visible = True
LcCarregado = False
End Sub

Private Sub MnAnterior_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enAnterior, Funcionario) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MnExcluir_Click()
On Error Resume Next
If Exclui(Funcionario) = 1 Then
      VinculaDados
End If
LcRegAtual = False
End Sub

Private Sub MnOrdenar_Click()
On Error Resume Next
FrmOrdena.Show , Me
End Sub

Private Sub MnPesquisar_Click()
On Error Resume Next
frmPesquisa.Show , Me
LcRegAtual = False
End Sub

Private Sub MnPrimeiro_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enPrimeiro, Funcionario) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MnSair_Click()
Unload Me
End Sub

Private Sub MnSalvar_Click()
Call SalvaRegistro(Funcionario)
LcRegAtual = True
VinculaDados
LcRegAtual = False

End Sub

Private Sub MnUltimo_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, Funcionario) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MSeguinte_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, Funcionario) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
HoraS.Text = Format(Time, "hh:mm:ss")
End Sub
Private Function DesabilitaCtr()
CmdPrimeiro.Enabled = False
CmdAnterior.Enabled = False
CmdUltimo.Enabled = False
CmdSeguinte.Enabled = False
MnMovimento.Enabled = False
MnRegistro.Enabled = False
CmdExcluir.Enabled = False
End Function
Function VinculaDados()
On Error Resume Next

If LcTipoDados = 1 Then NovoReg Else Call RegistroAtual(Funcionario)


txt(0).Text = GlCampo0
txt(1).Text = GlCampo1
txt(2).Text = GlCampo2
txt(3).Text = GlCampo3
txt(4).Text = GlCampo4
txt(5).Text = GlCampo5
txt(6).Text = GlCampo6
txt(7).Text = GlCampo7
txt(8).Text = GlCampo8
txt(9).Text = GlCampo9
txt(10).Text = GlCampo10
txt(11).Text = GlCampo11
txt(12).Text = GlCampo12
txt(13).Text = GlCampo13
txt(14).Text = GlCampo14
txt(15).Text = GlCampo15
txt(16).Text = GlCampo16
txt(17).Text = GlCampo17
txt(18).Text = GlCampo18
txt(19).Text = GlCampo19
txt(20).Text = GlCampo20
txt(21).Text = GlCampo21
txt(22).Text = GlCampo22

txt(1).SetFocus

CmdSalvar.Enabled = False
MnSalvar.Enabled = False

Exit Function
ErroVinculo:
Resume Next
End Function

Private Sub Txt_Change(Index As Integer)
Call Alterado

End Sub


Private Sub txt_GotFocus(Index As Integer)

If VerificaDuplicado(Index) Then
   txt(Index).SetFocus
End If

End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
Call MoveTecla(Index, KeyCode)
End Sub
