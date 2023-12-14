VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Sintegra 
   BackColor       =   &H00DDCFBF&
   Caption         =   "Sintegra"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10245
   ClipControls    =   0   'False
   Icon            =   "Sintegra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   10245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdProcessar 
      Caption         =   "&Processar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   31
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton CmdSintegra 
      Caption         =   "&Gera Sintegra"
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
      Height          =   495
      Left            =   8415
      TabIndex        =   30
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CheckBox Inventario 
      BackColor       =   &H00DDCFBF&
      Caption         =   "Incluir Inventario de Estoque"
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
      Left            =   120
      TabIndex        =   29
      Top             =   840
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDCFBF&
      Caption         =   "Considerações sobre o Inventário"
      Height          =   1455
      Left            =   120
      TabIndex        =   28
      Top             =   1080
      Width           =   6495
      Begin VB.CheckBox IncluirZero 
         BackColor       =   &H00DDCFBF&
         Caption         =   "Incluir Produto com Quantidade Zero"
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
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   3495
      End
      Begin VB.OptionButton OPtSintegraAnterior 
         BackColor       =   &H00DDCFBF&
         Caption         =   "Apartir do Sintegra Anterior"
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
         Left            =   2880
         TabIndex        =   35
         Top             =   480
         Width           =   2895
      End
      Begin VB.OptionButton OptApartir 
         BackColor       =   &H00DDCFBF&
         Caption         =   "Apartir de "
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
         TabIndex        =   33
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CheckBox OptArquivo 
         BackColor       =   &H00DDCFBF&
         Caption         =   "Incluir dados do Arquivo Sintegra"
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
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   3255
      End
      Begin MSMask.MaskEdBox DataInventario 
         Height          =   255
         Left            =   1440
         TabIndex        =   34
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   8
         Mask            =   "99/99/99"
         PromptChar      =   " "
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDCFBF&
      Caption         =   "FINALIDADES DA APRESENTAÇÃO DO ARQUIVO MAGNÉTICO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   5895
      Begin VB.ComboBox Finalidade 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   5535
      End
   End
   Begin MSMask.MaskEdBox Dataf 
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   5
      TabHeight       =   520
      BackColor       =   14536639
      TabCaption(0)   =   "Entradas"
      TabPicture(0)   =   "Sintegra.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Entrada"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "IcmsEntrada"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TotalEntrada"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Saídas"
      TabPicture(1)   =   "Sintegra.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(2)=   "Saida"
      Tab(1).Control(3)=   "IcmsSaida"
      Tab(1).Control(4)=   "TotalSaida"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Cupons Emitidos"
      TabPicture(2)   =   "Sintegra.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label10"
      Tab(2).Control(1)=   "Label2(1)"
      Tab(2).Control(2)=   "Cupom"
      Tab(2).Control(3)=   "IcmsCupom"
      Tab(2).Control(4)=   "TotalCupom"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Devoluções de Clientes"
      TabPicture(3)   =   "Sintegra.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label5"
      Tab(3).Control(1)=   "Label4"
      Tab(3).Control(2)=   "DevCliente"
      Tab(3).Control(3)=   "IcmsCliente"
      Tab(3).Control(4)=   "TotalCliente"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Resumo Sintegra."
      TabPicture(4)   =   "Sintegra.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Resumo"
      Tab(4).Control(1)=   "ResumoCfop"
      Tab(4).ControlCount=   2
      Begin MSFlexGridLib.MSFlexGrid ResumoCfop 
         Height          =   2415
         Left            =   -74760
         TabIndex        =   27
         Top             =   3240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4260
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedCols       =   0
      End
      Begin MSFlexGridLib.MSFlexGrid Resumo 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   26
         Top             =   1200
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   3201
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
      End
      Begin VB.TextBox TotalCliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -67080
         TabIndex        =   13
         Top             =   5640
         Width           =   1935
      End
      Begin VB.TextBox IcmsCliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -69960
         TabIndex        =   12
         Top             =   5640
         Width           =   1935
      End
      Begin VB.TextBox TotalEntrada 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7920
         TabIndex        =   11
         Top             =   5640
         Width           =   1935
      End
      Begin VB.TextBox IcmsEntrada 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5040
         TabIndex        =   10
         Top             =   5640
         Width           =   1935
      End
      Begin VB.TextBox TotalSaida 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -67080
         TabIndex        =   9
         Top             =   5640
         Width           =   1935
      End
      Begin VB.TextBox IcmsSaida 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -69960
         TabIndex        =   8
         Top             =   5640
         Width           =   1935
      End
      Begin VB.TextBox TotalCupom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -67080
         TabIndex        =   7
         Top             =   5640
         Width           =   1935
      End
      Begin VB.TextBox IcmsCupom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -69960
         TabIndex        =   6
         Top             =   5640
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid Saida 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   14
         Top             =   840
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   8281
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedCols       =   0
         ScrollBars      =   2
      End
      Begin MSFlexGridLib.MSFlexGrid Entrada 
         Height          =   4575
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   8070
         _Version        =   393216
         Rows            =   1
         Cols            =   9
         FixedCols       =   0
         ScrollBars      =   2
      End
      Begin MSFlexGridLib.MSFlexGrid Cupom 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   16
         Top             =   840
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   8281
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedCols       =   0
         ScrollBars      =   2
      End
      Begin MSFlexGridLib.MSFlexGrid DevCliente 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   17
         Top             =   960
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   8070
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedCols       =   0
         ScrollBars      =   2
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total "
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
         Left            =   -67800
         TabIndex        =   25
         Top             =   5760
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Icms"
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
         Left            =   -70920
         TabIndex        =   24
         Top             =   5760
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total "
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
         Left            =   7200
         TabIndex        =   23
         Top             =   5760
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Icms"
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
         Left            =   4080
         TabIndex        =   22
         Top             =   5760
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total "
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
         Left            =   -67800
         TabIndex        =   21
         Top             =   5760
         Width           =   510
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Icms"
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
         Left            =   -70920
         TabIndex        =   20
         Top             =   5760
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total "
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
         Left            =   -67800
         TabIndex        =   19
         Top             =   5760
         Width           =   510
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Icms"
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
         Left            =   -70920
         TabIndex        =   18
         Top             =   5760
         Width           =   900
      End
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo para Escituração do Sintegra"
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
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "Sintegra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private CSintegra As New BhForte
Private LcLocalSintegra As String
Sub gerar()
Call CmdProcessar_Click
Call CmdSintegra_Click
End Sub

Private Sub Check1_Click()

End Sub

Private Sub CmdProcessar_Click()
Dim LcCap As String
Dim StErros() As String
Dim a As Integer

'==> valida a digitacao
'==> Valida as datas
If Not IsDate(Datai.Text) Then
   MsgBox "A data de inicial não é valida.", 64, "Aviso"
   Datai.SetFocus
   SendKeys "{Home}+{end}"
   Exit Sub
End If
If Not IsDate(Dataf.Text) Then
   MsgBox "A data de final não é valida.", 64, "Aviso"
   Dataf.SetFocus
   SendKeys "{Home}+{end}"
   Exit Sub
End If

If CDate(Dataf.Text) < CDate(Datai.Text) Then
   MsgBox "O periodo informado não é valido.", 64, "Aviso"
   Datai.SetFocus
   SendKeys "{Home}+{end}"
   Exit Sub
End If
LcCap = Me.Caption
Csintegra.BuscaDados
Me.Caption = LcCap

If Len(Csintegra.erroS) = 0 Then
   Csintegra.ProcessarSintegra
   Csintegra.GerarResumo
   CmdSintegra.Enabled = True
Else
   StErros = Split(Csintegra.erroS, Chr(13))
   Load ErrosEncontrados
   For a = 0 To UBound(StErros)
       ErrosEncontrados.Erro.AddItem StErros(a)
   Next
   ErrosEncontrados.Show , Me
End If
Me.Caption = LcCap
End Sub

Private Sub CmdSintegra_Click()
LcCap = Me.Caption
Csintegra.EscreverSintegra
Me.Caption = LcCap
MsgBox "Aquivo gerando: " & Csintegra.LocalArmazenar, 64, "Aviso"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{Tab}"
   SendKeys "{Home}+{End}"
End If
End Sub

Private Sub Form_Load()
GeraGrid
LocalTemp
VerificaArquivoSisntegra
Finalidade.AddItem "1- Normal."
Finalidade.AddItem "2- Retificação total de arquivo: substituição total de informações prestadas pelo contribuinte referentes a este período."
Finalidade.AddItem "3- Retificação aditiva de arquivo: acréscimo de informação não incluída em arquivos já apresentados."
Finalidade.AddItem "5- Desfazimento: arquivo de informação referente a operações/prestações não efetivadas."
Finalidade.Text = "1- Normal."

End Sub
Sub VerificaArquivoSisntegra()
Dim Rs As adodb.Recordset
Dim StrSql As String
StrSql = "Select * from inventariosintegra LIMIT 1"
Set Rs = AbreRecordset(StrSql, True)
'MsgBox DEscricaoErro
If Not Rs.EOF Then
   OptArquivo.Enabled = True
Else
   OptArquivo.Enabled = False
   
End If
End Sub
Sub LocalTemp()
Dim Arquivo As String
Dim LocalPdv As String
'Dim Glbase As String
If Len(GLBase) = 0 Then GLBase = AcessoAdo.LocalBaseDados
Dim a As Integer
For a = Len(GLBase) To 1 Step -1
    If Mid(GLBase, a, 1) = "\" Then
       Exit For
    End If
Next
Arquivo = Mid(GLBase, 1, a)
Arquivo = Arquivo & "Configuracaopdv.txt"
LcLocalSintegra = LeIni("Pdv", "Sintegra", Arquivo)
End Sub
Sub GeraGrid()
On Error Resume Next
Entrada.Cols = 17
Saida.Cols = 15
DevFornecedores.Cols = 12
Cupom.Cols = 12
DevCliente.Cols = 12

Entrada.TextMatrix(0, 0) = "Nº Nota"
Entrada.TextMatrix(0, 1) = "Entrada"
Entrada.TextMatrix(0, 2) = "Fornecedor"
Entrada.TextMatrix(0, 3) = "Valor"
Entrada.TextMatrix(0, 4) = "ICMS"
Entrada.TextMatrix(0, 5) = "Valor ICMS"
Entrada.TextMatrix(0, 6) = "CodFornecedor"
Entrada.TextMatrix(0, 7) = "Modelo"
Entrada.TextMatrix(0, 8) = "Serie"
Entrada.TextMatrix(0, 9) = "Cfop" '==> Codigo Fiscal
Entrada.TextMatrix(0, 10) = "Situacao"
Entrada.TextMatrix(0, 11) = "Codigo"
Entrada.TextMatrix(0, 12) = "outras"
Entrada.TextMatrix(0, 13) = "Codigo"
Entrada.TextMatrix(0, 14) = "Valor total Ipi"
Entrada.TextMatrix(0, 15) = "SubSerie"
Entrada.TextMatrix(0, 16) = "TipoFrete"

Entrada.ColWidth(0) = 1000
Entrada.ColWidth(1) = 1000
Entrada.ColWidth(2) = 4200
Entrada.ColWidth(3) = 1000
Entrada.ColWidth(4) = 1000
Entrada.ColWidth(5) = 1000
Entrada.ColWidth(6) = 0
Entrada.ColWidth(7) = 0
Entrada.ColWidth(8) = 0
Entrada.ColWidth(9) = 0
Entrada.ColWidth(10) = 0
Entrada.ColWidth(11) = 0
Entrada.ColWidth(12) = 1000
Entrada.ColWidth(13) = 0
Entrada.ColWidth(14) = 0
Entrada.ColWidth(15) = 0
Entrada.ColWidth(16) = 0

Saida.TextMatrix(0, 0) = "Nº Nota"
Saida.TextMatrix(0, 1) = "Saida"
Saida.TextMatrix(0, 2) = "Cliente"
Saida.TextMatrix(0, 3) = "Valor"
Saida.TextMatrix(0, 4) = "ICMS"
Saida.TextMatrix(0, 5) = "Valor ICMS"
Saida.TextMatrix(0, 6) = "CodFornecedor"
Saida.TextMatrix(0, 7) = "Modelo"
Saida.TextMatrix(0, 8) = "Serie"
Saida.TextMatrix(0, 9) = "Cfop" '==> Codigo Fiscal
Saida.TextMatrix(0, 10) = "Situacao"
Saida.TextMatrix(0, 11) = "Codigo"
Saida.TextMatrix(0, 12) = "Desconto"
Saida.TextMatrix(0, 13) = "Acrescimo"
Saida.TextMatrix(0, 14) = "Ipi"

Saida.ColWidth(0) = 1000
Saida.ColWidth(1) = 1000
Saida.ColWidth(2) = 4200
Saida.ColWidth(3) = 1000
Saida.ColWidth(4) = 1000
Saida.ColWidth(5) = 1000
Saida.ColWidth(6) = 0
Saida.ColWidth(7) = 0
Saida.ColWidth(8) = 0
Saida.ColWidth(9) = 0
Saida.ColWidth(10) = 0
Saida.ColWidth(11) = 0
Saida.ColWidth(12) = 0
Saida.ColWidth(13) = 0
Saida.ColWidth(14) = 0

Cupom.TextMatrix(0, 0) = "Nº Nota"
Cupom.TextMatrix(0, 1) = "Emissão"
Cupom.TextMatrix(0, 2) = "Cliente"
Cupom.TextMatrix(0, 3) = "Valor"
Cupom.TextMatrix(0, 4) = "ICMS"
Cupom.TextMatrix(0, 5) = "Valor ICMS"
Cupom.TextMatrix(0, 6) = "CodFornecedor"
Cupom.TextMatrix(0, 7) = "Modelo"
Cupom.TextMatrix(0, 8) = "Serie"
Cupom.TextMatrix(0, 9) = "Cfop" '==> Codigo Fiscal
Cupom.TextMatrix(0, 10) = "Situacao"
Cupom.TextMatrix(0, 11) = "Codigo"


Cupom.ColWidth(0) = 1000
Cupom.ColWidth(1) = 1000
Cupom.ColWidth(2) = 4200
Cupom.ColWidth(3) = 1000
Cupom.ColWidth(4) = 1000
Cupom.ColWidth(5) = 1000
Cupom.ColWidth(6) = 0
Cupom.ColWidth(7) = 0
Cupom.ColWidth(8) = 0
Cupom.ColWidth(9) = 0
Cupom.ColWidth(10) = 0
Cupom.ColWidth(11) = 0

DevCliente.TextMatrix(0, 0) = "Nº Nota"
DevCliente.TextMatrix(0, 1) = "Entrada"
DevCliente.TextMatrix(0, 2) = "Cliente"
DevCliente.TextMatrix(0, 3) = "Valor"
DevCliente.TextMatrix(0, 4) = "ICMS"
DevCliente.TextMatrix(0, 5) = "Valor ICMS"
DevCliente.TextMatrix(0, 6) = "CodFornecedor"
DevCliente.TextMatrix(0, 7) = "Modelo"
DevCliente.TextMatrix(0, 8) = "Serie"
DevCliente.TextMatrix(0, 9) = "Cfop" '==> Codigo Fiscal
DevCliente.TextMatrix(0, 10) = "Situacao"
DevCliente.TextMatrix(0, 11) = "Codigo"

DevCliente.ColWidth(0) = 1000
DevCliente.ColWidth(1) = 1000
DevCliente.ColWidth(2) = 4200
DevCliente.ColWidth(3) = 1000
DevCliente.ColWidth(4) = 1000
DevCliente.ColWidth(5) = 1000
DevCliente.ColWidth(6) = 0
DevCliente.ColWidth(7) = 0
DevCliente.ColWidth(8) = 0
DevCliente.ColWidth(9) = 0
DevCliente.ColWidth(10) = 0
DevCliente.ColWidth(11) = 0

Resumo.TextMatrix(0, 0) = "Tipo"
Resumo.TextMatrix(0, 1) = "Valor Contabil"
Resumo.TextMatrix(0, 2) = "Base Cálculo"
Resumo.TextMatrix(0, 3) = "Icms"

Resumo.ColWidth(0) = 3000
Resumo.ColWidth(1) = 2000
Resumo.ColWidth(2) = 2000
Resumo.ColWidth(3) = 2000

ResumoCfop.TextMatrix(0, 0) = "CFOP"
ResumoCfop.TextMatrix(0, 1) = "Valor Total"
ResumoCfop.TextMatrix(0, 2) = "Base Calc."
ResumoCfop.TextMatrix(0, 3) = "Valor Icms"
ResumoCfop.TextMatrix(0, 4) = "Valor Isentas"
ResumoCfop.TextMatrix(0, 5) = "Valor Outras"

ResumoCfop.ColWidth(0) = 1000
ResumoCfop.ColWidth(1) = 1000
ResumoCfop.ColWidth(2) = 1000
ResumoCfop.ColWidth(3) = 2000
ResumoCfop.ColWidth(4) = 2000
ResumoCfop.ColWidth(5) = 2000

End Sub

Private Sub Inventario_Click()
Frame1.Enabled = Inventario.Value * -1
DataInventario.Enabled = Inventario.Value * -1
End Sub

Private Sub Label3_Click()

End Sub

Private Sub OptApartir_Click()
DataInventario.Enabled = OptApartir.Value
   
End Sub

Private Sub OptArquivo_Click()
DataInventario.Enabled = Not OptArquivo.Value
End Sub

Private Sub OPtSintegraAnterior_Click()
DataInventario.Enabled = Not OPtSintegraAnterior.Value
End Sub
