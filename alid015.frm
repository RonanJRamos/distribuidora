VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form alid015 
   BackColor       =   &H00DDF2FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Receitas"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Acrescimo 
      Height          =   285
      Left            =   4200
      TabIndex        =   7
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton botoes4 
      Caption         =   "..."
      Height          =   255
      Left            =   6600
      TabIndex        =   30
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox BotoesClienteAnterior 
      Height          =   285
      Left            =   7320
      TabIndex        =   29
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton botoes3 
      Caption         =   "..."
      Height          =   255
      Left            =   7920
      TabIndex        =   28
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox NossoNumero 
      Height          =   285
      Left            =   4920
      TabIndex        =   26
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox codigo 
      Height          =   375
      Left            =   5880
      TabIndex        =   25
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox nomereceita 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   3120
      Width           =   4575
   End
   Begin VB.TextBox nome 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   2040
      Width           =   6015
   End
   Begin VB.TextBox VALPAGO 
      Height          =   285
      Left            =   6720
      TabIndex        =   8
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox TPMONET 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Obs 
      Height          =   1005
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   4080
      Width           =   6855
   End
   Begin VB.TextBox VALOR 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   1590
      Width           =   2055
   End
   Begin VB.TextBox NF 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox CLIENTE 
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   2040
      Width           =   615
   End
   Begin MSComctlLib.Toolbar botoes1 
      Height          =   570
      Left            =   0
      TabIndex        =   14
      Top             =   5835
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   1005
      ButtonWidth     =   2064
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Primeiro F8"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "A&nterior F9"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Se&guinte F10"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Ultimo F11"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir Boleto"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Fechar F12"
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.Toolbar botoes 
      Height          =   570
      Left            =   0
      TabIndex        =   12
      Top             =   5175
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   1005
      ButtonWidth     =   1984
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Salvar F2"
            Key             =   "F2"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Incluir F3"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Consultar F4"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Alterar F4"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Excluir F5"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pes&quisar F6"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Ordenar F7"
            ImageIndex      =   14
         EndProperty
      EndProperty
      BorderStyle     =   1
      OLEDropMode     =   1
   End
   Begin MSComctlLib.StatusBar BarStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   6450
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox DTVENC 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox DATA 
      Height          =   285
      Left            =   5880
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSComctlLib.ImageList figuras 
      Left            =   3840
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alid015.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alid015.frx":019A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alid015.frx":0334
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alid015.frx":0446
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alid015.frx":0558
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alid015.frx":09AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alid015.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alid015.frx":0BCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alid015.frx":0CE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alid015.frx":1132
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alid015.frx":1584
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alid015.frx":1BFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alid015.frx":1D10
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alid015.frx":241E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox DTPAGTO 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Acrescimo"
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
      Index           =   8
      Left            =   3120
      TabIndex        =   31
      Top             =   3645
      Width           =   885
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Noso Numero"
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
      Index           =   7
      Left            =   3600
      TabIndex        =   27
      Top             =   1605
      Width           =   1155
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Pag"
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
      Index           =   6
      Left            =   5760
      TabIndex        =   24
      Top             =   3645
      Width           =   840
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DataPag"
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
      TabIndex        =   23
      Top             =   3645
      Width           =   750
   End
   Begin VB.Line Line 
      BorderWidth     =   2
      X1              =   -120
      X2              =   8520
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Titulo 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Monet"
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
      Index           =   14
      Left            =   120
      TabIndex        =   21
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Obs"
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
      Index           =   17
      Left            =   120
      TabIndex        =   20
      Top             =   4200
      Width           =   345
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
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
      TabIndex        =   19
      Top             =   885
      Width           =   975
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Cadastro"
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
      Left            =   4560
      TabIndex        =   18
      Top             =   885
      Width           =   1230
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
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
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   2145
      Width           =   600
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento"
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
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   2655
      Width           =   1005
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
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
      Left            =   120
      TabIndex        =   15
      Top             =   1635
      Width           =   450
   End
End
Attribute VB_Name = "alid015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private WithEvents conexaoAdo As ADODB.Connection
Private WithEvents RsReceita As ADODB.Recordset
Attribute RsReceita.VB_VarHelpID = -1
Private mblnAddMode As Boolean
Private LcSql       As String
Private LcAcao      As Integer
Private LcCarregado As Boolean

Private Function PodeExcluir() As Boolean
Dim RsGrupo As ADODB.Recordset
Dim Pode As Boolean
Dim LcCriterio As String
Dim afetados As Integer
'Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsGrupo = AbreRecordset("select * from Usuario where Nome='" & GlUsuario & "'")
If RsGrupo.EOF Then
   Pode = False
Else
   Pode = RsGrupo!ExcluiReceita
End If
RsGrupo.Close
Set RsGrupo = Nothing
PodeExcluir = Pode
End Function
Private Sub botoes_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case UCase(Trim(Button))
    Case Is = "&FECHAR F12"
        Unload Me
    Case Is = "&SALVAR F2"
        If LcAcao = 1 Then
           ' If IncluiRegistroReceita(Me, RsReceita) Then
            If AdicionaRegistro Then
                
                   GeraPainel Me, LcAcao, RsReceita
                   
                   MsgBox "Registros Salvos com Sucesso.", 64, "Aviso"
                   LimpaControles Me
                   Button.Enabled = False
                   DataCadastro.Text = Format(Date, "dd/mm/yy")
               ' Else
               '    MsgBox "Registro não foi Salvo", 64, "Aviso"
               ' End If
            Else
                   MsgBox "Erro Salvando Registros", 64, "Aviso"
            End If
        Else
        
            If AlteraRegistro Then
                 If IsNumeric(BotoesClienteAnterior.Text) Then
                   If CLng(CLIENTE.Text) <> CLng(BotoesClienteAnterior.Text) Then
                       If MsgBox("Altera as outras receitas com o codigo do cliente para o novo codigo?", vbYesNo, "Confirmação") = vbYes Then
                          Dim StrSql As String
                          Dim Afetado As Long
                          StrSql = "Update alid015 set Cliente='" & CLIENTE.Text & "' where cliente='" & BotoesClienteAnterior.Text & "'"
                          Afetado = ExecutaSql(StrSql)
                          MsgBox Afetado & " Registro(s) Alterados.", 64, "Aviso"
                          BotoesClienteAnterior.Text = ""
                       End If
                   End If
                 End If
                   GeraPainel Me, LcAcao, RsReceita
                   MsgBox "Registro Salvo com Sucesso.", 64, "Aviso"
                   RsReceita.Requery
            Else
                   MsgBox "Erro Salvando Registros", 64, "Aviso"
            End If
        
        End If
    Case Is = "&INCLUIR F3"
        LcAcao = 1
        DesabilitaNavegacao
        LimpaControles Me
        NF.Locked = False
        DataCadastro.Text = Format(Date, "dd/mm/yy")
        botoes.Buttons(4).Enabled = True
        botoes.Buttons(3).Enabled = True
        botoes.Buttons(5).Enabled = False
        botoes.Buttons(6).Enabled = False
        botoes.Buttons(7).Enabled = False

        GeraPainel Me, LcAcao, RsReceita
        NF.SetFocus
    Case Is = "&ALTERAR F4"
        If RsReceita.EOF And RsReceita.BOF Then
           MsgBox "Operação não Disponivel.", 64, "Sem Registro Atual"
           Exit Sub
        End If
        HabilitaNavegacao
        LcAcao = 2
        nome.SetFocus
        botoes.Buttons(2).Enabled = True
        botoes.Buttons(5).Enabled = True
        botoes.Buttons(6).Enabled = True
        botoes.Buttons(7).Enabled = True
        botoes.Buttons(3).Enabled = True
        botoes.Buttons(4).Enabled = False
        VincularTabela Me, RsReceita
        'VincularDados RsReceita, Me
        GeraPainel Me, LcAcao, RsReceita
        NF.Locked = True
    Case Is = "&CONSULTAR F4"
        If RsReceita.EOF And RsReceita.BOF Then
           MsgBox "Operação não Disponivel.", 64, "Sem Registro Atual"
           Exit Sub
        End If

        HabilitaNavegacao
        LcAcao = 3
        botoes.Buttons(6).Enabled = True
        botoes.Buttons(7).Enabled = True
        botoes.Buttons(2).Enabled = True
        botoes.Buttons(5).Enabled = False
        botoes.Buttons(4).Enabled = True
        botoes.Buttons(3).Enabled = False
        botoes.Buttons(1).Enabled = False
        VincularTabela Me, RsReceita
        GeraPainel Me, LcAcao, RsReceita
        nome.SetFocus
       
    Case Is = "&EXCLUIR F5"
        If Len(codigo.Text) = 0 Then Exit Sub
        If PodeExcluir() Then
              LcResposta = MsgBox("Confirma a Exclusão deste Registro ?", vbExclamation + vbYesNo, "Excluir Registro")
            If LcResposta = 7 Then Exit Sub
            If ExcluirRegistro(Me, CLng(codigo.Text), RsReceita) Then
               '==> Verifica a Quantidade de REgistro que foi excluido
               If LcRegistrosAfetados > 0 Then
                    
                    If RsReceita.EOF Then
                       If RsReceita.BOF Then
                          LcAcao = 1
                          LimpaControles Me
                          DesabilitaNavegacao
                       Else
                          RsReceita.MovePrevious
                          If RsReceita.BOF Then
                            LcAcao = 1
                            LimpaControles Me
                            DesabilitaNavegacao
                            botoes.Buttons(5).Enabled = False
                          Else
                            VincularTabela Me, RsReceita
                          End If
                       End If
                    Else
                        RsReceita.MoveNext
                        If RsReceita.EOF Then
                            LcAcao = 1
                            LimpaControles Me
                            DesabilitaNavegacao
                        Else
                            VincularTabela Me, RsReceita
                        End If
                    End If
                    GeraPainel Me, LcAcao, RsReceita
                    MsgBox "Registro Excluido com Sucesso.", 64, "Aviso"
               Else
                    MsgBox "O Registro Não Foi Encontrado para a Exclusão", 64, "Aviso"
               End If
            Else
               MsgBox "Erro Excluido Registro.", 64, "Erro Ocorrido"
            End If
            
        Else
            MsgBox "Usuário sem Permissão para Excluir", vbCritical, "Aviso"
        End If
        
        
    
    Case Is = "&ORDENAR F7"
        Load ordenar
        Call ordenar.ordenar(Me, RsReceita)
        ordenar.Show , Me
    Case Is = "PES&QUISAR F6"
        Load pesquisa
        Call pesquisa.CriaLista(Me, RsReceita)
        pesquisa.Show , Me

End Select


End Sub
Function GeraBoletoA4()
Dim ClBoleto As ControlaBoleto
Dim ClAcesso As New AcessoAdo.Acessos
Dim RsCedente As ADODB.Recordset
Dim RsClientes As Recordset
Dim RsCidade As Recordset
Dim db As Database
Dim StrSql As String
Dim a As Integer
Dim LcMargemBo As String
Dim Protesto    As String
Dim protesto1   As String
Dim LocalAr As String
Set ClBoleto = New ControlaBoleto
On Error GoTo erroImpressao

Set db = OpenDatabase(GLBase)
Set RsClientes = db.OpenRecordset("Select * from alid001 where codigo='" & CLIENTE.Text & "'")
LocalAr = App.EXEName & ".ini"
ClBoleto.NomeProjeto = LocalAr
StrSql = "Select * from contasacado"

Set RsCedente = AbreRecordset(StrSql, True)
Set RsCidade = db.OpenRecordset("select * from alid005 where Cod='" & RsClientes!Cidade & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)

If RsCedente.EOF Then
   MsgBox "A conta do cedente não foi cadastrada.", 64, "Aviso"
   Exit Function
End If

LcJuros = "             JUROS DE 5 % AO MES"
LcPag = "             ATE A DATA DO VENCIMENTO PAGAR EM QUALQUER BANCO / QUALQUER AGENCIA"
Protesto = "             NAO RECEBER APOS 4 (QUATRO) DIAS UTEIS DO VENCIMENTO."
protesto1 = "             SUJEITO A PROTESTO"
ClBoleto.BoletodoBanco = "Itau"
ClBoleto.Agencia = RsCedente!Agencia
ClBoleto.Cedente = RsCedente!NomeSacado
ClBoleto.Conta = RsCedente!Conta
ClBoleto.ContaDigito = RsCedente!DigitoConta
ClBoleto.Carteira = RsCedente!Carteira
ClBoleto.BairroSacado = RsClientes!Bairro & ""
ClBoleto.CepSacado = RsClientes!Cep & ""
ClBoleto.NossoNumero = NossoNumero.Text

If Not RsCidade.EOF Then
   ClBoleto.CidadeSacado = RsCidade!nome & ""
Else
  ClBoleto.CidadeSacado = ""
End If
ClBoleto.CnpjSacado = "CNPJ/CPF: " & RsClientes!CGC & ""
ClBoleto.DataDocumento = Format(DATA.Text, "dd/mm/yy")
ClBoleto.EnderecoSacado = RsClientes!End & ""
ClBoleto.EspecieDoc = "DP"
ClBoleto.EstadoSacado = RsClientes!Estado & ""
ClBoleto.IncricaoSacado = RsClientes!INSCEST & ""
ClBoleto.Instrucao1 = LcJuros
ClBoleto.Instrucao2 = Protesto
ClBoleto.Instrucao3 = protesto1
ClBoleto.Instrucao4 = ""
ClBoleto.NomeSacado = RsClientes!RAZAOSOC & ""
ClBoleto.RgSacado = ""
ClBoleto.Especie = "R$"
ClBoleto.Aceite = "N"
ClBoleto.NumeroDocumento = NF.Text
ClBoleto.Vencimento = DTVENC
ClBoleto.VALOR = Format(VALOR.Text, "Standard")
ClBoleto.GeraBoleto imprimir
    

Set ClBoleto = Nothing


Exit Function
erroImpressao:
MsgBox err.Description & " Nº:" & err.Number

'Resume 0

End Function


Private Sub botoes1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case UCase(Trim(Button))
    Case Is = UCase("Imprimir Boleto")
       GeraBoletoA4
    Case Is = "&FECHAR F12"
        Unload Me
    Case Is = "&PRIMEIRO F8"
        RsReceita.MoveFirst
        'RsReceita.Requery
        VincularTabela Me, RsReceita
    Case Is = "&ULTIMO F11"
        RsReceita.MoveLast
        'RsReceita.Requery
        VincularTabela Me, RsReceita
    Case Is = "A&NTERIOR F9"
        'RsReceita.Requery
        RsReceita.MovePrevious
        If Not RsReceita.BOF Then
           VincularTabela Me, RsReceita
        Else
           If Not RsReceita.EOF Then RsReceita.MoveNext
           MsgBox "Este é o Primeiro Registro.", 64, "Aviso"
        End If
    Case Is = "SE&GUINTE F10"
        'RsReceita.Requery
        RsReceita.MoveNext
        If Not RsReceita.EOF Then
           VincularTabela Me, RsReceita
        Else
           If Not RsReceita.BOF Then RsReceita.MovePrevious
           MsgBox "Este é o Ultimo Registro.", 64, "Aviso"
        End If

End Select


End Sub

Private Sub botoes3_Click()
GlCriterioSql = ""
    Load FrmPesquisaCliente
    FrmPesquisaCliente.Txt.Text = "" 'NomeCliente.Text
    FrmPesquisaCliente.ExibePesquisa
    FrmPesquisaCliente.Show , Me
End Sub

Private Sub botoes4_Click()
ExibeMonetario.Show , Me
End Sub

Private Sub CLIENTE_Change()
On Error Resume Next
Dim Rs As Recordset
Dim Lcs As String
Lcs = "Select * from  alid001  where codigo='" & CLIENTE.Text & "'"
AbreBase
Set Rs = Dbbase.OpenRecordset(Lcs, dbOpenDynaset, dbSeeChanges, dbOptimistic)
If Not Rs.EOF Then
   nome.Text = Rs!RAZAOSOC & ""
Else
   nome.Text = ""
End If
Rs.Close
Dbbase.Close
Set Rs = Nothing
Set Dbbase = Nothing

End Sub

Private Sub CmdPesquisaCliente_Click()

End Sub

Private Sub CmdBuscaTipo_Click()
ExibeMonetario.Show (Me)
End Sub

Private Sub Form_Activate()
On Error Resume Next
Set GlFormA = Me
Titulo.Caption = " Cadastro de Receitas"
If Not LcCarregado Then

  If RsReceita.EOF And RsReceita.BOF Then
     DesabilitaNavegacao
     Call LimpaControles(Me)
     botoes.Buttons(2).Enabled = False
     botoes.Buttons(6).Enabled = False
     botoes.Buttons(7).Enabled = False
     botoes.Buttons(5).Enabled = True
     botoes.Buttons(3).Enabled = True
     botoes.Buttons(4).Enabled = False
     DataCadastro.Text = Format(Date, "dd/mm/yy")
     LcAcao = 1
     
  Else
     RsReceita.MoveFirst
     VincularTabela Me, RsReceita
  End If
  GeraPainel Me, LcAcao, RsReceita

End If
LcCarregado = True
DoEvents
Me.Refresh
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then Call botoes_ButtonClick(botoes.Buttons(1))
If KeyCode = 114 Then Call botoes_ButtonClick(botoes.Buttons(2))

If KeyCode = 115 Then
   If botoes.Buttons(3).Enabled Then
      Call botoes_ButtonClick(botoes.Buttons(3))
   Else
      Call botoes_ButtonClick(botoes.Buttons(4))
   End If
End If
If KeyCode = 116 Then Call botoes_ButtonClick(botoes.Buttons(5))
If KeyCode = 117 Then Call botoes_ButtonClick(botoes.Buttons(6))
If KeyCode = 118 Then Call botoes_ButtonClick(botoes.Buttons(7))

If KeyCode = 119 Then Call botoes1_ButtonClick(botoes1.Buttons(1))
If KeyCode = 120 Then Call botoes1_ButtonClick(botoes1.Buttons(2))
If KeyCode = 121 Then Call botoes1_ButtonClick(botoes1.Buttons(3))
If KeyCode = 122 Then Call botoes1_ButtonClick(botoes1.Buttons(4))
If KeyCode = 123 Then Call botoes1_ButtonClick(botoes1.Buttons(5))

If KeyCode >= 32 And KeyCode <= 127 And LcAcao <> 3 Then
    botoes.Buttons(1).Enabled = True
Else
   KeyCode = 0
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
If LcAcao = 3 And KeyAscii <> 13 Then KeyAscii = 0

End Sub

Private Sub Form_Load()
On Error Resume Next
'Me.Top = 1740
'Me.Left = FrmPrincipal.Width '3090
Set LcF = Me
LcAcao = 3
botoes.Buttons(5).Enabled = False
'abreconexao
'Call 'abreconexao
Call ChamaRecord("nome")

End Sub
Function ChamaRecord(LcOrdem As String)
On Error Resume Next
Dim LcMat As Variant
Dim a As Long
LcSql = "Select * from " & Me.Name
Set RsReceita = AbreRecordset(LcSql)
RsReceita.Sort = "DOCUMENTO"
End Function
Function ordena(LcOrdem As String) As Boolean
On Error Resume Next
err.Number = 0
RsReceita.Sort = LcOrdem
'MsgBox Err.Number
If err.Number <> 0 Then
   MsgBox "Não é possivel Ordenar pela Seleção Escolhida," & Chr(13) & "Verifique se todos os campos escolhidos possuiem dados.", vbCritical + vbExclamation, "Ordem não Aceita"
   ordena = False
Else
   ordena = True
End If
GeraPainel Me, LcAcao, RsReceita
End Function

Private Sub DesabilitaNavegacao()
On Error Resume Next
Dim a As Long
For a = 1 To 4
   botoes1.Buttons(a).Enabled = False
Next
botoes.Buttons(2).Enabled = False
End Sub
Private Sub HabilitaNavegacao()
On Error Resume Next
Dim a As Long
For a = 1 To 4
   botoes1.Buttons(a).Enabled = True
Next
botoes.Buttons(2).Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'Call 'FechaConexao
'FechaConexao
RsReceita.Close
Set RsReceita = Nothing
LcCarregado = False
FrmPrincipal.SetFocus
End Sub

Private Sub TIPORD_Change()


End Sub

Private Sub nome_Change()
If LcAcao <> 3 Then botoes.Buttons(1).Enabled = True
End Sub

Private Sub nomereceita_Change()
If LcAcao <> 3 Then botoes.Buttons(1).Enabled = True
End Sub

Private Sub TPMONET_Change()
On Error Resume Next
Dim Rs As Recordset
Dim Lcs As String
Lcs = "Select * from  alid008  where TPMONET='" & TPMONET.Text & "'"
AbreBase
Set Rs = Dbbase.OpenRecordset(Lcs, dbOpenDynaset, dbSeeChanges, dbOptimistic)
If Not Rs.EOF Then
   nomereceita.Text = Rs!XTPMONET & ""
Else
   nomereceita.Text = ""
End If
Rs.Close
Dbbase.Close
Set Rs = Nothing
Set Dbbase = Nothing

End Sub
Function AlteraRegistroReceita(lcform As Form, LcCodigo As Long, Rs As ADODB.Recordset) As Boolean
On Error GoTo ErrAlteracao
Dim C           As Control
Dim LcNome      As String
Dim LcType      As Integer
Dim LcSql       As String
Dim LcCampos    As String
Dim LcValores   As String
Dim LcNomeL     As String
Dim LcPrimeiro  As Boolean
Dim LcIncluiCampo As Boolean

On Error Resume Next
err.Number = 0
LcNome = Rs!nome & ""
If err.Number <> 0 Then
   err.Number = 0
   LcNome = Rs!Descricao & ""
   If err.Number <> 0 Then
      err.Number = 0
      LcNome = Rs!RAZAOSOC & ""
      If err.Number <> 0 Then
         err.Number = 0
         LcNome = Rs!Endereco & ""
      End If
         If err.Number <> 0 Then
            err.Number = 0
            LcNome = Rs!NF & ""
            If err.Number <> 0 Then
               err.Number = 0
               LcNome = Rs!XTPMONET & ""
               If err.Number <> 0 Then
                  err.Number = 0
                  LcNome = Rs!NumNf & ""
                  If err.Number <> 0 Then
                     err.Number = 0
                     LcNome = Rs!cheque & ""
                     If err.Number <> 0 Then
                        err.Number = 0
                        LcNome = Rs!Doc & ""
                     End If
                  End If
               End If
            End If
         End If
   End If
End If
LcNomeL = LcNome
On Error GoTo ErrAlteracao
LcPrimeiro = True
LcComentario = "-AlteraRegistro- Criando Sql."
LcSql = ""
LcSql = "Update " & lcform.Name & " SET "
LcComentario = "-AlteraRegistro- Efetuando o Loop No Form para buscar os campos e Valores."
For Each C In lcform.Controls()
    LcNome = UCase(C.Name)
    LcIncluiCampo = False
    If LcNome <> "NOME" And LcNome <> "NOMERECEITA" And LcNome <> "TITULO" And LcNome <> "CODIGO" And LcNome <> "BOTOES1" And LcNome <> "BARSTATUS" And LcNome <> "LINE" And LcNome <> "LABEL" And LcNome <> "TAB" And LcNome <> "BOTOES" And LcNome <> "FIGURAS" Then
        LcComentario = "-AlteraRegistro- Setando o Tipo do Campo."
        LcType = Rs.Fields(LcNome).Type
        Select Case LcType
           Case 135
                 LcComentario = "-AlteraRegistro- Setando o Tipo Data."
                 If IsDate(C.Text) Then
                    Rs(LcNome).Value = Format(C.Text, "dd/mm/YY")
                 Else
                    Rs(LcNome).Value = Null
                 End If

            Case adDate
                 LcComentario = "-AlteraRegistro- Setando o Tipo Data."
                 If IsDate(C.Text) Then
                    Rs(LcNome).Value = Format(C.Text, "dd/mm/YY")
                 Else
                    Rs(LcNome).Value = Null
                 End If
            
            Case Is = adDBDate
                 LcComentario = "-AlteraRegistro- Setando o Tipo Data."
                 If IsDate(C.Text) Then
                    Rs(LcNome).Value = Format(C.Text, "dd/mm/yy")
                 Else
                    Rs(LcNome).Value = Null
                 End If
            Case Is = dbBoolean
                  LcComentario = "-AlteraRegistro- Setando o Tipo Boleano."
                  Rs(LcNome).Value = C.vaslue
            Case Is = adDouble
                 LcComentario = "-AlteraRegistro- Setando o Tipo Inteiro."
                 If IsNumeric(C.Text) Then
                    Rs(LcNome).Value = CDbl(C.Text)
                 End If
            
            Case Is = adDecimal
                 LcComentario = "-AlteraRegistro- Setando o Tipo Inteiro."
                 If IsNumeric(C.Text) Then
                    Rs(LcNome).Value = CDbl(C.Text)
                 End If
            
            Case Is = adInteger
                 LcComentario = "-AlteraRegistro- Setando o Tipo Inteiro."
                 If IsNumeric(C.Text) Then
                    Rs(LcNome).Value = CInt(C.Text)
                 
                 End If
            
            Case Is = adCurrency
                 LcComentario = "-AlteraRegistro- Setando o Tipo Inteiro."
                 If IsNumeric(C.Text) Then
                    Rs(LcNome).Value = CCur(C.Text)
                 End If
            
            Case adInteger
                 LcComentario = "-AlteraRegistro- Setando o Tipo Inteiro."
                 If IsNumeric(C.Text) Then
                    Rs(LcNome).Value = CInt(C.Text)
                 End If
            
            Case Is = adNumeric
                 LcComentario = "-AlteraRegistro- Setando o Tipo Numérico."
                 If IsNumeric(C.Text) Then
                    Rs(LcNome).Value = CDbl(C.Text)
                 End If
            Case Is = adLongVarChar
                   LcComentario = "-AlteraRegistro- Setando o Tipo String."
                   Rs(LcNome).Value = CLng(UCase(C.Text))
            Case Is = adChar
                   LcComentario = "-AlteraRegistro- Setando o Tipo String."
                   Rs(LcNome).Value = UCase(C.Text)
                    
            Case Is = adVarChar
                   Rs(LcNome).Value = UCase(C.Text)
        End Select
       
    End If
Next
LcComentario = "-AlteraRegistro- Efetuando a alteração."
Rs.Update
LcComentario = "-AlteraRegistro- Gravando o Log."
Call GravaLogSistema(lcform.Name, "ALTERAÇÂO", CLng(LcCodigo), LcNomeL)

LcComentario = "-AlteraRegistro- Atualizando o recordset."


'Rs.Requery

AlteraRegistroReceita = True
Saida:
Exit Function
ErrAlteracao:
logErro err.Number, err.Description, LcComentario
MsgBox err.Description & err.Number
AlteraRegistroReceita = False
Resume 0
GoTo Saida

End Function

Function IncluiRegistroReceita(lcform As Form, Rs As ADODB.Recordset) As Boolean
On Error GoTo ErrAlteracao
Dim C           As Control
Dim LcNome      As String
Dim LcType      As Integer
Dim LcSql       As String
Dim LcCampos    As String
Dim LcValores   As String
Dim LcNomeL     As String
Dim LcPrimeiro  As Boolean
Dim LcIncluiCampo As Boolean

On Error Resume Next
err.Number = 0
LcNome = Rs!nome & ""
If err.Number <> 0 Then
   err.Number = 0
   LcNome = Rs!Descricao & ""
   If err.Number <> 0 Then
      err.Number = 0
      LcNome = Rs!RAZAOSOC & ""
      If err.Number <> 0 Then
         err.Number = 0
         LcNome = Rs!Endereco & ""
      End If
         If err.Number <> 0 Then
            err.Number = 0
            LcNome = Rs!NF & ""
            If err.Number <> 0 Then
               err.Number = 0
               LcNome = Rs!XTPMONET & ""
               If err.Number <> 0 Then
                  err.Number = 0
                  LcNome = Rs!NumNf & ""
                  If err.Number <> 0 Then
                     err.Number = 0
                     LcNome = Rs!cheque & ""
                     If err.Number <> 0 Then
                        err.Number = 0
                        LcNome = Rs!Doc & ""
                     End If
                  End If
               End If
            End If
         End If
   End If
End If
LcNomeL = LcNome
On Error GoTo ErrAlteracao
LcPrimeiro = True
LcComentario = "-AlteraRegistro- Criando Sql."
LcSql = ""
LcSql = "Update " & lcform.Name & " SET "
LcComentario = "-AlteraRegistro- Efetuando o Loop No Form para buscar os campos e Valores."
Rs.AddNew
For Each C In lcform.Controls()
    LcNome = UCase(C.Name)
    LcIncluiCampo = False
    If LcNome <> "NOME" And LcNome <> "NOMERECEITA" And LcNome <> "TITULO" And LcNome <> "CODIGO" And LcNome <> "BOTOES1" And LcNome <> "BARSTATUS" And LcNome <> "LINE" And LcNome <> "LABEL" And LcNome <> "TAB" And LcNome <> "BOTOES" And LcNome <> "FIGURAS" Then
        LcComentario = "-AlteraRegistro- Setando o Tipo do Campo."
        LcType = Rs.Fields(LcNome).Type
        Select Case LcType
           Case 135
                 LcComentario = "-AlteraRegistro- Setando o Tipo Data."
                 If IsDate(C.Text) Then
                    Rs(LcNome).Value = Format(C.Text, "dd/mm/YY")
                 Else
                    Rs(LcNome).Value = Null
                 End If

            Case adDate
                 LcComentario = "-AlteraRegistro- Setando o Tipo Data."
                 If IsDate(C.Text) Then
                    Rs(LcNome).Value = Format(C.Text, "dd/mm/YY")
                 Else
                    Rs(LcNome).Value = Null
                 End If
            
            Case Is = adDBDate
                 LcComentario = "-AlteraRegistro- Setando o Tipo Data."
                 If IsDate(C.Text) Then
                    Rs(LcNome).Value = Format(C.Text, "dd/mm/yy")
                 Else
                    Rs(LcNome).Value = Null
                 End If
            Case Is = dbBoolean
                  LcComentario = "-AlteraRegistro- Setando o Tipo Boleano."
                  Rs(LcNome).Value = C.vaslue
            Case Is = adDouble
                 LcComentario = "-AlteraRegistro- Setando o Tipo Inteiro."
                 If IsNumeric(C.Text) Then
                    Rs(LcNome).Value = CDbl(C.Text)
                 End If
            
            Case Is = adDecimal
                 LcComentario = "-AlteraRegistro- Setando o Tipo Inteiro."
                 If IsNumeric(C.Text) Then
                    Rs(LcNome).Value = CDbl(C.Text)
                 End If
            
            Case Is = adInteger
                 LcComentario = "-AlteraRegistro- Setando o Tipo Inteiro."
                 If IsNumeric(C.Text) Then
                    Rs(LcNome).Value = CInt(C.Text)
                 
                 End If
            
            Case Is = adCurrency
                 LcComentario = "-AlteraRegistro- Setando o Tipo Inteiro."
                 If IsNumeric(C.Text) Then
                    Rs(LcNome).Value = CCur(C.Text)
                 End If
            
            Case adInteger
                 LcComentario = "-AlteraRegistro- Setando o Tipo Inteiro."
                 If IsNumeric(C.Text) Then
                    Rs(LcNome).Value = CInt(C.Text)
                 End If
            
            Case Is = adNumeric
                 LcComentario = "-AlteraRegistro- Setando o Tipo Numérico."
                 If IsNumeric(C.Text) Then
                    Rs(LcNome).Value = CDbl(C.Text)
                 End If
            Case Is = adLongVarChar
                   LcComentario = "-AlteraRegistro- Setando o Tipo String."
                   Rs(LcNome).Value = CLng(UCase(C.Text))
            Case Is = adChar
                   LcComentario = "-AlteraRegistro- Setando o Tipo String."
                   Rs(LcNome).Value = UCase(C.Text)
                    
            Case Is = adVarChar
                   Rs(LcNome).Value = UCase(C.Text)
        End Select
       
    End If
Next
LcComentario = "-AlteraRegistro- Efetuando a alteração."
Rs.Update
LcComentario = "-AlteraRegistro- Gravando o Log."
Call GravaLogSistema(lcform.Name, "ALTERAÇÂO", CLng(LcCodigo), LcNomeL)

LcComentario = "-AlteraRegistro- Atualizando o recordset."


'Rs.Requery

IncluiRegistroReceita = True
Saida:
Exit Function
ErrAlteracao:
logErro err.Number, err.Description, LcComentario
MsgBox err.Description & err.Number
IncluiRegistroReceita = False
Resume 0
GoTo Saida

End Function
Function AdicionaRegistro() As Boolean
On Error GoTo erroadi
'UPDATE ALID015 SET ALID015.DTPAGTO = Null
'==> Esta funcao adiciona um novo registro ao banco de dados
Dim StrSql      As String
Dim Afetado     As Integer
Dim a           As Integer
Dim Campos      As String
Dim Valores     As String
Dim NomeCampo   As String
Dim C As Control
Campos = ""
Valores = ""
'==> Processa os campos e valores
For Each C In Me.Controls()
    NomeCampo = UCase(C.Name)
    If Len(Campos) > 0 Then
       If Right(Campos, 1) <> "," Then Campos = Campos & ","
    End If
    If Len(Valores) > 0 Then
       If Right(Valores, 1) <> "," Then Valores = Valores & ","
    End If
    
    Select Case NomeCampo
        Case Is = "NF"
            Campos = Campos & NomeCampo
            Valores = Valores & "'" & NF.Text & "'"
        Case Is = "DATA"
            Campos = Campos & NomeCampo
            Valores = Valores & "'"
            If IsDate(DATA.Text) Then
                Valores = Valores & Format(DATA.Text, "yyyy-mm-dd") & "'"
            Else
               Valores = Valores & Null & "'"
            End If
        Case Is = "VALOR"
            Campos = Campos & NomeCampo
            If IsNumeric(VALOR.Text) Then
               Valores = Valores & Replace(VALOR.Text, ",", ".")
            Else
               Valores = Valores & "0"
            End If
        Case Is = "CLIENTE"
            Campos = Campos & NomeCampo
            Valores = Valores & "'" & CLIENTE.Text & "'"
        Case Is = "DTVENC"
            Campos = Campos & NomeCampo
            Valores = Valores & "'"
            If IsDate(DTVENC.Text) Then
                Valores = Valores & Format(DTVENC.Text, "yyyy-mm-dd") & "'"
            Else
               Valores = Valores & Null & "'"
            End If
        Case Is = "TPMONET"
            Campos = Campos & NomeCampo
            Valores = Valores & "'" & TPMONET.Text & "'"
        Case Is = "DTPAGTO"
            Campos = Campos & NomeCampo
            Valores = Valores & "'"
            If IsDate(DTPAGTO.Text) Then
                Valores = Valores & Format(DTPAGTO.Text, "yyyy-mm-dd") & "'"
            Else
               Valores = Valores & Null & "'"
            End If

        Case Is = "VALPAGO"
            Campos = Campos & NomeCampo
            If IsNumeric(VALPAGO.Text) Then
               Valores = Valores & Replace(VALPAGO.Text, ",", ".")
            Else
               Valores = Valores & "0"
            End If

        Case Is = "Obs"
            Campos = Campos & NomeCampo
            Valores = Valores & "'" & Obs.Text & "'"
    End Select

Next
If Right(Campos, 1) = "," Then
   Campos = Left(Campos, Len(Campos) - 1)
End If
If Right(Valores, 1) = "," Then
   Valores = Left(Valores, Len(Valores) - 1)
End If

StrSql = "insert into Alid015(" & Campos & ") Values(" & Valores & ")"
'MsgBox strSql
'abreconexao
Afetado = ExecutaSql(StrSql)

If Afetado = 1 Then AdicionaRegistro = True Else AdicionaRegistro = False
Exit Function

erroadi:
MsgBox err.Description & err.Number
    'Resume 0




End Function

Function AlteraRegistro() As Boolean
On Error GoTo erroadi
'UPDATE ALID015 SET ALID015.DTPAGTO = Null
'==> Esta funcao Altera o registro ao banco de dados
Dim StrSql      As String
Dim Afetado     As Integer
Dim a           As Integer
Dim Campos      As String
Dim NomeCampo   As String
Dim C As Control
Campos = ""
Valores = ""
'==> Processa os campos e valores
For Each C In Me.Controls()
    NomeCampo = UCase(C.Name)
    If Len(Campos) > 0 Then
       If Right(Campos, 1) <> "," Then Campos = Campos & ","
    End If
    
    Select Case NomeCampo
        Case Is = "NF"
            Campos = Campos & NomeCampo & "='" & NF.Text & "'"
        Case Is = "DATA"
            If IsDate(DATA.Text) Then
                Campos = Campos & NomeCampo & "='" & Format(DATA.Text, "yyyy-mm-dd") & "'"
            Else
               Campos = Campos & NomeCampo & "=Null"
            End If
        Case Is = "VALOR"
            If IsNumeric(VALOR.Text) Then
               Campos = Campos & NomeCampo & "=" & Replace(VALOR.Text, ",", ".")
            Else
               Campos = Campos & NomeCampo & "=0"
            End If
        Case Is = "CLIENTE"
            Campos = Campos & NomeCampo & "='" & CLIENTE.Text & "'"
        Case Is = "DTVENC"
            If IsDate(DTVENC.Text) Then
                Campos = Campos & NomeCampo & "='" & Format(DTVENC.Text, "yyyy-mm-dd") & "'"
            Else
               Campos = Campos & NomeCampo & "=Null"
            End If
        Case Is = "TPMONET"
            Campos = Campos & NomeCampo & "='" & TPMONET.Text & "'"
        Case Is = "DTPAGTO"
            If IsDate(DTPAGTO.Text) Then
                Campos = Campos & NomeCampo & "='" & Format(DTPAGTO.Text, "yyyy-mm-dd") & "'"
            Else
               Campos = Campos & NomeCampo & "=Null"
            End If

        Case Is = "VALPAGO"
            If IsNumeric(VALPAGO.Text) Then
               Campos = Campos & NomeCampo & "=" & Replace(VALPAGO.Text, ",", ".")
            Else
               Campos = Campos & NomeCampo & "=0"
            End If
        Case Is = "ACRESCIMO"
            If IsNumeric(Acrescimo.Text) Then
               Campos = Campos & NomeCampo & "=" & Replace(Acrescimo.Text, ",", ".")
            Else
               Campos = Campos & NomeCampo & "=0"
            End If
        Case Is = UCase("Obs")
            Campos = Campos & NomeCampo & "='" & Obs.Text & "'"
        Case Is = UCase("NossoNumero")
             Campos = Campos & NomeCampo & "='" & NossoNumero.Text & "'"
        
    End Select

Next
If Right(Campos, 1) = "," Then
   Campos = Left(Campos, Len(Campos) - 1)
End If


StrSql = "UPDATE ALID015 SET " & Campos & " where codigo=" & codigo.Text
'MsgBox strSql
Debug.Print StrSql
'abreconexao
Afetado = ExecutaSql(StrSql)
Debug.Print StrSql

'MsgBox DEscricaoErro
If Afetado = 1 Or Len(DEscricaoErro) = 0 Then AlteraRegistro = True Else AlteraRegistro = False
Exit Function

erroadi:
MsgBox err.Description & err.Number
'Resume 0

End Function

