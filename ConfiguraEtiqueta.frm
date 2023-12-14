VERSION 5.00
Begin VB.Form ConfiguraEtiqueta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuração das Etiquetas de Mala Direta"
   ClientHeight    =   6465
   ClientLeft      =   2490
   ClientTop       =   1260
   ClientWidth     =   7425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7425
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   375
      Left            =   2520
      TabIndex        =   23
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salvar F2"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "Espaçamentos"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   17
      Top             =   4440
      Width           =   7095
      Begin VB.TextBox Horizontal 
         Height          =   285
         Left            =   2520
         TabIndex        =   20
         Text            =   "2"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox vertical 
         Height          =   285
         Left            =   2520
         TabIndex        =   18
         Text            =   "4"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Distância Horizontal (em Linhas)"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label8 
         Caption         =   "Distância Veritcal (em Colunas)"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dados da Etiqueta"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   7095
      Begin VB.TextBox altura 
         Height          =   375
         Left            =   4680
         TabIndex        =   15
         Text            =   "4"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Largura 
         Height          =   405
         Left            =   1800
         TabIndex        =   13
         Text            =   "40"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Altura (em Linhas)"
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Largura (em Colunas)"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.TextBox Etiquetas 
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   10
      Text            =   "4"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Número de Etiquetas"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   7095
      Begin VB.TextBox Etiquetas 
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Text            =   "2"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Por Colunas"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Por Linha"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.TextBox Margem 
      Height          =   285
      Index           =   2
      Left            =   2160
      TabIndex        =   5
      Text            =   "0"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Margem 
      Height          =   285
      Index           =   1
      Left            =   5520
      TabIndex        =   3
      Text            =   "0"
      Top             =   720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Margens "
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
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7095
      Begin VB.TextBox Margem 
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   1
         Text            =   "0"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Esquerda (em colunas)"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Inferior (em Linhas)"
         Height          =   255
         Left            =   3600
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Superior (em Linhas)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "ConfiguraEtiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private a As Integer
Private Sub altura_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim RsEtiqueta As Recordset
AbreBase
Set RsEtiqueta = Dbbase.OpenRecordset("Etiqueta", dbOpenDynaset, dbSeeChanges, dbOptimistic)
If RsEtiqueta.EOF Then
   RsEtiqueta.AddNew
Else
   RsEtiqueta.Edit
End If
RsEtiqueta("LarguraVertical") = Val(altura.Text)
RsEtiqueta("LarguraHorizontal") = Val(Largura.Text)
RsEtiqueta("MargemEsquerda") = Val(Margem(2).Text)
RsEtiqueta("MargemSuperior") = Val(Margem(0).Text)
RsEtiqueta("MargemInferior") = Val(Margem(1).Text)
RsEtiqueta("DistanciaHorizontal") = Val(Horizontal.Text)
RsEtiqueta("DistanciaVertical") = Val(vertical.Text)
RsEtiqueta("EtiquetasLinha") = Val(Etiquetas(0).Text)
RsEtiqueta("EtiquetaColuna") = Val(Etiquetas(1).Text)
RsEtiqueta.Update
RsEtiqueta.Close
Set RsEtiqueta = Nothing
Unload Me
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Etiquetas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Load()
Dim RsEtiqueta As Recordset
AbreBase
Set RsEtiqueta = Dbbase.OpenRecordset("Etiqueta", dbOpenDynaset, dbSeeChanges, dbOptimistic)
If Not RsEtiqueta.EOF Then
   altura.Text = RsEtiqueta("LarguraVertical")
   Largura.Text = RsEtiqueta("LarguraHorizontal")
   Margem(2).Text = RsEtiqueta("MargemEsquerda")
   Margem(0).Text = RsEtiqueta("MargemSuperior")
   Margem(1).Text = RsEtiqueta("MargemInferior")
   Horizontal.Text = RsEtiqueta("DistanciaHorizontal")
   vertical.Text = RsEtiqueta("DistanciaVertical")
   Etiquetas(0).Text = RsEtiqueta("EtiquetasLinha")
   Etiquetas(1).Text = RsEtiqueta("EtiquetaColuna")
End If
RsEtiqueta.Close
Set RsEtiqueta = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
End Sub

Private Sub Horizontal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Largura_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Margem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub vertical_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub
