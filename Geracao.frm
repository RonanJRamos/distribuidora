VERSION 5.00
Begin VB.Form Geracao 
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdGerar 
      Caption         =   "Gerar"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.ListBox Meses 
      Height          =   2400
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5415
   End
   Begin VB.ComboBox Ano 
      Height          =   315
      ItemData        =   "Geracao.frx":0000
      Left            =   120
      List            =   "Geracao.frx":000D
      TabIndex        =   0
      Text            =   "04"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Ano"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Geracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdGerar_Click()
Dim Data1 As String
Dim data2 As String

Data1 = "01/01/" & Ano.Text
data2 = "31/01/" & Ano.Text
Sintegra.Datai.Text = Data1
Sintegra.Dataf.Text = data2
Sintegra.gerar
Meses.AddItem "Janeiro /" & Ano.Text & " - Ok"

Data1 = "01/02/" & Ano.Text
If Ano.Text = "04" Then data2 = "29/02/" & Ano.Text Else data2 = "28/02/" & Ano.Text
Sintegra.Datai.Text = Data1
Sintegra.Dataf.Text = data2
Sintegra.gerar
Meses.AddItem "Fevereiro /" & Ano.Text & " - Ok"

Data1 = "01/03/" & Ano.Text
data2 = "31/03/" & Ano.Text
Sintegra.Datai.Text = Data1
Sintegra.Dataf.Text = data2
Sintegra.gerar
Meses.AddItem "Março /" & Ano.Text & " - Ok"

Data1 = "01/04/" & Ano.Text
data2 = "30/04/" & Ano.Text
Sintegra.Datai.Text = Data1
Sintegra.Dataf.Text = data2
Sintegra.gerar
Meses.AddItem "Abril /" & Ano.Text & " - Ok"

Data1 = "01/05/" & Ano.Text
data2 = "31/05/" & Ano.Text
Sintegra.Datai.Text = Data1
Sintegra.Dataf.Text = data2
Sintegra.gerar
Meses.AddItem "Maio /" & Ano.Text & " - Ok"

Data1 = "01/06/" & Ano.Text
data2 = "30/06/" & Ano.Text
Sintegra.Datai.Text = Data1
Sintegra.Dataf.Text = data2
Sintegra.gerar
Meses.AddItem "Junho /" & Ano.Text & " - Ok"

Data1 = "01/07/" & Ano.Text
data2 = "31/07/" & Ano.Text
Sintegra.Datai.Text = Data1
Sintegra.Dataf.Text = data2
Sintegra.gerar
Meses.AddItem "Julho /" & Ano.Text & " - Ok"

Data1 = "01/08/" & Ano.Text
data2 = "31/08/" & Ano.Text
Sintegra.Datai.Text = Data1
Sintegra.Dataf.Text = data2
Sintegra.gerar
Meses.AddItem "Agosto /" & Ano.Text & " - Ok"

Data1 = "01/09/" & Ano.Text
data2 = "30/09/" & Ano.Text
Sintegra.Datai.Text = Data1
Sintegra.Dataf.Text = data2
Sintegra.gerar
Meses.AddItem "Setembro /" & Ano.Text & " - Ok"

Data1 = "01/10/" & Ano.Text
data2 = "31/10/" & Ano.Text
Sintegra.Datai.Text = Data1
Sintegra.Dataf.Text = data2
Sintegra.gerar
Meses.AddItem "Outubro /" & Ano.Text & " - Ok"

Data1 = "01/11/" & Ano.Text
data2 = "30/11/" & Ano.Text
Sintegra.Datai.Text = Data1
Sintegra.Dataf.Text = data2
Sintegra.gerar
Meses.AddItem "Novembro /" & Ano.Text & " - Ok"

Data1 = "01/12/" & Ano.Text
data2 = "31/12/" & Ano.Text
Sintegra.Datai.Text = Data1
Sintegra.Dataf.Text = data2
Sintegra.gerar
Meses.AddItem "Dezembro /" & Ano.Text & " - Ok"

End Sub
