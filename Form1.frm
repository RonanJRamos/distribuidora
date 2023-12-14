VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
' Read the computer's name
Dim compname As String * 255, cname As String
X = GetComputerName(compname, 255)
' Trim blank spaces and ending vbNullChar
cname = RTrim(LTrim(compname))
cname = Left(cname, Len(cname) - 1)
Text1.Text = cname
End Sub
