VERSION 5.00
Begin VB.Form Sobre 
   Caption         =   "Sobre Censura"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   Icon            =   "Sobre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4665
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   1080
      Picture         =   "Sobre.frx":6852
      ScaleHeight     =   1035
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label LabelDesc 
      Alignment       =   2  'Center
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4335
   End
End
Attribute VB_Name = "Sobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    LabelDesc.Caption = "2011 - Censura Fácil - Uso Livre - Proibida a venda" & vbCrLf & "Desenvolvido por Thiago Amaral Guarnieri" _
    & vbCrLf & vbCrLf & "Para sugestões, dúvidas ou caso considere fazer uma contribuição, favor contactar pelo email torinokinsei@hotmail.com."
End Sub
