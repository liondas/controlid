VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ControlID"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar"
      Height          =   495
      Left            =   1575
      TabIndex        =   0
      Top             =   525
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3300
      TabIndex        =   1
      Top             =   525
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim rep As New RepCid.RepCid
Dim er As ErrosRep
Dim gravou As Boolean
Dim log As String

Private Sub Command1_Click()
  
  Label1.Caption = "Conectando..."
  DoEvents
  
  er = rep.iDClass_Conectar("192.168.1.200", "admin", "admin")
  'er = rep.Conectar_vb6("192.168.1.200", 443, 0)
  
  If er = ErrosRep_OK Then
      
      'Dim doc As String, tipo As Long, cei As String, razao As String, endereco As String
      'If rep.LerEmpregador(doc, tipo, cei, razao, endereco) Then
      '    Label1.Caption = "OK empregador Lido: " + doc + " - " + razao
      'Else
      '    Label1.Caption = "Erro ao ler empregador"
      'End If
      
      'Dim sn As String, tam As String, ret As String, up As String, cor As String, pap As String, nsr As String
      'If rep.LerInfo_vb6(sn, tam, ret, up, cor, pap, nsr) Then
      '    Label1.Caption = "OK Lido: " & sn & " " & tam & " " & ret & " " & up & " " & cor & " " & pap & " " & nsr
      'Else
      '    Label1.Caption = "Erro ao ler info"
      'End If
      
      'If rep.GravarDataHora(2016, 8, 24, 10, 11, 12, lGravou) Then
      '  Label1.Caption = "OK"
      'Else
      '  Label1.Caption = "FALHA"
      'End If
      
      'If rep.iDClass_GravarEmpregador("01.245.055/0001-24", 0, "", "Teste Empregador", "Teste Endereco", "", lGravou) Then
      If rep.iDClass_GravarEmpregador("111111110001110000", 1, "0", "Teste Empregador", "Teste Endereco", "0", gravou) Then
        Label1.Caption = "Comando executado com sucesso"
      Else
        rep.GetLastLog log
        Label1.Caption = log
      End If
      
      
  End If
End Sub

