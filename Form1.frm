VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ControlID"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   5190
      Left            =   2475
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "Form1.frx":0000
      Top             =   375
      Width           =   4290
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Ler Empregador"
      Height          =   420
      Left            =   300
      TabIndex        =   8
      Top             =   4125
      Width           =   1650
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ler Informações"
      Height          =   420
      Left            =   300
      TabIndex        =   7
      Top             =   3675
      Width           =   1650
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Sair"
      Height          =   420
      Left            =   300
      TabIndex        =   5
      Top             =   4575
      Width           =   1650
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ajustar Data e Hora"
      Height          =   420
      Left            =   300
      TabIndex        =   4
      Top             =   3225
      Width           =   1650
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Desconectar"
      Height          =   420
      Left            =   300
      TabIndex        =   3
      Top             =   2775
      Width           =   1650
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   300
      TabIndex        =   1
      Text            =   "192.168.1.200"
      Top             =   1875
      Width           =   1590
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar"
      Height          =   420
      Left            =   300
      TabIndex        =   0
      Top             =   2325
      Width           =   1650
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "www.liondas.com.br"
      Height          =   240
      Left            =   300
      TabIndex        =   6
      Top             =   5475
      Width           =   1590
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   600
      Picture         =   "Form1.frx":0006
      Stretch         =   -1  'True
      Top             =   375
      Width           =   990
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   -75
      X2              =   7650
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IP do Equipamento"
      Height          =   240
      Left            =   375
      TabIndex        =   2
      Top             =   1575
      Width           =   1440
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   5490
      Left            =   2325
      Top             =   225
      Width           =   4590
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim rep As New RepCid.RepCid
Dim er As ErrosRep
Dim gravou As Boolean
Dim log As String

'Inicio do programa
Private Sub Form_Load()
  ModoDesconectado
  Text2.Text = "Clique em 'Conectar' para iniciar"
End Sub

'Fim do programa
Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

'Finaliza o programa
Private Sub Command4_Click()
  Unload Me
End Sub

'Conectar
Private Sub Command1_Click()
  Mensagem "Conectando em " & Text1.Text & "..."
  'Tenta a conexão utilizando a senha Web padrão
  er = rep.iDClass_Conectar(Text1.Text, "admin", "admin")
  If er = ErrosRep_OK Then
     'Conexao com sucesso
     Mensagem "Conectado em " & Text1.Text
     ModoConectado
  Else
    'Tratamento do erro de conexão
    Select Case Err
      Case ErrosRep_ErroAutenticacao
        log = "Autenticação"
      Case ErrosRep_ErroConexao
        log = "Conexão"
      Case ErrosRep_ErroNaoOcioso
        log = "Não Ocioso"
      Case ErrosRep_ErroOutro
        log = "Outro"
      Case Else
        log = "Desconhecido"
    End Select
    Mensagem "Erro " & log
  End If
End Sub

'Desconectar
Private Sub Command2_Click()
  rep.Desconectar
  Mensagem "Desconectado"
  ModoDesconectado
End Sub

'Ajustar data e hora com a hora do computador
Private Sub Command3_Click()
  Mensagem "Ajustando data e hora..."
  Dim dia As Long, mes As Long, ano As Long, hora As Long, minuto As Long, segundo As Long
  dia = Format(Date, "dd")
  mes = Format(Date, "mm")
  ano = Format(Date, "yyyy")
  hora = Format(Time, "hh")
  minuto = Format(Time, "nn")
  segundo = Format(Time, "ss")
  If rep.GravarDataHora(ano, mes, dia, hora, minuto, segundo, gravou) Then
    Mensagem "Data e hora ajustada com sucesso"
  Else
    rep.GetLastLog log
    Mensagem log
  End If
End Sub

'Receber informações do rep
Private Sub Command5_Click()
  Dim sn As String, tam As String, ret As String, up As String, cor As String, pap As String, nsr As String
  Mensagem "Recebendo informações..."
  If rep.LerInfo_vb6(sn, tam, ret, up, cor, pap, nsr) Then
    Mensagem "Informações Recebidas: " & vbCrLf & _
               "  serial: " & sn & vbCrLf & _
               "  tamanho bobina " & tam & vbCrLf & _
               "  restante bobina " & ret & vbCrLf & _
               "  tempo ligado " & up & vbCrLf & _
               "  cortes " & cor & vbCrLf & _
               "  metros impressos " & pap & vbCrLf & _
               "  nsr atual " & nsr
  Else
    rep.GetLastLog log
    Mensagem log
  End If
End Sub

'Receber informações do empregador
Private Sub Command6_Click()
  Dim doc As String, tipo As Long, cei As String, razao As String, endereco As String
  Mensagem "Recebendo empregador..."
  If rep.LerEmpregador(doc, tipo, cei, razao, endereco) Then
    Mensagem "Empregador Recebido: " & vbCrLf & _
               "  documento: " & doc & vbCrLf & _
               "  tipo: " & tipo & vbCrLf & _
               "  cei: " & cei & vbCrLf & _
               "  razão: " & razao & vbCrLf & _
               "  endereco: " & endereco
  Else
    rep.GetLastLog log
    Mensagem log
  End If
End Sub

'Exibe a mensagem na tela
Private Sub Mensagem(ByVal sMen As String)
  Text2.Text = sMen & vbCrLf & Text2.Text
  DoEvents
End Sub

'Habilita ou desabilita os botões na tela
Private Sub ModoConectado()
  Command1.Enabled = False
  Command2.Enabled = True
  Command3.Enabled = True
  Command4.Enabled = False
  Command5.Enabled = True
  Command6.Enabled = True
End Sub
Private Sub ModoDesconectado()
  Command1.Enabled = True
  Command2.Enabled = False
  Command3.Enabled = False
  Command4.Enabled = True
  Command5.Enabled = False
  Command6.Enabled = False
End Sub


'If rep.iDClass_GravarEmpregador("111111110001110000", 1, "0", "Teste Empregador", "Teste Endereco", "0", gravou) Then
'  Label1.Caption = "Comando executado com sucesso"
'Else
'  rep.GetLastLog log
'  Label1.Caption = log
'End If


