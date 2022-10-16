VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   16185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Exercício 3"
      Height          =   1815
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "Gerar arquivo"
         Height          =   360
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Exercício 1"
      Height          =   1455
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   3135
      Begin VB.ComboBox cboPeriodo 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Exercício 1"
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.Timer tmrTeste 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2640
         Top             =   960
      End
      Begin VB.CommandButton CMDTeste 
         Caption         =   "&Teste"
         Height          =   360
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   990
      End
      Begin VB.Label lblTeste 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Width           =   465
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Option Explicit

Dim contador As Integer
Dim contSegundo As Integer
Dim Periodo As Integer


Private Sub cboPeriodo_Click()
    If cboPeriodo.ListIndex = 0 Then
        Periodo = 1
    ElseIf cboPeriodo.ListIndex = 1 Then
        Periodo = 7
    ElseIf cboPeriodo.ListIndex = 2 Then
        Periodo = 15
    ElseIf True Then
        Periodo = 30
    End If
End Sub

Private Sub CMDTeste_Click()
    tmrTeste.Enabled = True
    tmrTeste.Interval = 1000
    tmrTeste_Timer
End Sub

Private Sub Command1_Click()

GeraCadastro

End Sub

Private Sub Form_Load()
    contador = 0
    lblTeste.Caption = "Clique para iniciar o contador"
    
    cboPeriodo.AddItem "Diário"
    cboPeriodo.AddItem "Semanal"
    cboPeriodo.AddItem "Quinzenal"
    cboPeriodo.AddItem "Mensal"
    
End Sub

Private Sub tmrTeste_Timer()
    contador = contador + 1
    
    If contador = 1 Then
        lblTeste.Caption = "Processo"
    ElseIf contador = 2 Then
        lblTeste.Caption = "em"
    ElseIf contador = 3 Then
        lblTeste.Caption = "teste"
    Else
        Exit Sub
    End If
    
End Sub

Public Function GeraCadastro()

    Dim Conexao  As New ADODB.Connection
    Dim Caminho  As String
    Dim Provedor As String
    Dim Linha    As String
    Dim ConsultaSql      As String
    Dim RsTemp   As New ADODB.Recordset
  
    Dim Arquivo  As String
    Dim NumArq   As Integer
    Dim Data     As String
 
    Data = Format(Now, "dd/mm/yyyy")
    Arquivo = "c:\cadastro.txt"
    NumArq = FreeFile
 
    Open Arquivo For Output As #NumArq
 
    Provedor = "Provider=Microsoft.Jet.OLEDB.4.0;Password="""";" & "Data Source=" & "c:\BANCO.MDB;Persist Security Info=True"
    ConsultaSql = "SELECT * FROM MSFUNCIONARIOS ORDER BY FunNom"

    RsTemp.CursorLocation = adUseClient
    RsTemp.Open ConsultaSql, Provedor, adOpenStatic
   
    Set RsTemp.ActiveConnection = Nothing
   
    If RsTemp.RecordCount > 0 Then
        Do While Not RsTemp.EOF
            Print #NumArq, Tab(1); Left(RsTemp!FunNom, 30);
            Print #NumArq, Tab(31); Space(1);
            Print #NumArq, Tab(32); Format(RsTemp!FunDatAdm, "ddmmYYYY");
            Print #NumArq, Tab(42); Space(1);
            Print #NumArq, Tab(43); Left(Format(RsTemp!FunSal, "##,##0.00"), 54);
           
           ' Print #NumArq, Tab(1); Left("EDUARDO_NUNES_DE_SOUZA_SEIXASssss", 30);
           ' Print #NumArq, Tab(31); Space(1);
           ' Print #NumArq, Tab(32); Format(Data, "dd/mm/YYYY");
           ' Print #NumArq, Tab(42); Space(1);
          '  Print #NumArq, Tab(43); Format("15000,98", "##,##0.00");
            RsTemp.MoveNext
        Loop
        Close #NumArq
    End If

End Function

