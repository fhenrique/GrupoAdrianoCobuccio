VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XYZ Cartões de Crédito"
   ClientHeight    =   8340
   ClientLeft      =   5055
   ClientTop       =   1500
   ClientWidth     =   9000
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Ações:"
      Height          =   885
      Left            =   75
      TabIndex        =   22
      Top             =   7080
      Width           =   8880
      Begin MSComDlg.CommonDialog dlgExportacao 
         Left            =   6675
         Top             =   195
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton btnExcluir 
         Height          =   405
         Left            =   2580
         Picture         =   "frmPrincipal.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Excluir uma transação."
         Top             =   255
         Width           =   990
      End
      Begin VB.CommandButton btnEditar 
         Height          =   405
         Left            =   1380
         Picture         =   "frmPrincipal.frx":10D6E
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Editar uma transação já inserida."
         Top             =   255
         Width           =   990
      End
      Begin VB.CommandButton btnNovo 
         Height          =   405
         Left            =   210
         Picture         =   "frmPrincipal.frx":112FC
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Inserir uma nova transação."
         Top             =   255
         Width           =   945
      End
      Begin VB.CommandButton btnExportar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exportar para Excel"
         Height          =   720
         Left            =   7245
         Picture         =   "frmPrincipal.frx":1182E
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   1560
      End
      Begin VB.CommandButton btnCancelar 
         Enabled         =   0   'False
         Height          =   405
         Left            =   5085
         Picture         =   "frmPrincipal.frx":11B55
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Cancelar a ação atual."
         Top             =   255
         Width           =   1230
      End
      Begin VB.CommandButton btnSalvar 
         Enabled         =   0   'False
         Height          =   405
         Left            =   3855
         Picture         =   "frmPrincipal.frx":1228B
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Salvar a ação atual."
         Top             =   255
         Width           =   990
      End
   End
   Begin VB.Frame panConsulta 
      Caption         =   "Consultar:"
      Height          =   705
      Left            =   90
      TabIndex        =   14
      Top             =   6285
      Width           =   8895
      Begin VB.OptionButton optNumeroCartao 
         Caption         =   "Número do cartão:"
         Height          =   285
         Left            =   60
         TabIndex        =   21
         ToolTipText     =   "Consultar por número do cartão."
         Top             =   315
         Width           =   1650
      End
      Begin VB.TextBox txtOptnNumeroCartao 
         Height          =   285
         Left            =   1725
         TabIndex        =   20
         Top             =   300
         Width           =   1650
      End
      Begin VB.OptionButton optDtTransacao 
         Caption         =   "Data da Transação:"
         Height          =   225
         Left            =   3480
         TabIndex        =   19
         ToolTipText     =   "Consultar por data da transação."
         Top             =   315
         Width           =   1755
      End
      Begin VB.TextBox txtOptnDataTransacao 
         Height          =   285
         Left            =   5220
         TabIndex        =   18
         Top             =   285
         Width           =   1095
      End
      Begin VB.OptionButton optValTransacao 
         Caption         =   "Valor:"
         Height          =   225
         Left            =   6435
         TabIndex        =   17
         ToolTipText     =   "Consultoar por Valor da Transação."
         Top             =   315
         Width           =   705
      End
      Begin VB.TextBox txtOptnValorTransacao 
         Height          =   285
         Left            =   7185
         TabIndex        =   16
         Top             =   285
         Width           =   1065
      End
      Begin VB.CommandButton btnConsultar 
         Height          =   360
         Left            =   8310
         Picture         =   "frmPrincipal.frx":127F7
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Clique para consultar"
         Top             =   210
         Width           =   435
      End
   End
   Begin VB.Frame panHistTransacoes 
      Caption         =   "Histórico de Transações"
      Height          =   3300
      Left            =   45
      TabIndex        =   12
      Top             =   2880
      Width           =   8910
      Begin VB.CommandButton btnTeste 
         Caption         =   "Categorias"
         Height          =   390
         Left            =   7560
         TabIndex        =   29
         Top             =   2730
         Width           =   1185
      End
      Begin MSDataGridLib.DataGrid dgLogTransacoes 
         Height          =   2940
         Left            =   75
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Lista de transações já inseridas."
         Top             =   240
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   5186
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2250
      Left            =   45
      TabIndex        =   7
      Top             =   555
      Width           =   8895
      Begin VB.TextBox txtNumeroCartao 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1560
         MaxLength       =   16
         TabIndex        =   1
         ToolTipText     =   "Digite o número do cartão."
         Top             =   825
         Width           =   1650
      End
      Begin VB.TextBox txtDataTransacao 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1575
         MaxLength       =   10
         TabIndex        =   0
         ToolTipText     =   "Digite a data no formato DD/MM/AAAA. Ex: 22/04/1980"
         Top             =   255
         Width           =   1095
      End
      Begin VB.TextBox txtValorTransacao 
         Height          =   300
         Left            =   4935
         TabIndex        =   2
         ToolTipText     =   "Digite o valor da transação."
         Top             =   825
         Width           =   1065
      End
      Begin VB.TextBox txtDescricao 
         Height          =   615
         Left            =   105
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Digite a descrição da transação."
         Top             =   1530
         Width           =   8625
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data da Transação:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   315
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número do Cartão:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   885
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Valor da Transação:"
         Height          =   195
         Left            =   3465
         TabIndex        =   9
         Top             =   870
         Width           =   1440
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1290
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Height          =   645
      Left            =   -15
      TabIndex        =   4
      Top             =   -75
      Width           =   8970
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         Caption         =   "Gerenciamento de Transações"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   135
         TabIndex        =   5
         Top             =   135
         Width           =   8025
      End
   End
   Begin VB.Label lblAcao 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   75
      TabIndex        =   6
      Top             =   8025
      Width           =   8835
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnFind As ADODB.Connection
Dim rsFind As ADODB.Recordset
Dim strConnFind As String
Dim sqlFind As String

Private Sub btnCancelar_Click()
    habilitaBotoes (acao.cancelando)
    habilitaCampos (False)
    
    panConsulta.Enabled = True
    
    setaAcao ("")
End Sub

Private Sub btnConsultar_Click()
    If (RTrim(LTrim(txtOptnNumeroCartao.Text)) = "") And ((RTrim(LTrim(txtOptnDataTransacao.Text)) = "")) And ((RTrim(LTrim(txtOptnValorTransacao.Text)) = "")) Then
        MsgBox "Informe uma das 3 opções de consulta:" & Chr(13) & Chr(13) & "- Número do cartão" & Chr(13) & "- Data da Transação" & Chr(13) & "- Valor", vbInformation
        Exit Sub
    End If
    
    
    Dim valorFormatado As String
    
    If optNumeroCartao Then
        sqlFind = "SELECT * FROM Transacao WHERE Numero_Cartao = " & txtOptnNumeroCartao.Text
    ElseIf optDtTransacao Then
        sqlFind = "SELECT * FROM Transacao WHERE Data_Transacao= '" & txtOptnDataTransacao.Text & "'"
    ElseIf optValTransacao Then
        If Len(txtOptnValorTransacao.Text) < 7 Then
            valorFormatado = Replace(txtOptnValorTransacao.Text, ",", ".")
            sqlFind = "SELECT * FROM Transacao WHERE Valor_Transacao = " & valorFormatado
        Else
            valorFormatado = Replace(txtOptnValorTransacao.Text, ".", "")
            valorFormatado = Replace(valorFormatado, ",", ".")
            sqlFind = "SELECT Id_Transacao, Numero_Cartao, Format(Valor_Transacao,'N', 'pt-BR') as Valor_Transacao, Data_Transacao, Descricao FROM Transacao WHERE Valor_Transacao = " & valorFormatado
        End If

    End If
    
    Set rsFind = New ADODB.Recordset
    rsFind.Open sqlFind, cnFind, adOpenStatic, adLockOptimistic
    
    Set dgLogTransacoes.DataSource = rsFind
    dgLogTransacoes.Columns(3).NumberFormat = "dd/mm/yyyy"
    
End Sub

Private Sub btnEditar_Click()
    
    If strTrasacaoSelecionada = 0 Then
        MsgBox "É necessário selecionar uma transação no Histórico de Transações!", vbOKOnly + vbInformation, "Informar Transação"
        Exit Sub
    End If
    
    habilitaCampos (True)
    habilitaBotoes (acao.editando)
    
    strAcao = editando
    

End Sub

Private Sub btnExcluir_Click()
    
    If strTrasacaoSelecionada = 0 Then
        MsgBox "É necessário selecionar uma transação no Histórico de Transações!", vbOKOnly + vbInformation, "Informar Transação"
        Exit Sub
    End If
    
    
    If MsgBox("Confirma a exclusão desta transação?", vbYesNo + vbExclamation, "Confirmação") = vbYes Then
        habilitaBotoes (acao.excluindo)
        strAcao = excluindo
        excluir
    End If
    
End Sub

Private Sub btnExportar_Click()
    Dim connExporta As ADODB.Connection
    Dim rsExporta As ADODB.Recordset
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim contador1 As Integer
    Dim strCaminho As String
    
    ' Conecta ao banco de dados
    Set connExporta = New ADODB.Connection
    connExporta.Open "Provider=SQLOLEDB;Data Source=" & strServidor & ";Initial Catalog=" & strDatabase & ";User Id=" & strUsuario & ";Password=" & strSenha & ";"
    
    ' Executa a consulta
    Set rsExporta = New ADODB.Recordset
    rsExporta.Open "SELECT Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, dbo.fn_CategoriaTransacao(Valor_Transacao) AS Categoria FROM Transacao WHERE Data_Transacao >= DATEADD(month, -1, GETDATE())", connExporta, adOpenStatic, adLockReadOnly
    
    ' Cria o objeto Excel
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    ' Define o título das colunas
    xlSheet.Cells(1, 1).Value = "Numero_Cartao"
    xlSheet.Cells(1, 2).Value = "Valor_Transacao"
    xlSheet.Cells(1, 3).Value = "Data_Transacao"
    xlSheet.Cells(1, 4).Value = "Descricao"
    xlSheet.Cells(1, 5).Value = "Categoria"
    
    ' Preenche os dados
    contador1 = 2
    While Not rsExporta.EOF
        xlSheet.Cells(contador1, 1).Value = rsExporta!Numero_Cartao
        xlSheet.Cells(contador1, 2).Value = rsExporta!Valor_Transacao
        xlSheet.Cells(contador1, 3).Value = rsExporta!Data_Transacao
        xlSheet.Cells(contador1, 4).Value = rsExporta!Descricao
        xlSheet.Cells(contador1, 5).Value = rsExporta!Categoria
        contador1 = contador1 + 1
        rsExporta.MoveNext
    Wend
    
    ' Abre a caixa de diálogo para salvar o arquivo
    dlgExportacao.Filter = "Arquivos do Excel(*.xlsx)|*.xls"
    dlgExportacao.DefaultExt = "xlsx"
    
    dlgExportacao.ShowSave
    
    strCaminho = dlgExportacao.FileName
    
    
    'strCaminho = Application.GetSaveAsFilename(FileFilter:="Arquivos Excel (*.xls), *.xls", Title:="Salvar arquivo Excel")
    
    If strCaminho <> "" Then
        xlBook.SaveAs strCaminho
        MsgBox "Arquivo salvo com sucesso em: " & strCaminho, vbInformation, "Exportação"
    Else
        MsgBox "Exportação cancelada!", vbExclamation, "Exportação"
    End If
    
    ' Fecha os objetos
    xlBook.Close
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    rsExporta.Close
    connExporta.Close
    Set rsExporta = Nothing
    Set connExporta = Nothing
End Sub

Private Sub btnNovo_Click()
    habilitaCampos (True)
    habilitaBotoes (acao.incluindo)
    
    strAcao = incluindo
    
    panConsulta.Enabled = False
    
    setaAcao " Inserindo nova transação!"
        
End Sub

Private Sub btnSalvar_Click()
    If consisteCampos Then
        If strAcao = incluindo Then
            inserir txtNumeroCartao.Text, txtValorTransacao.Text, txtDataTransacao.Text, txtDescricao.Text
            habilitaCampos (False)
            habilitaBotoes (salvando)
            setaAcao ("")
            limpaCampos
            atualizaDataGrid
        ElseIf strAcao = editando Then
            editar txtNumeroCartao.Text, txtValorTransacao.Text, txtDataTransacao.Text, txtDescricao.Text
            habilitaCampos (False)
            setaAcao ("")
            atualizaDataGrid
        End If
        
        btnSalvar.Enabled = False
        panConsulta.Enabled = True
    End If
End Sub



Private Sub btnTeste_Click()
    sqlFind = "SELECT  Numero_Cartao, Valor_Transacao, dbo.fn_CategoriaTransacao(Valor_Transacao) AS Categoria FROM Transacao;"
    
    Set rsFind = New ADODB.Recordset
    rsFind.Open sqlFind, cnFind, adOpenStatic, adLockOptimistic
    
    Set dgLogTransacoes.DataSource = rsFind
End Sub

Public Sub dgLogTransacoes_Click()
    If strTrasacaoSelecionada > 0 Then
        setaAcao (" Navegando.")
    End If
End Sub

Public Sub dgLogTransacoes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    
    With dgLogTransacoes
            strAcao = navegando
            strTrasacaoSelecionada = dgLogTransacoes.Columns("Id_Transacao").Text
            txtDataTransacao.Text = Format$(.Columns("Data_Transacao").Text, "dd/mm/yyyy")
            txtNumeroCartao.Text = .Columns("Numero_Cartao").Text
            txtValorTransacao.Text = Format(.Columns("Valor_Transacao").Text, "#,##0.00")
            txtDescricao.Text = .Columns("Descricao").Text
    End With
End Sub

Private Sub Form_Load()
    'frmPrincipal.dgLogTransacoes.Columns(0).Width = 0
    
    Arquivo_Dados
    habilitaCampos (False)
    txtDataTransacao.Text = Date
    
    atualizaDataGrid
    
    If rsRecordSet.RecordCount > 0 Then
        strTrasacaoSelecionada = dgLogTransacoes.Columns("Id_Transacao").Text
    Else
        strTrasacaoSelecionada = 0
    End If

    strTrasacaoSelecionada = 0
    
    ' 1. Configurar a conexão
    Set cnFind = New ADODB.Connection
    strConnFind = "Provider=SQLOLEDB;Data Source=" & strServidor & ";Initial Catalog=" & strDatabase & ";User Id=" & strUsuario & ";Password=" & strSenha & ";"
    cnFind.Open strConnFind
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
     ' Fechar conexão ao sair
    If Not rsFind Is Nothing Then
        If rsFind.State = adStateOpen Then rsFind.Close
        Set rsFind = Nothing
    End If
    If Not cnFind Is Nothing Then
        If cnFind.State = adStateOpen Then cnFind.Close
        Set cnFind = Nothing
    End If
End Sub



Private Sub optDtTransacao_Click()
    txtOptnDataTransacao.BackColor = &HFFC0C0
    txtOptnNumeroCartao.BackColor = &HFFFFFF
    txtOptnValorTransacao.BackColor = &HFFFFFF
End Sub

Private Sub optNumeroCartao_Click()
    txtOptnNumeroCartao.BackColor = &HFFC0C0
    txtOptnDataTransacao.BackColor = &HFFFFFF
    txtOptnValorTransacao.BackColor = &HFFFFFF
End Sub

Private Sub optValTransacao_Click()
    txtOptnValorTransacao.BackColor = &HFFC0C0
    txtOptnNumeroCartao.BackColor = &HFFFFFF
    txtOptnDataTransacao.BackColor = &HFFFFFF
End Sub

Private Sub txtDataTransacao_GotFocus()
    txtDataTransacao.BackColor = &HFFC0C0
End Sub

Private Sub txtDataTransacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNumeroCartao.SetFocus
    End If
End Sub

Private Sub txtDataTransacao_LostFocus()

    If (RTrim(LTrim(txtDataTransacao.Text <> ""))) Then
        If (Not ValidaData(txtDataTransacao.Text)) Then
            MsgBox "Informe uma data válida! Ex: " & Date, vbInformation
            txtDataTransacao.Text = ""
            txtDataTransacao.SetFocus
        Else
            txtDataTransacao.BackColor = &HFFFFFF
        End If
    End If
    
    txtDataTransacao.BackColor = &HFFFFFF
    
End Sub


Private Sub txtDescricao_GotFocus()
    txtDescricao.BackColor = &HFFC0C0
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDataTransacao.SetFocus
    End If
End Sub

Private Sub txtDescricao_LostFocus()
    txtDescricao.BackColor = &HFFFFFF
End Sub

Private Sub txtNumeroCartao_GotFocus()
    txtNumeroCartao.BackColor = &HFFC0C0
End Sub

Private Sub txtNumeroCartao_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
        Case 13
        Case 48 To 57
        Case Else
            MsgBox "Este campo aceita somente números!", vbInformation
            KeyAscii = 0
    End Select
    
    If KeyAscii = 13 Then
        txtValorTransacao.SetFocus
    End If
End Sub

Private Sub txtNumeroCartao_LostFocus()
    txtNumeroCartao.BackColor = &HFFFFFF
End Sub

Private Sub txtOptnDataTransacao_GotFocus()
    optDtTransacao.Value = 1
    txtOptnDataTransacao.BackColor = &HFFC0C0
    txtOptnNumeroCartao.BackColor = &HFFFFFF
    txtOptnValorTransacao.BackColor = &HFFFFFF
End Sub

Private Sub txtOptnDataTransacao_LostFocus()
    With txtOptnDataTransacao
        If (RTrim(LTrim(.Text)) <> "") And (Not ValidaData(.Text)) Then
            MsgBox "Informe uma data válida! Ex: " & Date, vbInformation
            .Text = ""
            .SetFocus
        Else
            .BackColor = &HFFFFFF
        End If
    End With
End Sub

Private Sub txtOptnNumeroCartao_GotFocus()
    optNumeroCartao.Value = 1
    txtOptnNumeroCartao.BackColor = &HFFC0C0
    txtOptnDataTransacao.BackColor = &HFFFFFF
    txtOptnValorTransacao.BackColor = &HFFFFFF
End Sub

Private Sub txtOptnNumeroCartao_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
        Case 13
        Case 48 To 57
        Case Else
            MsgBox "Este campo aceita somente números!", vbInformation
            KeyAscii = 0
    End Select
End Sub

Private Sub txtOptnValorTransacao_Change()
    Dim ValorTexto As String
    ValorTexto = Replace(txtOptnValorTransacao.Text, "", "")
    ValorTexto = Replace(ValorTexto, ".", "")
    ValorTexto = Replace(ValorTexto, ",", "")
    
    With txtOptnValorTransacao
        If IsNumeric(ValorTexto) And ValorTexto <> "" Then
            .Text = Format(CDbl(ValorTexto) / 100, "#,##0.00")
            .SelStart = Len(.Text)
        End If
    End With
End Sub

Private Sub txtOptnValorTransacao_GotFocus()
    optValTransacao.Value = 1
    txtOptnValorTransacao.BackColor = &HFFC0C0
    txtOptnNumeroCartao.BackColor = &HFFFFFF
    txtOptnDataTransacao.BackColor = &HFFFFFF
End Sub

Private Sub txtValorTransacao_Change()
    
    If strAcao = incluindo Then
        Dim ValorTexto As String
        ValorTexto = Replace(txtValorTransacao.Text, "", "")
        ValorTexto = Replace(ValorTexto, ".", "")
        ValorTexto = Replace(ValorTexto, ",", "")
        
        If IsNumeric(ValorTexto) And ValorTexto <> "" Then
            txtValorTransacao.Text = Format(CDbl(ValorTexto) / 100, "#,##0.00")
            txtValorTransacao.SelStart = Len(txtValorTransacao.Text)
        End If
    End If
End Sub

Private Sub txtValorTransacao_GotFocus()
    txtValorTransacao.BackColor = &HFFC0C0
End Sub

Private Sub txtValorTransacao_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
        Case 13
        Case 44
        Case 46
        Case 48 To 57
        Case Else
            MsgBox "Este campo aceita somente números!", vbInformation
            KeyAscii = 0
    End Select
    
    If KeyAscii = 13 Then
        txtDescricao.SetFocus
    End If
End Sub

Private Sub txtValorTransacao_LostFocus()
    txtValorTransacao.BackColor = &HFFFFFF
End Sub

Private Sub habilitaCampos(blnAcao As Boolean)
    txtDataTransacao.Enabled = blnAcao
    txtNumeroCartao.Enabled = blnAcao
    txtValorTransacao.Enabled = blnAcao
    txtDescricao.Enabled = blnAcao
    
    If blnAcao Then
        txtDataTransacao.SetFocus
    End If
End Sub

Private Sub habilitaBotoesClick(varAcao As acao)

    If varAcao = incluindo Then
        btnNovo.Enabled = False
        btnEditar.Enabled = False
        btnExcluir.Enabled = False
        btnSalvar.Enabled = True
        btnCancelar.Enabled = True
        btnConsultar.Enabled = False
        btnExportar.Enabled = False
        txtDataTransacao.Text = Date
        txtNumeroCartao.Text = ""
        txtValorTransacao.Text = ""
        txtDescricao.Text = ""
        dgLogTransacoes.Enabled = False
    ElseIf varAcao = editando Then
        btnNovo.Enabled = False
        btnEditar.Enabled = False
        btnExcluir.Enabled = False
        btnSalvar.Enabled = True
        btnCancelar.Enabled = True
        btnConsultar.Enabled = False
        btnExportar.Enabled = False
        panConsulta.Enabled = False
        setaAcao " Editando transação já inserida!"
    ElseIf varAcao = excluindo Then
            btnNovo.Enabled = True
            btnEditar.Enabled = True
            btnExcluir.Enabled = True
            btnConsultar.Enabled = True
            btnExportar.Enabled = True
            panConsulta.Enabled = True
            limpaCampos
    ElseIf varAcao = cancelando Then
        btnNovo.Enabled = True
        btnEditar.Enabled = True
        btnExcluir.Enabled = True
        btnSalvar.Enabled = False
        btnCancelar.Enabled = False
        btnConsultar.Enabled = True
        btnExportar.Enabled = True
        dgLogTransacoes.Enabled = True
        strTrasacaoSelecionada = 0
    ElseIf varAcao = salvando Then
        btnNovo.Enabled = True
        btnEditar.Enabled = True
        btnExcluir.Enabled = True
        btnSalvar.Enabled = False
        btnCancelar.Enabled = False
        btnConsultar.Enabled = True
        btnExportar.Enabled = True
        dgLogTransacoes.Enabled = True
    End If

End Sub

Private Sub habilitaBotoes(varAcao As Integer)

    If varAcao = incluindo Then
        btnNovo.Enabled = False
        btnEditar.Enabled = False
        btnExcluir.Enabled = False
        btnSalvar.Enabled = True
        btnCancelar.Enabled = True
        btnConsultar.Enabled = False
        btnExportar.Enabled = False
        txtDataTransacao.Text = Date
        txtNumeroCartao.Text = ""
        txtValorTransacao.Text = ""
        txtDescricao.Text = ""
        dgLogTransacoes.Enabled = False
    ElseIf varAcao = editando Then
        btnNovo.Enabled = False
        btnEditar.Enabled = False
        btnExcluir.Enabled = False
        btnSalvar.Enabled = True
        btnCancelar.Enabled = True
        btnConsultar.Enabled = False
        btnExportar.Enabled = False
        panConsulta.Enabled = False
        setaAcao " Editando transação já inserida!"
    ElseIf varAcao = excluindo Then
            btnNovo.Enabled = True
            btnEditar.Enabled = True
            btnExcluir.Enabled = True
            btnConsultar.Enabled = True
            btnExportar.Enabled = True
            panConsulta.Enabled = True
            limpaCampos
    ElseIf varAcao = cancelando Then
        btnNovo.Enabled = True
        btnEditar.Enabled = True
        btnExcluir.Enabled = True
        btnSalvar.Enabled = False
        btnCancelar.Enabled = False
        btnConsultar.Enabled = True
        btnExportar.Enabled = True
        dgLogTransacoes.Enabled = True
        strTrasacaoSelecionada = 0
    ElseIf varAcao = salvando Then
        btnNovo.Enabled = True
        btnEditar.Enabled = True
        btnExcluir.Enabled = True
        btnSalvar.Enabled = False
        btnCancelar.Enabled = False
        btnConsultar.Enabled = True
        btnExportar.Enabled = True
        dgLogTransacoes.Enabled = True
    End If

End Sub



Private Function consisteCampos() As Boolean
    
    If (RTrim(LTrim(txtDataTransacao.Text)) = "") Then
        consisteCampos = False
        MsgBox "Informe a data da transação!", vbInformation
        txtDataTransacao.SetFocus
        Exit Function
    ElseIf (RTrim(LTrim(txtNumeroCartao.Text)) = "") Then
        consisteCampos = False
        MsgBox "Informe o número do cartão!", vbInformation
        txtNumeroCartao.SetFocus
        Exit Function
    ElseIf (RTrim(LTrim(txtValorTransacao.Text)) = "") Then
        consisteCampos = False
        MsgBox "Informe o valor da transação!", vbInformation
        txtValorTransacao.SetFocus
        Exit Function
    ElseIf (RTrim(LTrim(txtDescricao.Text)) = "") Then
        consisteCampos = False
        MsgBox "Informe a descrição da transação!", vbInformation
        txtDescricao.SetFocus
        Exit Function
    End If
    
    consisteCampos = True

End Function

Private Sub limpaCampos()
    txtDataTransacao.Text = ""
    txtNumeroCartao.Text = ""
    txtValorTransacao.Text = ""
    txtDescricao.Text = ""
End Sub

Private Sub findPorNumero(paramNumeroCartao As Integer)
    Dim cnFind As ADODB.Connection
    Dim rsFind As ADODB.Recordset
    Dim strSQLFind As String
    
    Set cnFind = New ADODB.Connection
    
    Set rsFind = New ADODB.Recordset
    rsFind.CursorLocation = adUseClient
    
    strSQLFind = "SELECT Id_Transacao, Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao FROM Transacao Where Numero_Cartao = " & paramNumeroCartao
    
    cnFind.ConnectionString = "Provider=SQLOLEDB;Data Source=" & strServidor & ";Initial Catalog=" & strDatabase & ";User Id=" & strUsuario & ";Password=" & strSenha & ";"
    
    rsFind.Open strSQLFind, cnFind, adOpenStatic, adLockReadOnly
    Set dgLogTransacoes.DataSource = rsFind
    
    If rsFind.State = adStateOpen Then rsFind.Close
    
    On Error Resume Next
    cnFind.Open
    If cnFind.State = adStateClosed Then
        MsgBox "Erro ao conectar: " & Err.Description
        Exit Sub
    End If

End Sub

    
