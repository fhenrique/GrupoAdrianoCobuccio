Attribute VB_Name = "Modulo"
Option Explicit

'Função para leitura das informações do arquivo INI
Public Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
    
Public strServidor As String
Public strDatabase As String
Public strUsuario As String
Public strSenha As String
Public strTrasacaoSelecionada As Integer

Public rsRecordSet As ADODB.Recordset

Public Enum acao
    [incluindo] = 0
    [editando] = 1
    [excluindo] = 2
    [cancelando] = 3
    [salvando] = 4
    [navegando] = 5
End Enum


Public strAcao As acao

'Função para buscar informações no arquivo INI
Public Function GetINISetting(ByVal sHeading As String, _
    ByVal sKey As String, _
    sINIFileName) As String
     
    Const cparmLen = 50
    Dim sReturn As String * cparmLen
    Dim sDefault As String * cparmLen
    Dim lLength As Long
    lLength = GetPrivateProfileString(sHeading, _
            sKey, _
            sDefault, _
            sReturn, _
            cparmLen, _
            sINIFileName)
             
    GetINISetting = Mid(sReturn, 1, lLength)
End Function

Public Sub Arquivo_Dados()
    Dim ConfigFile As String
    Dim atalho_track1 As String
    Dim strTemp As String

    ConfigFile = App.Path & "\config.INI"
    
    strServidor = GetINISetting("BANCODEDADOS", "Servidor", ConfigFile)
    strDatabase = GetINISetting("BANCODEDADOS", "Database", ConfigFile)
    strUsuario = GetINISetting("BANCODEDADOS", "Usuario", ConfigFile)
    strSenha = GetINISetting("BANCODEDADOS", "Senha", ConfigFile)
End Sub

Function ValidaData(dataTexto As String) As Boolean
    If IsDate(dataTexto) Then
        ValidaData = True
    Else
        ValidaData = False
    End If
End Function


Public Sub inserir(NumCart As String, ValTransac As Double, dtTransac As Date, DescTransac As String)
' Para a conexão com o SQL Server, está referenciada a "Microsoft ActiveX Data Objects 2.8 Library"

    Dim conn As ADODB.Connection
    Dim strConn As String
    Dim strSQL As String
    
    strConn = "Provider=SQLOLEDB;Data Source=" & strServidor & ";Initial Catalog=" & strDatabase & ";User Id=" & strUsuario & ";Password=" & strSenha & ";"
    Set conn = New ADODB.Connection
    
    On Error GoTo ErrorHandler
    
    conn.Open strConn
    
    strSQL = "Insert Into Transacao(Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao) Values('" & NumCart & "'," & Replace(ValTransac, ",", ".") & ",'" & dtTransac & "', '" & DescTransac & "')"
    
    conn.Execute strSQL
    
    conn.Close
    Set conn = Nothing
    
    MsgBox "Transação inserida com sucesso!", vbOKOnly + vbInformation, "Informação"
    
    With frmPrincipal
        .btnNovo.Enabled = True
        .btnEditar.Enabled = True
        .btnExcluir.Enabled = True
        .btnSalvar.Enabled = False
        .btnCancelar.Enabled = False
    End With
    
    Exit Sub

ErrorHandler:
    MsgBox "Erro: " & Err.Description, vbCritical
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
        Set conn = Nothing
    End If
End Sub

Public Sub editar(NumCart As String, ValTransac As Double, dtTransac As Date, DescTransac As String)
Dim conn As ADODB.Connection
Dim strConn As String
Dim sql As String

strConn = "Provider=SQLOLEDB;Data Source=" & strServidor & ";Initial Catalog=" & strDatabase & ";User Id=" & strUsuario & ";Password=" & strSenha & ";"

Set conn = New ADODB.Connection

On Error GoTo ErrorHandler

conn.Open strConn

sql = "UPDATE Transacao SET Numero_Cartao = '" & NumCart & "', Valor_Transacao = " & Replace(ValTransac, ",", ".") & ", Data_Transacao = '" & dtTransac & "', Descricao = '" & DescTransac & "' Where Id_Transacao = " & strTrasacaoSelecionada

conn.Execute sql

MsgBox "Registro atualizado com sucesso!", vbInformation

conn.Close
Set conn = Nothing
Exit Sub

ErrorHandler:
    MsgBox "Erro: " & Err.Description, vbCritical
    If conn.State = adStateOpen Then conn.Close
    Set conn = Nothing
End Sub

Public Sub excluir()
    Dim cn As ADODB.Connection
    Dim strSQL As String
    Dim idExcluir As Integer
    
    Set cn = New ADODB.Connection
    cn.ConnectionString = "Provider=SQLOLEDB;Data Source=" & strServidor & ";Initial Catalog=" & strDatabase & ";User Id=" & strUsuario & ";Password=" & strSenha & ";"
    cn.Open
    
    idExcluir = strTrasacaoSelecionada
    
    strSQL = "DELETE FROM Transacao WHERE id_Transacao = " & idExcluir
    
    On Error Resume Next
    cn.Execute strSQL
    
    If Err.Number = 0 Then
        atualizaDataGrid
        MsgBox "Registro excluído com sucesso!", vbInformation
        strTrasacaoSelecionada = 0
        setaAcao ("")
    Else
        MsgBox "Erro: " & Err.Description, vbCritical
    End If
    
    cn.Close
    Set cn = Nothing
End Sub

Public Sub setaAcao(strTexto As String)
    frmPrincipal.lblAcao.Caption = strTexto
End Sub

Public Sub atualizaDataGrid()
    Dim conn As ADODB.Connection
    Dim strConn As String
    Dim strSQL As String
    
    Set conn = New ADODB.Connection
    
    strConn = "Provider=SQLOLEDB;Data Source=" & strServidor & ";Initial Catalog=" & strDatabase & ";User Id=" & strUsuario & ";Password=" & strSenha & ";"
    
    conn.Open strConn
    
    Set rsRecordSet = New ADODB.Recordset
    strSQL = "SELECT * FROM Transacao"
    rsRecordSet.Open strSQL, conn, adOpenStatic, adLockReadOnly
    
    Set frmPrincipal.dgLogTransacoes.DataSource = rsRecordSet
    
    If rsRecordSet.RecordCount = 0 Then
        strTrasacaoSelecionada = 0
    End If
    
    frmPrincipal.dgLogTransacoes.Columns(3).NumberFormat = "dd/mm/yyyy"
    frmPrincipal.dgLogTransacoes_Click
    
    frmPrincipal.dgLogTransacoes.Columns(0).Width = 0
End Sub

