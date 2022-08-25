Private Sub insertBd(nome As String, cpf As String, valor As Double)

    Dim sql As String
    Dim connection As ADODB.connection
    Set connection = New ADODB.connection
    
    'define a conexão com o banco de dados
    connection.ConnectionString = "DSN=excel_homologa;UID=homologa;PWD=homologa;"
    connection.Open

    'Verifica se obteve sucesso para montar a conexao
    If connection.State = adStateOpen Then
        sql = "INSERT INTO LANC_FOLHA_PAGAMENTO VALUES (" & "'" & nome & "'" & "," & "'" & cpf & "'" & "," & valor & ")"
        
        MsgBox "SQL EXECUTADO: " & sql
    
        connection.Execute sql, regafetados, adExecuteNoRecords
    Else
        MsgBox "Erro ao realizar conexão com o banco", vbCritical
    End If
    
    'Fecha a conexao a cada chamada
    connection.Close

End Sub

Public Sub ImportarArquivoTXT()
    
    'Declarando varivel que ira receber resultado da interação com o arquivo
    Dim Arquivo As String
    Dim Linha As String
    Dim Detalhe As String


    Dim Favorecido As String
    Dim ValorPagamento As String
    Dim CodCPF As String
    
    MsgBox "SELECIONE O ARQUIVO DE FOLHA DE PAGAMENTO."
    
    'Opção que permite escolher o aquivo que será aberto.
    With Application.FileDialog(msoFileDialogFilePicker)
        .Show
        Arquivo = .SelectedItems(1)
    End With
        
    'Especifica a partir de qual célula da Planilha deve começar a importação
    Dim rg As Range
    Set rg = Range("A5")
    
    'Abre o arquivo TXT
    Open Arquivo For Input As #1

    'Faz a Leitura do arquivo TXT Linha a Linha até o fim
    Do Until EOF(1)
       Line Input #1, Linha
       
       'Capturando as informações de cada Linha do arquivo TXT e insere nas variáveis

       Detalhe = Mid$(Linha, 14, 1)          '  14-14: Detalhe "A" ou "B"
       Favorecido = Mid$(Linha, 44, 30)      '  44-73: Nome do Favorecido
       ValorPagamento = Mid$(Linha, 120, 13) '120-134: Valor do Lançamento

      'Inserindo os Dados na Planilha

      
       If Detalhe = "A" Then 'Faz a leitura do Detalhe A
            rg.Offset(0, 0) = Favorecido
            rg.Offset(0, 2) = ValorPagamento
          
           Else
            If Detalhe = "B" Then              'Faz a leitura do Detalhe B
                Set rg = rg.Offset(-1, 0)      'volta uma linha acima na planilha e insere CPF
                CodCPF = Mid$(Linha, 19, 14)   '19-32: CPF
                rg.Offset(0, 1) = CodCPF       'insere CPF na 3º Coluna da Planilha. 1º Coluna é Zero.
            End If
        End If
        
        Set rg = rg.Offset(1, 0)               'próxima linha da planilha
    
    Loop

    Close #1

    'verifica qual é a última linha da planilha, será importado outro arquivo TXT na sequencia.
    Dim Ultimalinha As Long
    Ultimalinha = Range("A" & Rows.Count).End(xlUp).Row
    Ultimalinha = Ultimalinha + 1

End Sub

Public Sub ImportarWinthor()
     Dim ult_linha As Long
     Dim nome As String
     Dim cpf As String
     Dim valor As Double
     
     ult_linha = Planilha1.Cells(Planilha1.Rows.Count, 1).End(xlUp).Row
     
     For i = 7 To ult_linha
        nome = Planilha1.Cells(i, 1).Value
        cpf = Planilha1.Cells(i, 2).Value
        valor = Planilha1.Cells(i, 3).Value
        
        'Chamando função que faz insert do registro
        Call insertBd(nome, cpf, valor)
     Next i
     
 If Err <> 0 Then
    MsgBox "Erro ao tentar inserir registro no banco de dados: " & Err.Description, vbCritical
    Err.Clear
 End If
    
End Sub
