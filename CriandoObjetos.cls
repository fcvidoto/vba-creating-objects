VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CriandoObjetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Abre um arquivo no path selecionado
Public Sub OpenFile(ByVal path As String)
    
    'Abre o arquivo do path
    Call Shell(path)

End Sub

'Retorna o nome do arquivo a partir do endereço
Public Function NomeDoArquivo(ByVal pathEndereco As String) As String
    
    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    Dim pasta As Scripting.folder
    Dim arquivo As Scripting.File
    
    'Set o arquivo do endereço
    Set arquivo = fso.GetFile(pathEndereco)

    'Retorna o nome do arquivo
    NomeDoArquivo = arquivo.Name
    
End Function

'Seta o workbook na propriedade da classe
Public Function SetaWorkbook(ByVal pathEndereco As String, _
                             ByRef InstanciaExcel As InstanciaExcel) As Workbook
    
    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    Dim pasta As Scripting.folder
    Dim arquivo As Scripting.File
    Dim Workbook As Workbook
    Dim fileName As String
    
    'Set o arquivo do endereço
    Set arquivo = fso.GetFile(pathEndereco)
    
    'Nome do workbook
    fileName = arquivo.Name
    
    'Procura pelo workbook na instancia
    For Each Workbook In InstanciaExcel.xls.Workbooks
        
        'Se achar o workbook em algum workbook
        If Workbook.Name = fileName Then
            
            'Return
            Set SetaWorkbook = Workbook
            Exit Function
        End If
    Next Workbook

    'Retorna o nome do arquivo
    'NomeDoArquivo = arquivo.Name

End Function

'Renomeia o arquivo
Public Sub RenomeiaArquivo(ByVal pathDoArquivo As String, _
                           ByVal j As Integer, _
                           ByRef arrayFiles As Variant)

    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    Dim pasta As Scripting.folder
    Dim arquivo As Scripting.File
    Dim novoNome As String
    
    'Seta o endereço do folder aqui
    Set arquivo = fso.GetFile(pathDoArquivo)
        
    'Substitui a extenção de j & '.xlsx"
    If InStr(1, arquivo.Name, ".xlsx") > 0 Then
        novoNome = Replace(arquivo.Name, ".xlsx", "")
        novoNome = novoNome & j & ".xlsx"
    
    'Se for xls
    ElseIf InStr(1, arquivo.Name, ".xls") > 0 Then
        novoNome = Replace(arquivo.Name, ".xls", "")
        novoNome = novoNome & j & ".xls"
    End If
    
    'Renomeia o arquivo
    arquivo.Name = novoNome
    
End Sub

'Válida se o arquivo é Excel ou não
Public Function IsAccessBackEndString(ByVal fileName As String) As Boolean
    
    'Verifica os últimos caracteres
    If Right(fileName, 5) = "accdb" Then
        
        IsAccessBackEndString = True
    End If
    
End Function

'Válida se o arquivo é Excel ou não
Public Function IsExcelString(ByVal fileName As String) As Boolean
    
    'Verifica os últimos caracteres
    If Right(fileName, 3) = "xls" Or _
        Right(fileName, 4) = "xlsx" Then
        
        IsExcelString = True
    End If
    
End Function

'Cria uma pasta no endereço do Sub
Public Sub CriarPasta(ByRef path As String)
    
    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    
    'Cria o folder aqui
    fso.CreateFolder (path)
    
End Sub

'Deletar a pasta
Public Sub DeletarPasta(ByVal pastaPath As String)

    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    
    fso.DeleteFolder pastaPath
    
End Sub

'Deletar arquivo
Public Sub DeletarArquivo(ByVal arquivoPath As String)
    
On Error GoTo trata

    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    
    'Antes de deletar o arquivo (FECHA A INSTANCIA DA MEMORIA antes de deletar)
    fso.DeleteFile arquivoPath, True

Exit Sub

'Tratamento se o arquivo não pode ser excluído
trata:
    
    MsgBox "O arquivo: '" & arquivoPath & "' não pode ser excluído", vbCritical, "ERRO"
    
End Sub

'Itera pelos arquivos dentro da pasta
Public Sub IteraPelosArquivos(ByVal path As String)
    
    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    Dim pasta As Scripting.folder
    Dim arquivo As Scripting.File
    
    'Seta o endereço do folder aqui
    Set pasta = fso.GetFolder(path)
    
    'Itera pela lista de arquivos dentro da pasta
    For Each arquivo In pasta.Files
        Debug.Print arquivo.Name
    Next arquivo

End Sub

'Itera pelos arquivos dentro da pasta
Public Sub IteraPelasSubPastas(ByVal path As String)
    
    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    Dim pasta As Scripting.folder
    
    'Seta o endereço do folder aqui
    Set pasta = fso.GetFolder(path)
    
    'Itera pela lista de arquivos dentro da pasta
    For Each pasta In pasta.SubFolders
        Debug.Print pasta.path
    Next pasta

End Sub

'Verifica se a pasta existe
Public Function PastaExiste(ByVal path As String) As Boolean

    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    
    'Se a pasta existe, retorna true
    If fso.FolderExists(path) = True Then
        PastaExiste = True
    End If
    
End Function

'Verifica se é array ou não
Public Function EhArray(ByRef validaArray As Variant) As Boolean
    On Error GoTo saiDosub

    Dim i As Variant
    
    'Itera pela lista de arrays
    For Each i In validaArray
    Next i

    'É array
    EhArray = True

Exit Function

'Não é array
saiDosub:
    EhArray = False
End Function


'Verifica se o arquivo existe
Public Function ArquivoExiste(ByVal path As String) As Boolean

    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    
    'Se a pasta existe, retorna true
    If fso.FileExists(path) = True Then
        ArquivoExiste = True
    End If
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub MEDELETE()

    Dim path As String: path = "C:\Users\Fernando\Documents\Teste\"
    Dim i As Integer
    Dim CriandoObjetos As CriandoObjetos: Set CriandoObjetos = New CriandoObjetos
    Dim arrayDePastas(1 To 3) As String
    
    'se a pasta não existe, cria a pasta path
    If CriandoObjetos.PastaExiste(path) = False Then
        CriandoObjetos.CriarPasta (path)
    End If
    
    'Estrutura das pastas
    arrayDePastas(1) = path & "Fernando"
    arrayDePastas(2) = path & "César"
    arrayDePastas(3) = path & "Vidoto"
    
    'Dentro da pasta Teste, criar a estrutura de pastas
    For i = 1 To UBound(arrayDePastas)
        
        'Se a pasta na estrutura de pastas não existe, criar ela no folder
        If CriandoObjetos.PastaExiste(arrayDePastas(i)) = False Then
            CriandoObjetos.CriarPasta (arrayDePastas(i))
        End If
        
    Next i
    
End Sub

'Cria a tabela dinamica com argumentos
Public Sub CriaTabelaDinamica(ByVal sheetName As String, _
                              ByVal nomeDaPivot As String, _
                              ByVal linhaInicial As Integer, _
                              ByVal colunaInicial As Integer, _
                              ByVal sourceDaPivotTable As Range, _
                              ByVal colunaInicialDeLinha As String, _
                              ByVal colunaInicialValor As String, _
                              ByRef wbook As Workbook)
    
    Dim destinoDaPivot As String: destinoDaPivot = sheetName & "!R" & linhaInicial & "C" & colunaInicial
    
    'Cria a tabela dinamica apenas com argumentos
    wbook.PivotCaches.Create(SourceType:=xlDatabase, _
                                          SourceData:=sourceDaPivotTable, _
                                          Version:=xlPivotTableVersion14). _
                                          CreatePivotTable TableDestination:=destinoDaPivot, _
                                          TableName:=nomeDaPivot, _
                                          DefaultVersion:=xlPivotTableVersion14
                                          
                                          
    '******* Define os filtros na pivot recém criada
    With wbook.Sheets(sheetName).PivotTables(nomeDaPivot)
            
        'Coluna de produto
        .PivotFields(colunaInicialDeLinha).Orientation = xlRowField
        .PivotFields(colunaInicialDeLinha).Position = 1
        
        'Valor agregado
        .AddDataField wbook.Sheets(sheetName).PivotTables(nomeDaPivot).PivotFields(colunaInicialValor), "Contagem de " & colunaInicialValor, xlCount
    
    End With
    
End Sub

'Cria uma nova sheet ao workbook
Public Sub CriarSheet(ByVal nomeDaSheet As String, _
                      ByRef workbookName As Workbook)

    'Adiciona a sheet
    workbookName.Sheets.Add
    workbookName.Sheets(1).Name = nomeDaSheet

End Sub

'Descobre o tamanho do arquivo
Public Function FileSize(ByVal filePath As String) As Double

    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    Dim arquivo As Scripting.File
    
    'Seta a variável do arquivo
    Set arquivo = fso.GetFile(filePath)
    
    'Retorna o tamanho do arquivo
    FileSize = arquivo.Size
    
End Function

'Descobre o tamanho da pasta
Public Function FolderSize(ByVal folderPath As String) As Double

    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    Dim pasta As Scripting.folder
    
    'Seta a variável do arquivo
    Set pasta = fso.GetFolder(folderPath)
    
    'Retorna o tamanho do arquivo
    FolderSize = pasta.Size
    
End Function

'Salva o arquivo no local desejado
Public Sub SalvaArquivo(ByVal pathDoFolder As String, _
                        ByRef sheetAclonar As String)
On Error GoTo exitSub
    'Disabilita msgs
    'Application.DisplayAlerts = False
    
    'Cria o novo workbook nesse método
    Sheets(sheetAclonar).Copy
    
    'Salva o workbook no local especídifico
    ActiveWorkbook.SaveAs pathDoFolder
    ActiveWorkbook.Close False
    
    'Reabilida msg
    'Application.DisplayAlerts = True
    
    Sheets("Consolidar_Contestacoes").Range("G13").Value = pathDoFolder
    MsgBox "Sucesso! Salvo no local '" & pathDoFolder & "'", vbInformation, "Sucesso"

Exit Sub

exitSub:
    
    'Fecha o workbook sem salvar
    ActiveWorkbook.Close False

End Sub

'Desoculta a sheet específica
Public Sub OcultaSheet(ByVal sheetName As String)

    Dim sheet As Worksheet
    
    'Itera por todas as sheets do workbook
    For Each sheet In ActiveWorkbook.Worksheets
        
        'Se o nome da sheet for o escolhido, OCULTA
        If sheet.Name = sheetName Then
            sheet.Visible = xlSheetHidden
            Exit Sub
        End If
    Next sheet

End Sub

'Oculta a sheet específica
Public Sub DesOcultaSheet(ByVal sheetName As String)

    Dim sheet As Worksheet
    
    'Itera por todas as sheets do workbook
    For Each sheet In ActiveWorkbook.Worksheets
        
        'Se o nome da sheet for o escolhido, OCULTA
        If sheet.Name = sheetName Then
            sheet.Visible = xlSheetVisible
            Exit Sub
        End If
    Next sheet

End Sub

'Fecha o workbook se estiver aberto
Public Sub FechaWorkbook(ByVal nomeDoWorkbookComExtensao As String)

    Dim ArquivosAbertos As Workbook
        
    Dim xls As Excel.Application: Set xls = GetObject(, "Excel.Application")
        
    'Itera por todos os workbooks abertos
    For Each ArquivosAbertos In xls.Workbooks
        
        'Fecha o workbook específico E não salva as alterações
        If ArquivosAbertos.Name = nomeDoWorkbookComExtensao Or _
           ArquivosAbertos.Name = "formulario.xlsx" Then
            ArquivosAbertos.Close False
            Exit Sub
        End If
    Next ArquivosAbertos

End Sub

'Desprotege a planilha
Sub DesprotegerPlanilha(ByRef sheet As Worksheet)

Dim i, j, k, l, m, i1, i2, i3, i4, i5, i6, n As Integer
On Error Resume Next
    
    For i = 65 To 66
        For j = 65 To 66
            For k = 65 To 66
                For l = 65 To 66
                    For m = 65 To 66
                        For i1 = 65 To 66
                            For i2 = 65 To 66
                                For i3 = 65 To 66
                                    For i4 = 65 To 66
                                        For i5 = 65 To 66
                                            For i6 = 65 To 66
                                                For n = 32 To 126
                                                    sheet.Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                                                    
                                                    If ActiveSheet.ProtectContents = False Then
                                                        'MsgBox "A planilha foi desprotegida com sucesso!!!"
                                                        Exit Sub
                                                    End If
                                                        
                                                Next
                                            Next
                                        Next
                                    Next
                                Next
                            Next
                        Next
                    Next
                Next
            Next
        Next
    Next
    
End Sub



