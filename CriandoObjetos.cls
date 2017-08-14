VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CriandoObjetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Abre um arquivo no path selecionado
Public Sub OpenFile(ByVal path As String)
    Call Shell(path) 'Abre o arquivo do path
End Sub

'Retorna o nome do arquivo a partir do endereço
Public Function NomeDoArquivo(ByVal pathEndereco As String) As String
    Dim fso As Object
    Dim pasta As Scripting.folder
    Dim arquivo As Scripting.File
    '------------------------'
    Set fso = createObject("Scripting.FileSystemObject") 'New Scripting.FileSystemObject
    Set arquivo = fso.GetFile(pathEndereco) 'Set o arquivo do endereço
    NomeDoArquivo = arquivo.name 'Retorna o nome do arquivo
End Function

'Renomeia o arquivo
Public Sub RenomeiaArquivo(ByVal pathDoArquivo As String, _
                           ByVal j As Integer, _
                           ByRef arrayFiles As Variant)
    Dim fso As Object
    Dim pasta As Object
    Dim arquivo As Object
    Dim novoNome As String
    '--------------------------------'
    Set fso = createObject("Scripting.FileSystemObject") 'New Scripting.FileSystemObject
    Set arquivo = fso.GetFile(pathDoArquivo) 'Seta o endereço do folder aqui
    '--------------------------------'
    If InStr(1, arquivo.name, ".xlsx") > 0 Then 'Substitui a extenção de j & '.xlsx"
      novoNome = Replace(arquivo.name, ".xlsx", "")
      novoNome = novoNome & j & ".xlsx"
    ElseIf InStr(1, arquivo.name, ".xls") > 0 Then  'Se for xls
      novoNome = Replace(arquivo.name, ".xls", "")
      novoNome = novoNome & j & ".xls"
    End If
    arquivo.name = novoNome 'Renomeia o arquivo
End Sub

'Válida se o arquivo é Excel ou não
Public Function IsAccessBackEndString(ByVal fileName As String) As Boolean
    If Right(fileName, 5) = "accdb" Then 'Verifica os últimos caracteres
      IsAccessBackEndString = True
    End If
End Function

'Válida se o arquivo é Excel ou não
Public Function IsExcelString(ByVal fileName As String) As Boolean
    If Right(fileName, 3) = "xls" Or _ 'Verifica os últimos caracteres
        Right(fileName, 4) = "xlsx" Then
                IsExcelString = True
    End If
End Function

'Cria uma pasta no endereço do Sub
Public Sub CriarPasta(ByRef path As String)
  Dim fso As Object
  Set fso = createObject("Scripting.FileSystemObject") 'New Scripting.FileSystemObject
  fso.CreateFolder (path) 'Cria o folder aqui
End Sub

'Deletar a pasta
Public Sub DeletarPasta(ByVal pastaPath As String)
  Dim fso As Object
  Set fso = createObject("Scripting.FileSystemObject") 'New Scripting.FileSystemObject
  fso.DeleteFolder pastaPath
End Sub

'Deletar arquivo
Public Sub DeletarArquivo(ByVal arquivoPath As String)
On Error GoTo trata
  Dim fso As Object
  Set fso = createObject("Scripting.FileSystemObject") 'New Scripting.FileSystemObject
  fso.DeleteFile arquivoPath, True 'Antes de deletar o arquivo (FECHA A INSTANCIA DA MEMORIA antes de deletar)
Exit Sub
trata:
'Tratamento se o arquivo não pode ser excluído
    MsgBox "O arquivo: '" & arquivoPath & "' não pode ser excluído", vbCritical, "ERRO"
End Sub

'Itera pelos arquivos dentro da pasta
Public Sub IteraPelosArquivos(ByVal path As String)
  Dim fso As Object
  Set fso = createObject("Scripting.FileSystemObject") 'New Scripting.FileSystemObject
  Dim pasta As Scripting.folder
  Dim arquivo As Scripting.File
  '------------------'
  Set pasta = fso.GetFolder(path) 'Seta o endereço do folder aqui
  For Each arquivo In pasta.Files 'Itera pela lista de arquivos dentro da pasta
      Debug.Print arquivo.name
  Next arquivo
End Sub

'Itera pelos arquivos dentro da pasta
Public Sub IteraPelasSubPastas(ByVal path As String)
    Dim pasta As Object
    Dim fso As Object
    Set fso = createObject("Scripting.FileSystemObject") 'New Scripting.FileSystemObject
    Set pasta = fso.GetFolder(path)     'Seta o endereço do folder aqui
    '------------------------------------'
    For Each pasta In pasta.SubFolders 'Itera pela lista de arquivos dentro da pasta
      Debug.Print pasta.path
    Next pasta
End Sub

'Verifica se a pasta existe
Public Function PastaExiste(ByVal path As String) As Boolean
    Dim fso As Object
    Set fso = createObject("Scripting.FileSystemObject") 'New Scripting.FileSystemObject
    '----------------------------'
    If fso.FolderExists(path) = True Then 'Se a pasta existe, retorna true
      PastaExiste = True
    End If
End Function

'Verifica se é array ou não
Public Function EhArray(ByRef validaArray As Variant) As Boolean
On Error GoTo saiDosub
    Dim i As Variant
    For Each i In validaArray 'Itera pela lista de arrays
    Next i
    EhArray = True 'É array
Exit Function
saiDosub: 'Não é array
    EhArray = False
End Function

'Verifica se o arquivo existe
Public Function ArquivoExiste(ByVal path As String) As Boolean
    Dim fso As Object
    Set fso = createObject("Scripting.FileSystemObject") 'New Scripting.FileSystemObject
    '-------------------------'
    'Se a pasta existe, retorna true
    If fso.FileExists(path) = True Then
      ArquivoExiste = True
    End If
End Function

'Descobre o tamanho do arquivo
Public Function FileSize(ByVal filePath As String) As Double
  Dim arquivo As Object
    Dim fso As Object
    Set fso = createObject("Scripting.FileSystemObject") 'New Scripting.FileSystemObject
  Set arquivo = fso.GetFile(filePath) 'Seta a variável do arquivo
  FileSize = arquivo.Size 'Retorna o tamanho do arquivo
End Function

'Descobre o tamanho da pasta
Public Function FolderSize(ByVal folderPath As String) As Double
  Dim pasta As Object
    Dim fso As Object
    Set fso = createObject("Scripting.FileSystemObject") 'New Scripting.FileSystemObject
  Set pasta = fso.GetFolder(folderPath) 'Seta a variável do arquivo
  FolderSize = pasta.Size 'Retorna o tamanho do arquivo
End Function

