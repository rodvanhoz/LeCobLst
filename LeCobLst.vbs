' Le o arquivo de chaves dos sistemas por prefeitura
'
' retorno:
'	0 -> Sem erros
'	1 -> Erro de processamento
'
' Data: 14/07/2019
' Author: Rodrigo Vanhoz Ribeiro
'
' Versão: 1
'
' Alteracoes informar abaixo
'

Option Explicit

' checagem de parametros
If WScript.Arguments.Count <> 2 And WScript.Arguments.Count <> 3 Then
	WScript.Echo "uso: [Caminho arquivo] [Nome Prefeitura]"
	WScript.Echo ""
	WScript.Quit( 1 )
End if

' Constantes para uso de manipulacao de arquivos
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
 
' parametros recebidos
dim nomearq, nomeprefeitura
nomearq = WScript.Arguments.Unnamed(0)
nomeprefeitura = WScript.Arguments.Unnamed(1)


' variaveis de arquivo
dim fs, arq, linhas, totallinhas, linha, cont, contstru, contachados, nrolinha
Set fs  = CreateObject( "scripting.filesystemobject" )
Set arq = fs.OpenTextFile( nomearq, ForReading, TristateFalse )

' variaveis da estrutura do arquivo
Dim codempresa, nomeempresa, cnpj, ddd, fone
Dim codparceiro, nomeparceiro, datavalidade, chave
Dim datageracao
Dim email

contachados = 0

' carregando arquivo
linhas = Split(arq.ReadAll, Chr(13) & Chr(10))
totalLinhas = arq.Line
arq.Close

' percorrendo arquivo
For cont = 0 to (UBound(linhas) - 1)
	linha = linhas(cont)
	
	If cont >= 8 Then
		If IsNumeric(Mid(linha, 1, 5)) Then
			codempresa = Mid(linha, 1, 5)
			nomeempresa = Mid(linha, 6, 42)
			ddd = Mid(linha, 66, 5)
			fone = Mid(linha, 71, 10)
			
			nrolinha = cont + 1
			contstru = 1
		
		ElseIf contstru = 1 And Not IsNumeric(Mid(linha, 1, 5)) Then
			codparceiro = Mid(linha, 6, 5)
			nomeparceiro = Mid(linha, 11, 36)
			datavalidade = Mid(linha, 47, 10)
			chave = Mid(linha, 57, 24)
			
			contstru = 2
		
		ElseIf contstru = 2 And Not IsNumeric(Mid(linha, 1, 5)) Then
			datageracao = Mid(linha, 6, 10)
			
			contstru = 3
		
		ElseIf contstru = 3 And Not IsNumeric(Mid(linha, 1, 5)) Then
			email = Mid(linha, 14, 67)
			
			contstru = 4
		
		Elseif contstru = 4 And Not IsNumeric(Mid(linha, 1, 5)) Then
			If InStr(UCase(nomeempresa), UCase(nomeprefeitura)) Then
				contachados = contachados + 1
				WScript.Echo "Linha Arquivo..: " & CStr(nrolinha)
				WScript.Echo "Nome Empresa...: " & Trim(codempresa) & "-" & Trim(nomeempresa)
				WScript.Echo "Chave Ativacao.: " & Trim(chave)
				WScript.Echo "Data Validade..: " & Trim(datavalidade)
				WScript.Echo ""
				
				contstru = 0
			End If
		End if
	End If
Next

WScript.Echo "---------------------"
WScript.Echo "Total Achados: " & CStr(contachados)
WScript.Echo ""

arq.Close

WScript.Quit(0)