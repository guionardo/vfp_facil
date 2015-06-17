# PRG´s com funções genéricas

## gs_funcoes.prg
| Função | Descrição          |
| ------------- | ----------- |
| STRZERO      | Análogo a STR, com preenchimento de zeros a esquerda.|
| DefaultTo     | Retorna o valor default de um parâmetro   |
| NetInfo | Retorna //COMPUTADOR/USUÁRIO/Sessão (CONSOLE ou RDP) |
| ArqVersao | Retorna a versão gravada nos detalhes do executável |

## gs_api.org
| Função | Descrição |
| ------------- | ----------- |
| GSAPIInit | Inicia declarações de APIs do windows |
| AProcessos | Obtém lista de processos rodando sob o nome lcProcesso |
| CreateShortCut | Cria atalho do windows para o aplicativo |
| GetFileSizeAPI | Retorna o tamanho de um arquivo usando a API do Windows | 
| GetWorkGroup | Retorna o WORKGROUP ou Domínio da máquina local |
| IsUserAdmin | Retorna se o usuário logado no windows é administrador |
| GetWinVer	| Retorna a versão do windows |
| Createprocess | Cria um processo do windows para executar uma aplicação |
| GetStartupInfo | Usada em Createprocesse, cria uma estrutura STARTUP |
| ReduceMemory | Reduz o consumo de memória |
| MoveArquivo | Tenta mover ou copiar arquivo(s) |
| ExcluiArquivo	| Tenta excluir um arquivo |
| FileExtract |	Extrai arquivos incorporados no executável |
| FileOpWithProgressBar | Copia/Move arquivo usando o shell |