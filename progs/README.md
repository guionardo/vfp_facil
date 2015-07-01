# PRG´s com funções genéricas

## gs_api.prg
| Função | Descrição |
| ------------- | ----------- |
| GSAPIInit | Inicia declarações de APIs do windows |
| WinApiErrMsg | Obtém a mensagem de erro do sistema a partir do código retornado por GetLastError() |
| FileOpMessage | Executa função de operação de arquivo e retorna mensagem caso M.plFileOp_ShowError for .T. |
| ExcluirArquivo | Tenta Excluir arquivo |
| MoverArquivo | Tenta Mover Arquivo |
| CopiarArquivo | Tenta Copiar Arquivo |
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

## gs_config.prg
| Função | Descrição          |
| ------------- | ----------- |
| Config_Init | Inicializa as configurações |
| SetConfig | Grava uma chave de configuração |
| GetConfig | Lê uma chave de configuração |

## gs_dialogs.prg
| Função | Descrição |
| ------ | --------- |
| Erros | Mostra mensagens de Erros |
| Sim | Pergunta |
| Mensagem | Mostra mensagem |


## gs_funcoes.prg
| Função | Descrição          |
| ------------- | ----------- |
| STRZERO      | Análogo a STR, com preenchimento de zeros a esquerda.|
| DefaultTo     | Retorna o valor default de um parâmetro   |
| NetInfo | Retorna //COMPUTADOR/USUÁRIO/Sessão (CONSOLE ou RDP) |
| ArqVersao | Retorna a versão gravada nos detalhes do executável |

## gs_init.prg

Executa o SET PROCEDURE para adicionar os PRG da suite.
Inicializa as APIs

## gs_reportform.prg
| Função | Descrição |
| ------ | --------- |
| ReportForm | Imprime/visualiza report |
| GetImpressoraPadrao | Obtém a impressora padrão do report |
| SetImpressoraPadrao | Atualiza a impressora padrão do report |


## gs_tabelas.prg
| Função | Descrição |
| ------ | --------- |
| PushAlias | Grava Alias/Recno/Order do alias atual em uma lista FIFO |
| PopAlias | Restaura Alias/Recno/Order do último alias registrado por PushAlias |

## gs_waitcenter.prg
| Função | Descrição |
| ------ | --------- |
| WaitCenter | Wait window centralizado |

## gs_hwinfo.prg
| Função | Descrição |
| ------ | --------- |
| GetMotherboardInfo | Obtém informações da placa mãe |
| GetVolumeInfo	| Retorna objeto com informações da do volume no disco |
| GetDiskInfo | Retorna informações sobre o disco |
