********************************************************************************
* Controle de Metadata de Tabelas
*
* Formato do METADATA
* #DBF:ALIAS
* CAMPO TIPO(TAMANHO)
* INDICE:EXPRESSAO
*
DEFINE CLASS MetaData AS COLLECTION
	Arquivo 	= ""
	Progresso 	= .T.
	BASE		= ""
	LogFunc		= "GSAddLog"	&& Função de log
	MsgFunc		= "Erros"		&& Função de mensagem
	Mensagem	= ""			&& Mensagem

	FUNCTION INIT
		LPARAMETERS lcArq
		IF TYPE("M.PCBase")="U"
			PUBLIC PCBase
		ENDIF
		IF VARTYPE(M.PCBase)#"C" OR !DIRECTORY(m.PCBase)
			m.PCBase = ADDBS(JUSTPATH(APPLICATION.SERVERNAME))+"DB"
			IF !DIRECTORY(m.PCBase)
				MKDIR (m.PCBase)
			ENDIF
		ENDIF
		THIS.BASE = m.PCBase

		THIS.Arquivo = m.lcArq
	ENDFUNC

	FUNCTION Arquivo_Assign
		LPARAMETERS lcArq
		IF (VARTYPE(m.lcArq)#"C") OR !FILE(m.lcArq)
			RETURN .T.
		ENDIF
		THIS.Arquivo = m.lcArq

		LOCAL lnMDC, lcMDT, lnI, lcMD
		m.lcMDT = FILETOSTR(m.lcArq)
		m.lnMDC = GETWORDCOUNT(m.lcMDT,'!')

		IF THIS.Progresso
			GSProgressFormInit(m.lnMDC)
		ENDIF

		LOCAL loRet
		LOCAL ARRAY laMD(5)
		FOR m.lnI = 1 TO m.lnMDC
			m.lcMD = GETWORDNUM(m.lcMDT,m.lnI,"!")
			m.loRet = THIS.MetaDataParser(m.lcMD)
			IF !m.loRet.OK
				LOOP
			ENDIF
			IF THIS.Progresso
				GSProgressNovoProcesso(m.loRet.DBF+":"+m.loRet.ALIAS)
				GSProgressProcessa(100)
			ENDIF
			THIS.ADD(m.loRet,m.loRet.ALIAS)
			IF !THIS.CheckStruct(m.loRet.DBF,m.loRet.ALIAS,m.loRet.FIELDS,m.loRet.INDEXES)
				Erros(THIS.Mensagem,"Metadata Inválido!")
				RETURN .F.
			ENDIF
		NEXT
		IF THIS.Progresso
			GSProgressFormClose()
		ENDIF
		RETURN .T.

	ENDFUNC

	FUNCTION MetaDataParser
		LPARAMETERS lcMetadata
		LOCAL lnI, lnL, lnP,lnMW, lcL, lcDBF, lcAlias, lcFields, lcIndexes, lcErro, lcRemoveCampos
		m.lnMW = SET("Memowidth")
		SET MEMOWIDTH TO 128
		m.lnL = MEMLINES(m.lcMetadata)
		STORE "" TO m.lcDBF, m.lcAlias, m.lcFields, m.lcIndexes, lcErro, lcRemoveCampos
		FOR m.lnI = 1 TO m.lnL
			m.lcL = ALLTRIM(MLINE(m.lcMetadata,m.lnI))
			IF EMPTY(m.lcL)
				LOOP
			ENDIF
* Verifica comentário e remove do campo
			m.lnP = AT('&',m.lcL)
			IF m.lnP>0
				m.lcL = ALLTRIM(LEFT(m.lcL,m.lnP-1))
			ENDIF
			DO CASE
* Identifica DBF e Alias
				CASE (LEFT(m.lcL,1)=="#") AND (":"$m.lcL)
					m.lcL = SUBSTR(m.lcL,2)
					m.lcDBF = GETWORDNUM(m.lcL,1,':')
					m.lcAlias = GETWORDNUM(m.lcL,2,':')
* Identifica Campo
				CASE (" "$m.lcL) AND ("("$m.lcL) AND (")"$m.lcL)
					IF LEFT(m.lcL,1)=="-"	&& Campos a remover
						m.lcRemoveCampos = m.lcRemoveCampos + SUBSTR(m.lcL,2)+';'
					ELSE
						m.lcFields = m.lcFields + m.lcL + ';'
					ENDIF
* Identifica Índice
				CASE ":"$m.lcL
					m.lcIndexes = m.lcIndexes + m.lcL+";"
				OTHERWISE
					m.lcErro = m.lcL+"\n"
			ENDCASE
		NEXT
		SET MEMOWIDTH TO (m.lnMW)
		m.lcFields = LEFT(m.lcFields,LEN(m.lcFields)-1)
		m.lcIndexes = LEFT(m.lcIndexes,LEN(m.lcIndexes)-1)
		LOCAL loRet
		m.loRet = CREATEOBJECT("EMPTY")
		ADDPROPERTY(m.loRet,"DBF",m.lcDBF)
		ADDPROPERTY(m.loRet,"ALIAS",m.lcAlias)
		ADDPROPERTY(m.loRet,"FIELDS",m.lcFields)
		ADDPROPERTY(m.loRet,"INDEXES",m.lcIndexes)
		ADDPROPERTY(m.loRet,"ERRO",m.lcErro)
		ADDPROPERTY(m.loRet,"REMOVECAMPOS",m.lcRemoveCampos)
		ADDPROPERTY(m.loRet,"OK",EMPTY(m.lcErro))
		RETURN m.loRet
	ENDFUNC


*
***********************************************************
* TABGetStructChange
* Obtém o código SQL de alteração da estrutura da tabela
*
	FUNCTION StructChange
		LPARAMETERS cTable,cAlias, cStruct
		LOCAL cSQL, lEstavaAberta, nF, cFN, cFT, nI, oExc AS EXCEPTION, lErro, lExclusiva
*
* Se a tabela não estiver aberta, abre em modo compartilhado
*
		m.lEstavaAberta = USED(m.cAlias)
		m.lExclusiva = m.lEstavaAberta AND ISEXCLUSIVE(m.cAlias)
		IF USED(m.cAlias)
			m.cTable = DBF(m.cAlias)
		ELSE
			m.cTable = FORCEEXT(THIS.FULLPATH(m.cTable),"DBF")
			IF !FILE(m.cTable)
				m.cSQL = 'CREATE TABLE "'+m.cTable+'" ('+STRTRAN(m.cStruct,';',',')+')'
				RETURN m.cSQL
			ENDIF
		ENDIF
		m.lErro = .F.
		TRY
			IF ! m.lEstavaAberta
				USE (m.cTable) IN 0 ALIAS (m.cAlias) SHARED
			ENDIF
		CATCH TO oExc && Captura exceções ignorando a falta do indice
			IF !USED(m.cAlias)
				m.lErro = .T.
			ENDIF
		ENDTRY
		IF m.lErro
			RETURN ""
		ENDIF


*
* Verifica a estrutura da tabela com o metadata informado
*
		m.nF = GETWORDCOUNT(m.cStruct,';')
		m.cSQL = ''
		FOR m.nI=1 TO m.nF
			m.cFN = GETWORDNUM(GETWORDNUM(m.cStruct,m.nI,';'),1) && Nome do campo
			m.cFT = GETWORDNUM(GETWORDNUM(m.cStruct,m.nI,';'),2) && Tipo do campo
			IF !THIS.IsField(m.cFN,m.cAlias) && Se o campo não existir
				m.cSQL = m.cSQL+' ADD '+m.cFN+' '+m.cFT+' '
			ELSE
				IF !THIS.IsFieldType(m.cFN,m.cAlias,m.cFT)
					m.cSQL = m.cSQL+' ALTER '+m.cFN+' '+m.cFT+' '
				ENDIF
			ENDIF
		ENDFOR

		IF m.lEstavaAberta
			IF m.lExclusiva
				USE IN (m.cAlias)
				USE (m.cTable) IN 0 ALIAS (m.cAlias) EXCLUSIVE
			ENDIF
		ELSE
			USE IN (m.cAlias)
		ENDIF

		IF !EMPTY(m.cSQL)
			m.cSQL = 'ALTER TABLE '+m.cAlias+' '+m.cSQL
		ENDIF
		RETURN m.cSQL
	ENDFUNC

****
* IsField
* Retorna se o campo do alias existe
*****
	FUNCTION IsField
		LPARAMETERS lcField AS STRING, lcAlias AS STRING
		RETURN !EMPTY(FIELD(m.lcField,m.lcAlias))
	ENDFUNC

****
* TABIsFieldType
* Retorna se o campo do alias é do tipo informado
*****
	FUNCTION IsFieldType
		LPARAMETERS cField AS STRING, cAlias AS STRING, cType AS STRING
		m.cType = UPPER(STRTRAN(m.cType,' ',''))
		RETURN m.cType==THIS.FieldType(m.cField,m.cAlias)
	ENDFUNC

****
* TABFieldType
* Retorna o tipo do campo no formato de criação do SQL
* Por exemplo: C(10), T(8), M(4), N(12,2), etc
*****
	FUNCTION FieldType
		LPARAMETERS lcField AS STRING, lcAlias AS STRING
		LOCAL ARRAY laFields(1)

		IF !THIS.IsField(m.lcField,m.lcAlias)
			RETURN 'U'
		ENDIF

		LOCAL lnFieldCount, lnFieldIndex, lnC
*
* Localiza informações do campo na base
*
		m.lnFieldCount=AFIELDS(m.laFields,m.lcAlias)
		m.lnC = 1
		m.lnFieldIndex = 0
		DO WHILE (m.lnC <=m.lnFieldCount) AND (m.lnFieldIndex=0)
			IF (UPPER(ALLTRIM(m.laFields(m.lnC ,1)))=UPPER(ALLTRIM(m.lcField)))
				m.lnFieldIndex=m.lnC
			ELSE
				m.lnC =m.lnC + 1
			ENDIF
		ENDDO
		IF m.lnFieldIndex=0
* Caso não encontre um campo retorna "U"
			RETURN "U"
		ENDIF

		LOCAL lcRet,lcX
* Identifica tipo do campo
		m.lcRet = m.laFields(m.lnFieldIndex,2)	&& Tipo do campo
		m.lcX = ''
* Tamanho
		IF m.laFields(m.lnFieldIndex,3)>0
			m.lcX = ALLTRIM(STR(m.laFields(m.lnFieldIndex,3),3,0))
		ENDIF
* Decimais
		IF m.laFields(m.lnFieldIndex,4)>0
			m.lcX = m.lcX+','+ALLTRIM(STR(m.laFields(m.lnFieldIndex,4),2,0))
		ENDIF
		IF !EMPTY(m.lcX)
			m.lcX='('+m.lcX+')'
		ENDIF
		RETURN m.lcRet+m.lcX
	ENDFUNC

****
*
* Retorna o caminho completo do DBF
*
***
	FUNCTION FULLPATH
		LPARAMETERS lcDBF
		RETURN FORCEPATH(m.lcDBF,THIS.BASE)
	ENDFUNC

****
*
* Tenta abrir tabela
* DBF, Alias, Exclusivo, Retorno
****
	FUNCTION TryUse
		LPARAMETERS lcDBF, lcAlias, llExclusive, lcRet
		IF !FILE(m.lcDBF)
			RETURN .F.
		ENDIF
		IF USED(m.lcAlias)
			IF ISEXCLUSIVE(m.lcAlias) OR !m.llExclusive
&& Se a tabela já estiver aberta exclusivamente ou não é necessário abrí-la exclusivamente, sai
				RETURN .T.
			ENDIF
			USE IN (m.lcAlias)
		ENDIF
		LOCAL lcModo, llErro
		m.lcModo = IIF(m.llExclusive,"EXCL","SHARED")
		TRY
			IF VARTYPE(m.lcRet)#"C"
				m.lcRet = ""
			ENDIF
			USE (m.lcDBF) IN 0 ALIAS (m.lcAlias) &lcModo
		CATCH TO oExc
			IF oExc.ERRORNO != 1707 && Ignora se o erro for a falta do CDX
				m.lcRet = m.lcRet + TRANSFORM(oExc.ERRORNO)+":"+oExc.MESSAGE
			ENDIF
		ENDTRY
		IF !USED(m.lcAlias)	&& Se houve um erro anteriormente, tenta novamente
			TRY
				USE (m.lcDBF) IN 0 ALIAS (m.lcAlias) &lcModo
			CATCH
			ENDTRY
		ENDIF
		RETURN USED(m.lcAlias)
	ENDFUNC

****
*
* Log
*
****
	PROCEDURE LOG
		LPARAMETERS lcMsg
		IF VARTYPE(THIS.LogFunc)#"C"
			RETURN
		ENDIF
		LOCAL lcL
		m.lcL = THIS.LogFunc+"(m.lcMsg)"
		&lcL
	ENDPROC

****
*
* Msg
*
****
	PROCEDURE Msg
		LPARAMETERS lcMsg, llErro
		THIS.Mensagem = IIF(m.llErro,'ERRO: ','')+m.lcMsg
		IF VARTYPE(THIS.MsgFunc)#"C"
			RETURN
		ENDIF
		LOCAL lcL
		m.lcL = THIS.MsgFunc+"(m.lcMsg,m.llErro)"
		&lcL
	ENDPROC

*
************************************************************
* ARQExclusivo
* Retorna se o arquivo está disponível para acesso exclusivo
* lcArquivo	-> Nome do arquivo
* Retorna boolean
	FUNCTION ARQExclusivo
		LPARAMETERS lcArquivo, llPrecisaExistir
		IF VARTYPE(m.llPrecisaExistir)#"L"
			m.llPrecisaExistir = .T.
		ENDIF
		IF !FILE(m.lcArquivo)
			RETURN !m.llPrecisaExistir
		ENDIF
		LOCAL lH
		m.lH = FOPEN(m.lcArquivo,12)
		FCLOSE(m.lH)
		RETURN m.lH>-1
	ENDFUNC

****
* CheckStruct
*
****
	FUNCTION CheckStruct
		LPARAMETERS cTable AS STRING, cAlias AS STRING, cStruct AS STRING, cIndex AS STRING, lFechar AS Boolean
		LOCAL cSQL AS STRING,nI AS INTEGER, nF AS INTEGER, nFN AS STRING, cFT AS STRING,;
			cFIndex AS STRING, cOnError AS STRING, oExc AS EXCEPTION, cRet, lErro, cDBF, cCDX,;
			lExclusivo
		m.cDBF = THIS.FULLPATH(FORCEEXT(m.cTable,"DBF"))
		m.cCDX = FORCEEXT(m.cDBF,'CDX')

		m.cSQL = THIS.StructChange(m.cTable,m.cAlias,m.cStruct)
		m.lErro = .F.
*
* cSQL conterá as alterações para criar/alterar a tabela, ou vazio se não for necessário
*
		m.cRet = ''
		IF VARTYPE(m.lFechar)#"L"
			m.lFechar = .F.
		ENDIF

		IF EMPTY(m.cSQL)
			IF !THIS.TryUse(m.cDBF,m.cAlias,.F.,@cRet)
				THIS.Msg("Não foi possível abrir a tabela "+m.cAlias+'\n'+;
					"Erro:"+m.cRet,.T.)
				THIS.LOG("Metadata.CheckStruct: USE "+m.cDBF+" = "+m.cRet)
				RETURN .F.
			ENDIF

		ELSE
			m.cRet = 'Metadata.CheckStruct '+m.cSQL
			IF "CREATE TABLE"$m.cSQL
				TRY
					&cSQL
					USE
				CATCH TO oExc
					m.cRet = m.cRet +': ('+TRANSFORM(m.oExc.ERRORNO)+') '+m.oExc.MESSAGE
					m.lErro = .T.
				ENDTRY
				IF m.lErro
					THIS.Msg(m.cRet,.T.)
					THIS.LOG(m.cRet)
					RETURN .F.
				ENDIF
				THIS.TryUse(m.cDBF,m.cAlias)
&& Chegando aqui, criou a tabela com sucesso
			ELSE
				IF (!USED(m.cAlias)) AND !THIS.TryUse(m.cDBF, m.cAlias, .T., @cRet)
					THIS.Msg("Não foi possível abrir a tabela "+m.cAlias)
					THIS.LOG("Metadata.CheckStruct: USE "+m.cDBF+" = "+m.cRet)
					RETURN .F.
				ENDIF
&& Se for alterar uma tabela deve estar aberta exclusivamente
				m.lExclusivo = ISEXCLUSIVE(m.cAlias)
				IF !m.lExclusivo && Se não estiver aberto exclusivamente, fecha e tenta abrir novamente
					m.cDBF = DBF(m.cAlias)
					USE IN (m.cAlias)
					IF !THIS.ARQExclusivo(m.cDBF)
						THIS.LOG('Tabela '+m.cDBF+' USE EXCLUSIVE FALHOU!')
						THIS.Msg('É necessário abrir a tabela '+m.cAlias+' exclusivamente, mas existem outros usuários a acessando.\n'+;
							'Execute o sistema em modo EXCLUSIVO para garantir o acesso às tabelas.',.T.)
						RETURN .F.
					ENDIF
					m.cCDX = FORCEEXT(m.cDBF,'CDX')
					IF FILE(m.cCDX)
						IF !THIS.ARQExclusivo(m.cCDX)
							THIS.LOG('INDEX '+m.cCDX+' ACESSO EXCLUSIVO FALHOU!')
							THIS.Msg('É necessário reconstruir os índices da tabela '+m.cAlias+'\n'+;
								'mas o arquivo '+m.cCDX+' não está disponível exclusivamente!\n'+;
								'Execute o sistema em modo EXCLUSIVO para garantir o acesso às tabelas.',.T.)
							RETURN .F.
						ENDIF
						ERASE (m.cCDX) && Exclui o arquivo de índice para forçar a reindexação
					ENDIF

					m.llErro = !THIS.TryUse(m.cDBF,m.cAlias,.T.,@cRet)
					THIS.LOG(m.cRet)
					IF m.llErro
						THIS.Msg(m.cRet,.T.)
						RETURN .F.
					ENDIF
&& Chegando aqui, a tabela está aberta exclusivamente
				ENDIF
				TRY
					THIS.LOG('Metadata.CheckStruct: '+m.cSQL)
					&cSQL
				CATCH TO oExc
					m.cRet = m.cRet +': ('+TRANSFORM(m.oExc.ERRORNO)+') '+m.oExc.MESSAGE
					m.lErro = .T.
				ENDTRY
				IF !m.lExclusivo
					IF USED(m.cAlias)
						USE IN (m.cAlias)
					ENDIF
					THIS.TryUse(m.cDBF,m.cAlias)
				ENDIF
				IF m.lErro
					THIS.LOG(m.cRet)
					THIS.Msg(m.cRet,.T.)
					RETURN .F.
				ENDIF
			ENDIF

		ENDIF
		IF EMPTY(m.cIndex)
			RETURN .T.
		ENDIF
*
* Verificando os índices
*
		m.cCDX = FORCEEXT(m.cDBF,'CDX')
		m.nF = GETWORDCOUNT(m.cIndex,';')
		m.cSQL = ''

		IF TYPE("M.Indexar")="U"
			LOCAL Indexar
			m.Indexar = .F.
		ENDIF

		IF (M.Indexar OR !FILE(m.cCDX))
			m.lExclusivo = ISEXCLUSIVE(m.cAlias)
			USE IN (m.cAlias)
			IF !THIS.ARQExclusivo(m.cCDX,.F.)
				THIS.LOG('INDEX '+m.cCDX+' ACESSO EXCLUSIVO FALHOU!')
				THIS.Msg('00-É necessário reconstruir os índices da tabela '+m.cAlias+'\n'+;
					'mas o arquivo '+m.cCDX+' não está disponível exclusivamente!\n'+;
					'Execute o sistema em modo EXCLUSIVO para garantir o acesso às tabelas.',.T.)
				RETURN .F.
			ENDIF
			IF FILE(m.cCDX)
				ERASE (m.cCDX)
			ENDIF

			IF !THIS.ARQExclusivo(m.cDBF)
				THIS.LOG('Tabela '+m.cDBF+' USE EXCLUSIVE FALHOU!')
				THIS.Msg('01-É necessário abrir a tabela '+m.cAlias+' exclusivamente, mas existem outros usuários a acessando.\n'+;
					'Execute o sistema em modo EXCLUSIVO para garantir o acesso às tabelas.',.T.)
				RETURN .F.
			ENDIF
			m.cRet = 'TABCheckStruct '

			IF !THIS.TryUse(m.cDBF,m.cAlias,.T.,@cRet)
				THIS.LOG(m.cRet)
				THIS.Msg(m.cRet,.T.)
				RETURN .F.
			ENDIF

			SELECT (m.cAlias)
			m.nF = GETWORDCOUNT(m.cIndex,';')
			FOR m.nI=1 TO m.nF
				m.cFN = GETWORDNUM(GETWORDNUM(m.cIndex,m.nI,';'),1,':') && Tag do indice
				m.cFT = GETWORDNUM(GETWORDNUM(m.cIndex,m.nI,';'),2,':') && Expressão
				IF EMPTY(m.cFN) OR EMPTY(m.cFT)
					THIS.LOG('Indice de '+m.cTable+" ("+m.cAlias+") ERRO DE SINTAXE: TAG = ["+m.cFN+"] EXPRESSÃO = ["+m.cFT+"]")
					THIS.Msg('ERRO DE SINTAXE DE ÍNDICE: Contate o suporte!',.T.)
					LOOP
				ENDIF
				WAIT WINDOW "Indexando "+m.cTable+' ('+m.cAlias+') = '+m.cFN NOWAIT
				THIS.LOG("Indexando "+m.cTable+' ('+m.cAlias+') = '+m.cFN)
				TRY
					INDEX ON &cFT TAG &cFN
				CATCH TO oExc
					THIS.LOG("INDEX ON "+m.cFT+" TAG "+m.cFN+" -> "+oExc.MESSAGE)
					THIS.Msg("Não foi possível indexar a tabela "+m.cTable+' ('+m.cAlias+')\n'+;
						'Erro: '+oExc.MESSAGE,.T.)
				ENDTRY
			ENDFOR
			IF !M.lExclusivo
				USE IN (m.cAlias)
				USE (m.cDBF) IN 0 ALIAS (m.cAlias) SHARED
			ENDIF
		ENDIF
		WAIT CLEAR

		IF !USED(m.cAlias)
			THIS.Msg('Tabela '+m.cTable+'('+m.cAlias+') não foi aberta!',.T.)
			RETURN .F.
		ENDIF
		IF m.lFechar AND USED(m.cAlias)
			USE IN (m.cAlias)
		ENDIF
		RETURN .T.
	ENDFUNC

****
*
* Retorna o SQL de criação da TABELA
*
****
	FUNCTION TabelaSQL
		LPARAMETERS lcDBF


	ENDFUNC
ENDDEFINE
