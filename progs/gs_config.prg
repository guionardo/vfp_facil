****
*
* Configurações
*
****

****
*
* Inicializa as configurações
*
****
FUNCTION Config_Init
	LPARAMETERS lcPathBase, lcTableName
	IF TYPE("M.PubConfigKey1")="U"
		PUBLIC PubConfigKey1
	ENDIF
	IF TYPE("M.PubConfigKey2")="U"
		PUBLIC PubConfigKey2
	ENDIF
	IF (VARTYPE(m.PubConfigKey1)#"C") OR ;
			(VARTYPE(m.PubConfigKey2)#"C") OR ;
			(LEN(m.PubConfigKey1)#255) OR ;
			(LEN(m.PubConfigKey2)#255)
		LOCAL lnI
		STORE "" TO ;
			M.PubConfigKey1,;
			M.PubConfigKey2
		FOR m.lnI = 1 TO 255
			m.PubConfigKey1 = m.PubConfigKey1 + CHR(m.lnI)
		NEXT
		m.PubConfigKey2 = SUBSTR(m.PubConfigKey1,128)+LEFT(m.PubConfigKey1,127)
	ENDIF

	IF USED("CONFIG")
		RETURN .T.
	ENDIF
	IF VARTYPE(m.lcTableName)#"C"
		m.lcTableName = JUSTSTEM(APPLICATION.SERVERNAME)+"_CFG"
	ENDIF
	m.lcTableName = FORCEEXT(m.lcTableName,'DBF')
	IF VARTYPE(m.lcPathBase)#"C"
		m.lcPathBase = ADDBS(JUSTPATH(APPLICATION.SERVERNAME))+"DB"
	ENDIF

	IF !DIRECTORY(m.lcPathBase)
		MKDIR (m.lcPathBase)
	ENDIF

	IF !DIRECTORY(m.lcPathBase)
		Erros("Não foi possível criar a pasta de dados\n"+;
			M.lcPathBase+"\n","Inicialização de Configuração")
		RETURN .F.
	ENDIF

	m.lcTableName = FORCEPATH(m.lcTableName,m.lcPathBase)

	IF !FILE(m.lcTableName)
		CREATE TABLE (m.lcTableName) (;
			SECAO 		C(10),;
			CHAVE 		C(10),;
			CODUSU 		C(3),;
			LASTUPDATE 	T(8),;
			VALOR 		M(4),;
			SECURE		L(1),;
			HASH 		I(4))
		USE
	ENDIF
	USE (m.lcTableName) IN 0 ALIAS CONFIG SHARED

	RETURN USED("CONFIG")
ENDFUNC

****
*
* Grava Configuração
*
*
* lcSecao	Seção (default = nome do aplicativo)
* lcChave	Chave
* lcVal		Valor
* lcUsu		Código do usuário (default = SCREEN.CODUSU ou "")
* llSecure	Valor criptografado (algoritmo simples, de substituição de caracteres)
*
* Retorna .t. em caso de sucesso
****
FUNCTION SetConfig
	LPARAMETERS lcSecao, lcChave, lcVal, lcUsu, llSecure
	IF !Config_Init()
		RETURN .F.
	ENDIF
	m.lcSecao = DefaultTo(m.lcSecao,JUSTSTEM(APPLICATION.SERVERNAME))
	IF VARTYPE(m.lcChave)#"C"
		Erros("SetConfig: Parâmetro inválido (CHAVE)")
		RETURN .F.
	ENDIF

	IF VARTYPE(m.lcUsu)#"C"
		IF PEMSTATUS(_SCREEN,"CODUSU",5)
			m.lcUsu = _SCREEN.CODUSU
		ELSE
			m.lcUsu = ""
		ENDIF
	ENDIF

	LOCAL lcAl, llRet
	m.lcAl = ALIAS()
	SELECT CONFIG
	LOCATE FOR (SECAO+CHAVE+CODUSU)==(PADR(m.lcSecao,10)+PADR(m.lcChave,10)+PADR(m.lcUsu,3))
	IF m.llSecure
		m.lcVal = CHRTRAN(m.lcVal,m.PubConfigKey1,m.PubConfigKey2)
	ENDIF
	IF !FOUND()
		APPEND BLANK
		REPLACE ;
			SECAO 		WITH m.lcSecao,;
			CHAVE 		WITH m.lcChave,;
			CODUSU		WITH m.lcUsu
		FLUSH IN CONFIG
		UNLOCK IN CONFIG
	ENDIF

	IF RegLock("CONFIG")
		REPLACE ;
			LASTUPDATE	WITH DATETIME(),;
			VALOR		WITH m.lcVal,;
			SECURE		WITH m.llSecure

		REPLACE HASH WITH VAL(SYS(2017,"HASH",1,2))

		FLUSH IN CONFIG
		UNLOCK IN CONFIG
		m.llRet = .T.
	ENDIF

	IF USED(m.lcAl)
		SELECT (m.lcAl)
	ENDIF
	RETURN m.llRet
ENDFUNC

****
*
* Obtém Valor de configuração
*
* lcSecao	Seção (default = nome do aplicativo)
* lcChave	Chave
* lcDefault	Valor padrão se não encontrar a chave
* lcUsu		Código do usuário (default = SCREEN.CODUSU ou "")
*
* Retorna Valor ou "" caso não encontrado
*
****
FUNCTION GetConfig
	LPARAMETERS lcSecao, lcChave, lcDefault, lcUsu

	IF !Config_Init()
		RETURN ""
	ENDIF
	m.lcSecao = DefaultTo(m.lcSecao,JUSTSTEM(APPLICATION.SERVERNAME))
	IF VARTYPE(m.lcChave)#"C"
		Erros("GetConfig: Parâmetro inválido (CHAVE)")
		RETURN ""
	ENDIF

	IF VARTYPE(m.lcUsu)#"C"
		IF PEMSTATUS(_SCREEN,"CODUSU",5)
			m.lcUsu = _SCREEN.CODUSU
		ELSE
			m.lcUsu = ""
		ENDIF
	ENDIF

	LOCAL lcAl, llRet
	m.lcAl = ALIAS()
	SELECT CONFIG
	LOCATE FOR (SECAO+CHAVE+CODUSU)==(PADR(m.lcSecao,10)+PADR(m.lcChave,10)+PADR(m.lcUsu,3))
	IF !FOUND()
		RETURN m.lcDefault
	ENDIF

	LOCAL lnHash
	IF CONFIG.HASH # VAL(SYS(2017,"HASH",1,2))
		AlertBox("Registro alterado indevidamente na tabela de configurações!",32,"ERRO: INTEGRIDADE COMPROMETIDA",10)
	ENDIF

	IF CONFIG.SECURE
		m.lcRet = CHRTRAN(CONFIG.VALOR,m.PubConfigKey2,m.PubConfigKey1)
	ELSE
		m.lcRet = CONFIG.VALOR
	ENDIF

	IF USED(m.lcAl)
		SELECT (m.lcAl)
	ENDIF
	RETURN m.lcRet
ENDFUNC
