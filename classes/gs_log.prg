*
* Controle de log
*
* 2015-06-09
* Guionardo Furlan

***********************************************************
* GSInitLog
* Inicializa o arquivo de LOG texto da aplicação
*
FUNCTION GSInitLog
	IF !PEMSTATUS(_SCREEN,"GSLOG",5)
		ADDPROPERTY(_SCREEN,"GSLOG",.F.)
	ENDIF

	IF VARTYPE(_SCREEN.GSLOG)#"O"
		_SCREEN.GSLOG = CREATEOBJECT("LOG")
	ENDIF

	RETURN (VARTYPE(_SCREEN.GSLOG)=="O")
ENDPROC
*
***********************************************************
* GSCheckLog
* Verifica se o LOG foi iniciado
*
PROCEDURE GSCheckLog
	IF !GSInitLog()
		WAIT WINDOW "ERRO: ACLOG não inicializou!" TIMEOUT 5
	ENDIF
ENDPROC
*
***********************************************************
* GSAddLog
* Adiciona texto ao log
* lcMsg -> Texto
* llSemDataHora -> opcional, retira a data e hora do início da linha
PROCEDURE GSAddLog
	LPARAMETERS lcMsg

	IF GSInitLog()
		_SCREEN.GSLOG.AddLog(m.lcMsg)
	ENDIF

ENDPROC
*
***********************************************************
* GSCallStack
* Retorna o call stack atual
*
FUNCTION GSCallStack
	LOCAL cStack, nc, nSC
	m.cStack = ""
	LOCAL ARRAY aSI(1,6)
	m.nSC = ASTACKINFO(m.aSI)
	FOR m.nc = MAX(1,m.nSC-2) TO 1 STEP -1
		m.cStack = m.cStack + ;
			'Nvl '+TRANSFORM(m.aSI(m.nc,1))+'\n'+;
			IIF(!EMPTY(m.aSI(m.nc,4)),'Mód/Obj:'+TRANSFORM(m.aSI(m.nc,4)),'')+;
			IIF(!EMPTY(m.aSI(m.nc,3)),'('+TRANSFORM(m.aSI(m.nc,3))+')','')+;
			IIF(!EMPTY(m.aSI(m.nc,5)),'('+TRANSFORM(m.aSI(m.nc,5))+')','')+'\n'+;
			IIF(!EMPTY(m.aSI(m.nc,6)),'Fonte  :'+STRTRAN(TRANSFORM(m.aSI(m.nc,6)),CHR(13),'')+'\n','')+;
			IIF(m.nc>1,'***\n','')
	NEXT
	m.cStack = STRTRAN(m.cStack,'\n',CHR(13))
	RETURN m.cStack
ENDFUNC
****
*
* GSExceptionLog
*	Registra LOG de exceção
*
****
FUNCTION GSExceptionLog
	LPARAMETERS lcMsg
	IF VARTYPE(m.lcMsg)="O"
		m.lcMsg = "#"+TRANSFORM(m.lcMsg.ERRORNO)+" "+m.lcMsg.MESSAGE
	ENDIF
	LOCAL loExc
	m.loExc = CREATEOBJECT("GSLOG",FORCEEXT(JUSTFNAME(APPLICATION.SERVERNAME),".exception.log"))
	m.loExc.AddLog(m.lcMsg+" CS:"+GSCallStack())
ENDFUNC

*********************************************************************
* Classe de LOG
*
DEFINE CLASS LOG AS CUSTOM
	Arquivo = ""		&& Arquivo que armazena o LOG
	MaxLinha = 160		&& Tamanho máximo da linha
	LineSep = CHR(9) 	&& TAB
	AutoPurge = .T.		&& Faz a limpeza do arquivo automaticamente,
&& 	copiando as linhas do dia anterior para um backup

*
* lcArquivo: 	default		.\LOG\SISTEMA.LOG
*				nomesimples	.\LOG\NOMESIMPLES.LOG
*
	PROCEDURE INIT
		LPARAMETERS lcArquivo, lnMaxLinha, llIgnoraAutoPurge
		IF VARTYPE(m.lcArquivo)="N" AND BETWEEN(m.lcArquivo,20,250)
			m.lnMaxLinha = m.lcArquivo
		ELSE
			IF VARTYPE(m.lcArquivo)#"C"
				m.lcArquivo = ADDBS(JUSTPATH(APPLICATION.SERVERNAME))+"LOG\"+JUSTFNAME(FORCEEXT(APPLICATION.SERVERNAME,"LOG"))
			ENDIF
		ENDIF
		IF VARTYPE(m.lnMaxLinha)#"N" OR ((m.lnMaxLinha#0) AND !BETWEEN(m.lnMaxLinha,20,250))
			IF VARTYPE(m.lnMaxLinha)="L" AND m.lnMaxLinha	&& Define IgnoraAutoPurge
				THIS.AutoPurge = .F.
			ENDIF
			m.lnMaxLinha = 80
		ENDIF
		IF VARTYPE(m.llIgnoraAutoPurge)="L" AND THIS.AutoPurge
			THIS.AutoPurge = !m.llIgnoraAutoPurge
		ENDIF

		IF EMPTY(JUSTPATH(m.lcArquivo))
			m.lcArquivo = ADDBS(JUSTPATH(APPLICATION.SERVERNAME))+"LOG\"+m.lcArquivo
		ENDIF
		IF !DIRECTORY(JUSTPATH(m.lcArquivo))
			MKDIR (JUSTPATH(m.lcArquivo))
		ENDIF
		THIS.Arquivo_Assign(m.lcArquivo)
		THIS.MaxLinha = m.lnMaxLinha
	ENDPROC

	FUNCTION AddLog
		LPARAMETERS lcMsg

		m.lcMsg = THIS.DivideLinha(STRTRAN(TRANSFORM(m.lcMsg),'\n',CHR(10)))
		LOCAL lnI,lcCabeca,lcSaida
		m.lcCabeca = TTOC(DATETIME(),1)
		m.lcMsg = STRTRAN(m.lcMsg,"\n",THIS.LineSep)
		m.lcSaida = ''

		FOR m.lnI = 1 TO GETWORDCOUNT(m.lcMsg,THIS.LineSep)
			m.lcSaida = m.lcSaida + m.lcCabeca+" "+GETWORDNUM(m.lcMsg,m.lnI,THIS.LineSep)+CHR(13)+CHR(10)
		NEXT
		STRTOFILE(m.lcSaida ,THIS.Arquivo,.T.)
		RETURN .T.
	ENDFUNC

	FUNCTION DivideLinha
*
* Divide em linhas
*
		LPARAMETERS lcMsg
		IF (THIS.MaxLinha = 0) OR (LEN(m.lcMsg)<=THIS.MaxLinha)
			RETURN m.lcMsg
		ENDIF
		LOCAL lnML, lnI, lnJ, lnC
		m.lnML = SET("Memowidth")
		SET MEMOWIDTH TO (THIS.MaxLinha)

		m.lnJ = MEMLINES(m.lcMsg)
		m.lnC = ""
		FOR m.lnI = 1 TO m.lnJ
			m.lnC = m.lnC + MLINE(m.lcMsg,m.lnI,THIS.MaxLinha) + "\n"
		NEXT
		SET MEMOWIDTH TO (m.lnML)

		RETURN m.lnC
	ENDFUNC


	PROCEDURE Arquivo_Assign
		LPARAMETERS lcArquivo
		LOCAL lcPasta, oExc AS EXCEPTION
		m.lcPasta = JUSTPATH(m.lcArquivo)
		IF !DIRECTORY(m.lcPasta)
			LOCAL m.lnErro
			m.lnErro = 0
			TRY
				MKDIR (m.lcPasta)
			CATCH TO oExc
				IF oExc.ERRORNO = 202 && Invalid path or file name
					Erros("Tentativa de definição de pasta inválida.\n\n"+;
						m.lcPasta+"\n\n"+;
						"Será utilizada a pasta padrão da aplicação.\n","Classe de Log")
					m.lcPasta = JUSTPATH(APPLICATION.SERVERNAME)
				ELSE
					m.lnErro = oExc.ERRORNO
				ENDIF
			ENDTRY
			IF m.lnErro#0
				ERROR (m.lnErro),"Classe de Log"
			ENDIF
		ENDIF
		THIS.Arquivo = m.lcArquivo
		IF THIS.AutoPurge
			THIS.Purge()
		ENDIF
	ENDPROC
****
*
* Purge - Limpa o log diariamente, copiando os registros anteriores para arquivos identificados por dia
*
****
	FUNCTION Purge
		IF !FILE(THIS.Arquivo)
			RETURN .T.
		ENDIF
		LOCAL ldDC, loFL, loExc AS EXCEPTION
		TRY
			m.loFL = CREATEOBJECT('scripting.filesystemobject')
			IF VARTYPE(m.loFL)="O"
				m.ldDC = TTOD(m.loFL.GETFILE(THIS.Arquivo).DateCreated ) && DateLastModified
			ENDIF
		CATCH TO loExc
			Erros(m.loExc)
		ENDTRY

		WAIT WINDOW "ABRINDO ARQUIVO LOG "+THIS.Arquivo TIMEOUT 5 NOWAIT AT 4,1
* Verifica se a data de criação é anterior a hoje e faz o purge automaticamente
		IF m.ldDC>=DATE()
			RETURN .T.
		ENDIF
		LOCAL lcMsg, lcAl
		m.lcAl = ALIAS()
		m.lcMsg = "PURGE LOG ("+JUSTFNAME(THIS.Arquivo)+")..."
		WAIT WINDOW LEFT(m.lcMsg,255) NOWAIT AT 8,1
		CREATE CURSOR TEMPOLOG (DATA C(8),S1 C(1),HORA C(8),S C(1),MSG C(THIS.MaxLinha))
		SELECT TEMPOLOG
		APPEND FROM (THIS.Arquivo) SDF
		LOCAL ARRAY laPurge(1,1)
		m.lcMsg = m.lcMsg + CHR(13)+ "Log carregado: "+TRANSFORM(RECCOUNT("TEMPOLOG"))+" linhas."
		WAIT WINDOW LEFT(m.lcMsg,255) NOWAIT AT 12,1

		SELECT DISTINCT DATA FROM TEMPOLOG WHERE DATA#DTOS(DATE()) INTO ARRAY laPurge
		IF _TALLY = 0
			WAIT CLEAR
			USE IN TEMPOLOG
			IF USED(m.lcAl)
				SELECT (m.lcAl)
			ENDIF
			RETURN .T.
		ENDIF
		LOCAL lnI, lcBak, loExc AS EXCEPTION
		SELECT TEMPOLOG
		TRY
			FOR m.lnI = 1 TO ALEN(m.laPurge,1)
				IF LEN(TRANSFORM(VAL(m.laPurge(m.lnI,1))))<8
					m.lcBak = ADDBS(JUSTPATH(THIS.Arquivo))+'BKP\LOG\'+;
						STR(YEAR(FDATE(THIS.Arquivo)),4,0)+'\'+;
						STRZERO(MONTH(FDATE(THIS.Arquivo)),2,0)+'\'+;
						STRZERO(DAY(FDATE(THIS.Arquivo)),2,0)+'\'+;
						JUSTSTEM(THIS.Arquivo)+'.OLD.LOG'
				ELSE
					m.lcBak = ADDBS(JUSTPATH(THIS.Arquivo))+'BKP\LOG\'+;
						LEFT(m.laPurge(m.lnI,1),4)+'\'+;
						SUBSTR(m.laPurge(m.lnI,1),5,2)+'\'+;
						JUSTSTEM(THIS.Arquivo)+'.'+m.laPurge(m.lnI,1)+'.LOG'
				ENDIF
				m.lcMsg = m.lcMsg + CHR(13)+"Processando "+m.lcBak
				WAIT WINDOW LEFT(m.lcMsg,255) NOWAIT
				IF !DIRECTORY(JUSTPATH(m.lcBak))
					TRY
						MKDIR (JUSTPATH(m.lcBak))
					CATCH TO oExc
						Erros("Não foi possível criar a pasta "+JUSTPATH(m.lcBak))
					ENDTRY
				ENDIF
				COPY ALL TO (m.lcBak) FOR DATA==m.laPurge(m.lnI,1) SDF
				DELETE FOR DATA==m.laPurge(m.lnI,1)
			NEXT
			ERASE (THIS.Arquivo)
			IF FILE(THIS.Arquivo)
				WAIT WINDOW THIS.Arquivo+ " ainda existe!" TIMEOUT 3
			ENDIF
			COPY ALL TO (THIS.Arquivo) FOR DATA==DTOS(DATE()) SDF
		CATCH TO loExc
			Erros(m.loExc)
		ENDTRY
		IF USED(m.lcAl)
			SELECT (m.lcAl)
		ENDIF
		USE IN TEMPOLOG
		WAIT CLEAR
		RETURN .T.

	ENDFUNC
ENDDEFINE
