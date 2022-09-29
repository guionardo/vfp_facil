SET PROCEDURE TO GS_TRATAERROS ADDITIVE 

ON ERROR DO GSTrataErros WITH ERROR(),MESSAGE(),MESSAGE(1),PROGRAM(),LINENO()

*
****************************************************
* GSTrataErros
*
PROCEDURE GSTrataErros
	LPARAMETERS lnErro, lcMens0, lcMens1, lcProc, lnLineNo
	LOCAL arqant
* Evita chamadas recursivas a TRATAERROS
	DO SetTrataErrosSimples

	IF TYPE("M.SISTEMA")#"C"
		m.SISTEMA = "GUIOSOFT"
	ENDIF

	IF TYPE("M.VERSAO")#"C"
		m.VERSAO = ""
	ENDIF

	m.lnErro = IIF(VARTYPE(m.lnErro)#"N",0,m.lnErro)
	LOCAL lcUsuAtivo, lcUsuCod

*------------------
*
* Apresenta erros  de execu��o
*
	TRY
		m.lcMens0 = DefaultTo(m.lcMens0,"")
		m.lcMens1 = DefaultTo(m.lcMens1,"")
		m.lcProc =  DefaultTo(m.lcProc,"")
		m.lnLineNo = DefaultTo(m.lnLineNo,0)

		m.lcUsuCod = IIF(PEMSTATUS(_SCREEN,"CODUSU",5),_SCREEN.CODUSU,"   ")
		m.lcUsuAtivo = IIF(PEMSTATUS(_SCREEN,"APEUSU",5),_SCREEN.APEUSU,"SEM LOGIN")

		LOCAL oErr AS EXCEPTION

		LOCAL lSai, lRetry
		m.lSai = .NULL.
		m.lRetry = .F.
		DO CASE
			CASE M.lnErro = 5
				Erros('Aparentemente, um registro foi posicionado fora da capacidade de uma tabela.\n'+;
					'Uma causa poss�vel � a corrup��o do arquivo de �ndice.\n'+;
					'Proceda a reindexa��o das tabelas do sistema '+M.SISTEMA)
				m.lSai = .T.

			CASE M.lnErro = 0 .AND. ALLTRIM(UPPER(M.lcMens0)) == 'STRUCTURAL CDX FILE NOT FOUND.'
				m.lSai = .T.

			CASE m.lnErro = 13 AND TYPE("M.GSSaindo")="L"
*  Ignora caso os aliases estejam fechados antes do fechamento de forms

			CASE M.lnErro = 30
				Erros('\nFavor configurar a resolu��o de v�deo para 800x600\n')

			CASE M.lnErro = 43
				IF Sim('Esta opera��o est� impossibilitada por falta de mem�ria.\n' +;
						'Feche os demais aplicativos abertos e escolha uma das op��es abaixo...\n\n' +;
						'Deseja tentar novamente ?\n')
					ReduceMemory()
					m.lRetry = .T.
				ENDIF

			CASE m.lnErro = 56
				IF Sim('\nN�o h� espa�o suficiente em disco.\n'+M.lcMens0+'\n'+m.lcMens1+'\n\nTentar novamente?\n')
					m.lRetry = .T.
				ENDIF
			CASE M.lnErro = 125
				IF Sim('\nImpressora desligada. Tentar novamente ?\n')
					m.lRetry = .T.
				ENDIF

			CASE lnErro = 202
				IF Sim('O caminho indicado � inv�lido. Verifique a exist�ncia da pasta de destino do arquivo. Ex.: \AUTOCOM\EXPORTA\'+ SPACE(10)+'\n\n'+;
						'Deseja tentar novamente ? ')
					m.lRetry = .T.
				ELSE
					m.lSai = .F.
				ENDIF

			CASE TRANSFORM(M.lnErro) $ '1102-1105' .AND. 'MERGE' $ UPPER(m.lcProc )
				IF Sim('Impressora desligada. Tentar novamente ?')
					RETRY
				ELSE
					m.lSai = .T.
				ENDIF
			CASE m.lnErro = 1298 	&& Largura do report � maior do que a impressora
				Erros("A impressora selecionada n�o comporta a largura do relat�rio.\n"+;
					"Impressora padr�o do sistema: "+SET("Printer",2)+'\n'+;
					"Escolha a impressora adequada ou configure adequadamente a impressora padr�o do Windows")
				m.lSai = .F.

			CASE m.lnErro = 1683	&& Tag/order n�o existe
				Erros("O programa solicitou um �ndice inexistente.\n"+;
					"Acesse este programa usando o par�metro INDEXAR.\n"+;
					"Se mesmo assim o erro persistir, entre em contato com o suporte.\n\n"+;
					'\n. Programa: '+ M.lcProc +"("+TRANSFORM(m.lnLineNo)+")" )
				m.lSai = .T.

			CASE M.lnErro = 1705  OR; && Violacao de compartilhamento
				m.lnErro = 1707 OR ;&& O Arquivo de indices (CDX) n�o presente
				m.lnErro = 1164 && Browse structure has changed -> Dentro da fun��o c�digo ao chamar o MOSTRAFOTO
				m.lSai = .T.

			CASE M.lnErro = 1924 .AND. ALLTRIM(UPPER(M.lcProc)) == 'VARREAD_'
				Erros('Esta fun��o deve ser utilizada somente sobre um campo v�lido.')
				m.lSai = .T.

			CASE ' OLE' $ UPPER(M.lcMens0+M.lcMens1)
				Erros('Problemas no processamento da imagem. Erro n�:'+ALLTRIM(STR(M.mError)))
				m.lSai = .T.
			CASE M.lnErro = 1733
				Erros('Classe n�o encontrada!\n'+m.lcMens0+'\n'+m.lcMens1)
				m.lSai = .F.

			CASE ('NOT A TABLE' $ ALLTRIM(UPPER(M.lcMens0)))
				m.txtlinha = SUBSTR(m.lcMens1,AT('AUTOC-',UPPER(m.lcMens1),8))

				IF !ISALPHA(RIGHT(m.txtlinha,1))
					m.txtlinha = LEFT(m.txtlinha,7)
				ENDIF
				m.numero = SELECT(0)
				IF m.numero > 0
					SELE (m.numero)
				ENDIF
				Erros('\nAparentemente, um problema de "header" foi encontrado...\n\n'+;
					'A �ltima tabela aberta com sucesso foi: \n'+;
					DBF(ALIAS())+ ' ('+ALLTRIM(ALIAS())+')\n\n'+;
					'As tabelas s�o abertas no formato: "ordem alfab�tica", portanto\n'+;
					'sugiro executar o utilit�rio TIMEHEAD para a tabela seguinte a\n'+;
					'tabela indicada na linha acima.\n\n'+;
					'Ap�s o reparo, n�o esque�a de indicar o par�metro "INDEXAR".\n\n')

			OTHERWISE
				LOCAL lcCS
				m.lcCS = CallStackSimples(2)
				m.cME = 'Por favor, anote a msg e entre em contato com a Guiosoft Inform�tica, atrav�s de um dos canais, abaixo:\n'+;
					'* Fone: (47) 8805-0705\n* suporte@guionardofurlan.com.br\n* Skype guionardo\n\n' +;
					'\n. Erro('+TRANSFORM(m.lnErro)+') '+ ALLTRIM(lcMens0)          +;
					'\n. Programa: '         + M.lcProc +"("+TRANSFORM(m.lnLineNo)+")"+;
					'\n. Usu�rio: '          + m.lcUsuAtivo+;
					'\n. Net/ID: '           + NetInfo()+;
					'\n. Vers�o ('+M.SISTEMA+'): ' + TRANSFORM(M.VERSAO) + " "+ArqVersao(APPLICATION.SERVERNAME)+;
					'\n. Windows: '			 + GetWinVer()+;
					'\n. Tabela em uso: '    + ALLTRIM(ALIAS())+;
					'\n. Tempor�rios: '      + ALLTRIM(SYS(2023))    +;
					'\n. Data Sistema: '     + TTOC(DATETIME())  +;
					IIF(!EMPTY(ALLTRIM(m.lcMens1)),'\n. Comando: ' + ALLTRIM(M.lcMens1),'')+'\n\n\n'

* '\n. Callstack: '		 + m.lcCS +;


				_CLIPTEXT = STRTRAN(m.cME,'\n',CHR(13)+CHR(10))
				Erros(M.cME,"ERRO na execu��o do sistema "+m.SISTEMA)
				Mensagem("A mensagem de erro foi copiada para a �rea de transfer�ncia.")

				LOCAL loLog AS LOG
				m.loLog = CREATEOBJECT("LOG",JUSTSTEM(APPLICATION.SERVERNAME)+".EXCEPTION.LOG",0)
				m.loLog.addLog("Erro      #"+TRANSFORM(m.lnErro)+": "+ALLTRIM(m.lcMens0))
				m.loLog.addLog("Programa  "+M.lcProc+"("+TRANSFORM(m.lnLineNo)+")")
				m.loLog.addLog("Sistema   "+M.SISTEMA+" ("+TRANSFORM(m.VERSAO)+ " "+ArqVersao(APPLICATION.SERVERNAME)+")")
				m.loLog.addLog("Windows   "+GetWinVer())
				m.loLog.addLog("Net/ID    "+NetInfo()+" "+_SCREEN.NOMUSU)
				m.loLog.addLog("CallStack "+m.lcCS)
				m.loLog.addLog("*")
				m.loLog.DESTROY()
		ENDCASE
	CATCH TO m.oErr
		Erros(m.oErr)
	ENDTRY
	IF m.lRetry
		RETRY
	ENDIF
	IF !ISNULL(m.lSai)
		RETURN m.lSai
	ENDIF
	WAIT WINDOW NOWAIT 'Aguarde... Fechando as tabelas...'
	FLUSH
	WAIT WINDOW NOWAIT 'Todas as tabelas foram fechadas...'
	? CHR(7)+CHR(7)+CHR(7)
	WAIT CLEAR
	CANCEL

ENDPROC
****
*
* Define TrataErros Simples
*
****
PROCEDURE SetTrataErrosSimples
	ON ERROR DO MESSAGEBOX("Erro:"+TRANSFORM(ERROR())+CHR(13)+;
		"Mensagem:"+MESSAGE()+CHR(13)+;
		"C�digo:"+MESSAGE(1)+CHR(13)+;
		"Programa:"+PROGRAM()+CHR(13)+;
		"Linha:"+TRANSFORM( LINENO()))

*"Callstack:"+CallStackSimples()+CHR(13)+;

ENDPROC
****
*
* CallStackSimples
*
* Parametro lnIgnoraUltimos n�o mostra ultimos "n" da lista
*
****
FUNCTION CallStackSimples
	LPARAMETERS lnIgnoraUltimos
	LOCAL lnI,lcR,lnC,laCS[1],lnPath
	m.lcR = ""
	m.lnI = 1
	m.lnC = ASTACKINFO(m.laCS)
	m.lnIgnoraUltimos = DefaultTo(m.lnIgnoraUltimos,0)
	IF m.lnIgnoraUltimos>=m.lnC
		m.lnIgnoraUltimos = 0
	ENDIF
	m.lnPath = LEN(JUSTPATH(m.laCS(1,2)))+1
	FOR m.lnI = 1 TO m.lnC - m.lnIgnoraUltimos
		m.lcR = m.lcR + SUBSTR(m.laCS(m.lnI,2),IIF(EMPTY(JUSTPATH(m.laCS(m.lnI,2))),1,m.lnPath))+"("+TRANSFORM(m.laCS(m.lnI,5))+") / "
	NEXT
	IF LEN(m.lcR)>3
		m.lcR = LEFT(m.lcR,LEN(m.lcR)-3)
	ENDIF
	RETURN m.lcR
ENDFUNC
