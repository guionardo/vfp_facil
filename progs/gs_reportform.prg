***********************************************************
* ReportForm - Imprime/visualiza report
* lcREPORT	-> Nome do report
* lcTitulo	-> Titulo
* lcOpcoes	-> Opções
PROCEDURE ReportForm
	LPARAMETERS lcReport, lcTitulo, lcOpcoes
	LOCAL lcImpressoraPadrao, llExterno
	LOCAL lcCom, lcRepo, lcPrint


	* Impressão/visualização direta

	m.lcImpressoraPadrao = GetImpressoraPadrao(m.lcReport)
	IF EMPTY(m.lcImpressoraPadrao)
		m.lcImpressoraPadrao = SET("Printer",2)
	ENDIF

	LOCAL lcPRep
	IF TYPE("M.pcBASE")="C"
		m.lcPRep = m.PCBase
	ELSE
		m.lcPRep = JUSTPATH(APPLICATION.SERVERNAME)
	ENDIF
	IF FILE(FORCEPATH(m.lcReport+"_.FXP",ADDBS(M.lcPRep)+"REPORTS")) AND ;
			FILE(FORCEPATH(m.lcReport+"_.FXT",ADDBS(m.lcPRep)+"REPORTS"))
		m.lcRepo = FORCEPATH(m.lcReport+"_.FXP",ADDBS(M.lcPRep)+"REPORTS")
	ELSE
		m.lcRepo = m.lcReport
	ENDIF
	m.lcCom = "V"

	IF !m.lcCom$"VP"
		RETURN
	ENDIF

	* Copia report para área temporária
	LOCAL lcNovoReport, laFiles(1)
	LOCAL lcTemp
	m.lcTemp = ADDBS(JUSTPATH(APPLICATION.SERVERNAME))+"TEMP"
	*!*	IF !DIRECTORY(m.lcTemp)
	*!*		MKDIR (m.lcTemp)
	*!*	ENDIF
	m.lcNovoReport = ["]+ADDBS(SYS(2023))+"Rep"+STRTRAN(DefaultTo(m.lcTitulo,m.lcReport),' ','_')+["]

	IF ADIR(m.laFiles,FORCEEXT(m.lcRepo,"FR?"))=2	&& Se o report for encontrado no disco
		CopiarArquivo(FORCEEXT(m.lcRepo,"FRX"),FORCEEXT(m.lcNovoReport,"FRX"))
		CopiarArquivo(FORCEEXT(m.lcRepo,"FRT"),FORCEEXT(m.lcNovoReport,"FRT"))
	ELSE
		IF FILE(FORCEEXT(m.lcRepo,"FRX")) AND FILE(FORCEEXT(m.lcRepo,"FRT"))
			STRTOFILE(FILETOSTR(FORCEEXT(m.lcRepo,'FRX')),FORCEEXT(m.lcNovoReport,'FRX'))
			STRTOFILE(FILETOSTR(FORCEEXT(m.lcRepo,'FRT')),FORCEEXT(m.lcNovoReport,'FRT'))
		ENDIF
	ENDIF

	IF ADIR(m.laFiles,FORCEEXT(m.lcNovoReport,"FR?"))#2	&& Report não existe na pasta temporária
		Erros("Não foi possível gravar o report na área temporária!\n\n"+;
			m.lcNovoReport)
		RETURN
	ENDIF

	* Chegando aqui, os arquivos do report estarão na pasta temporária
	* renomeado para o título do report

	* Remove impressora padrão do report
	IF .T.
		LOCAL lnTenta
		m.lnTenta = 5
		DO WHILE m.lnTenta > 0
			TRY
				USE (FORCEEXT(m.lcNovoReport,"FRX")) IN 0 ALIAS FRXTEMP
			CATCH

			ENDTRY
			IF !USED("FRXTEMP")
				WAIT WINDOW "Abrindo o report "+m.lcNovoReport+" exclusivamente...("+TRANSFORM(m.lnTenta)+")" TIMEOUT 0.5
				m.lnTenta = m.lnTenta - 1
				LOOP
			ENDIF

			IF !ISREADONLY("FRXTEMP")
				m.lnTenta = 0
			ELSE
				WAIT WINDOW "Abrindo o report "+m.lcNovoReport+" exclusivamente...("+TRANSFORM(m.lnTenta)+")" TIMEOUT 0.5
				m.lnTenta = m.lnTenta - 1
			ENDIF
		ENDDO
		IF USED("FRXTEMP") AND !ISREADONLY("FRXTEMP")
			GOTO TOP IN FRXTEMP
			REPLACE ;
				EXPR WITH STRTRAN(FRXTEMP.EXPR,"DEVICE=","DEVICE_OLD="),;
				TAG WITH "",;
				TAG2 WITH "" ;
				IN FRXTEMP
		ELSE
			LOCAL lnI, lcAl
			m.lcAl = ""
			FOR m.lnI = 1 TO 240
				IF USED(m.lnI)
					m.lcAl = m.lcAl + ALIAS(m.lnI)+'='+DBF(m.lnI)+"\n"
				ENDIF
			NEXT
			Mensagem(m.lcAl,"Aliases")
			Mensagem("Não foi possível remover a propriedade de impressora padrão do report "+m.lcNovoReport)
		ENDIF
		IF USED("FRXTEMP")
			USE IN FRXTEMP
		ENDIF

	ENDIF

	* Define impressora destino do report
	LOCAL lcImpAnt
	m.lcImpAnt = SET("Printer",2)
	IF !EMPTY(m.lcImpressoraPadrao)
		IF m.lcImpAnt#m.lcImpressoraPadrao
			LOCAL loExc AS EXCEPTION
			TRY
				SET PRINTER TO NAME (ALLTRIM(m.lcImpressoraPadrao))
			CATCH TO m.loExc
				Erros("Erro na tentativa de definir a impressora padrão para\n\n"+;
					"["+m.lcImpressoraPadrao+"]\n\n"+;
					"Será usada a configuração padrão do Windows")
			ENDTRY

		ENDIF
	ELSE
		SET PRINTER TO DEFAULT
	ENDIF

	IF VARTYPE(m.lcTitulo)#"C"
		m.lcTitulo = "Visualizar relatório"
	ENDIF
	IF VARTYPE(m.lcOpcoes)#"C"
		m.lcOpcoes = "TO PRINTER PROMPT NODIALOG PREVIEW"
	ENDIF

	IF m.lcCom = "V"
		IF !"PREVIEW"$m.lcOpcoes
			m.lcOpcoes = m.lcOpcoes + " PREVIEW"
		ENDIF
	ELSE
		m.lcOpcoes = STRTRAN(m.lcOpcoes,"PREVIEW","")
	ENDIF
	LOCAL nWidth, nHeight
	IF _SCREEN.VISIBLE
		m.nWidth = _SCREEN.WIDTH
		m.nHeight = _SCREEN.HEIGHT
	ELSE
		IF _SCREEN.FORMCOUNT>0
			m.nWidth = _SCREEN.FORMS(1).WIDTH
			m.nHeight = _SCREEN.FORMS(1).HEIGHT
		ELSE
			m.nWidth = SYSMETRIC(21)
			m.nHeight = SYSMETRIC(22)
		ENDIF
	ENDIF
	nWidth = m.nWidth/FONTMETRIC(6,_SCREEN.FONTNAME,_SCREEN.FONTSIZE)
	nHeight = m.nHeight/FONTMETRIC(1,_SCREEN.FONTNAME,_SCREEN.FONTSIZE)-2

	DEFINE WINDOW xview AT 0,0 SIZE nHeight,nWidth ;
		SYSTEM CLOSE FLOAT GROW ZOOM ;
		TITLE (m.lcTitulo+" | Impressora Padrão:"+m.lcImpressoraPadrao)

	ACTIVATE WINDOW xview
*	ZOOM WINDOW xview MAX


	IF !FILE(FORCEEXT(m.lcNovoReport,"FRX"))
		Erros("Report "+m.lcNovoReport+" não foi encontrado!")
	ELSE
		KEYBOARD "{CTRL-F10}"
		REPORT FORM &lcNovoReport &lcOpcoes WINDOW xview &&(.NAME)
	ENDIF
	RELEASE WINDOW xview

	* Exclui arquivos temporários
	ExcluirArquivo(FORCEEXT(m.lcNovoReport,'FRX'))
	ExcluirArquivo(FORCEEXT(m.lcNovoReport,'FRT'))

	IF SET("Printer",2)#m.lcImpAnt
		SET PRINTER TO NAME (m.lcImpAnt)
	ENDIF

	* Verifica se o tamanho do report é menor que a impressão
	* https://www.berezniker.com/content/pages/visual-foxpro/enumerating-printer-forms
	* https://www.berezniker.com/content/pages/visual-foxpro/changing-windows-default-printer
	*!*		USE (m.lcNovoReport) IN 0 ALIAS TMPREPORTFORM SHARED
	*!*		GOTO TOP IN TMPREPORTFORM
	*!*		LOCAL lcExpr,lnPaperLength,lnPaperWidth,lnP
	*!*		m.lcExpr = TMPREPORTFORM.EXPR
	*!*		USE IN TMPREPORTFORM
	*!*		m.lnP = AT("PAPERLENGTH=",m.lcExpr)
	*!*		m.lnPaperLength = VAL(SUBSTR(m.lcExpr,m.lnP+13))
	*!*		m.lnP = AT("PAPERWIDTH=",m.lcExpr)
	*!*		m.lnPaperWidth = VAL(SUBSTR(m.lcExpr,m.lnP+12))


	RETURN

ENDPROC

*
*
FUNCTION GetImpressoraPadrao
	LPARAMETERS lcReport
	IF USED("IMPRESSORAPADRAO")
		IF SEEK(PADR(M.SISTEMA,20)+PADR(m.lcReport,30)+PADR(NetInfo(),40),"IMPRESSORAPADRAO",1)
			RETURN IMPRESSORAPADRAO.IMPPAD
		ENDIF
	ENDIF
	RETURN ""
ENDFUNC
*
FUNCTION SetImpressoraPadrao
	LPARAMETERS lcReport,lcImpPad
	LOCAL lcCODUSU,lcAPEUSU
	IF TYPE("M.CodigoUsu")="C"
		m.lcCODUSU = m.CodigoUsu
	ELSE
		m.lcCODUSU = ""
	ENDIF
	IF TYPE("M.ApeUsuAtivo")="C"
		m.lcAPEUSU = M.ApeUsuAtivo
	ELSE
		m.lcAPEUSU = ""
	ENDIF
	IF USED("IMPRESSORAPADRAO")
		IF SEEK(PADR(M.SISTEMA,20)+PADR(m.lcReport,30)+PADR(NetInfo(),40),"IMPRESSORAPADRAO",1)
			IF RegLock("IMPRESSORAPADRAO")
				REPLACE ;
					IMPPAD WITH m.lcImpPad,;
					CODUSU WITH m.lcCODUSU,;
					APEUSU WITH m.lcAPEUSU,;
					DATALT WITH DATETIME();
					IN IMPRESSORAPADRAO
				RegUnlock("IMPRESSORAPADRAO")
				RETURN .T.
			ENDIF
		ELSE
			INSERT INTO IMPRESSORAPADRAO (SISTEM,NOMREP,NETUSU,IMPPAD,CODUSU,APEUSU,DATALT) VALUES ;
				(M.SISTEMA, m.lcReport, NetInfo(), m.lcImpPad, m.lcCODUSU, m.lcAPEUSU, DATETIME())
			RETURN _TALLY>0
		ENDIF
	ENDIF
	RETURN .F.
ENDFUNC
