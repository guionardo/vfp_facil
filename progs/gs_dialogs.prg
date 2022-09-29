****
*
* Mostra Mensagem de Erro
*
****
FUNCTION Erros
	LPARAMETERS lcMsg,lcCaption
	IF VARTYPE(m.lcMsg)="O" AND (UPPER(m.lcMsg.CLASS)=="EXCEPTION")
		m.lcMsg = TRANSFORM(m.lcMsg.ERRORNO)+':'+m.lcMsg.MESSAGE+;
			IIF(!EMPTY(m.lcMsg.PROCEDURE),'\nPrograma:'+m.lcMsg.PROCEDURE,"")+;
			IIF(!EMPTY(m.lcMsg.LINENO),'\nLinha:'+TRANSFORM(m.lcMsg.LINENO),"")+;
			IIF(!EMPTY(m.lcMsg.LINECONTENTS),'\n'+m.lcMsg.LINECONTENTS,"")+;
			IIF(!EMPTY(m.lcMsg.USERVALUE),'\n'+TRANSFORM(m.lcMsg.USERVALUE),'')
*
* Cria log de exceções
*
		LOCAL loExceptLog AS LOG
		m.loExceptLog = CREATEOBJECT("LOG",FORCEEXT(JUSTFNAME(APPLICATION.SERVERNAME),"EXCEPTION.LOG"))
		m.loExceptLog.AddLog(m.lcMsg)
	ENDIF
	IF VARTYPE(m.lcMsg)#"C"
		m.lcMsg = TRANSFORM(m.lcMsg)
	ENDIF
	IF VARTYPE(m.lcCaption)#"C"
		m.lcCaption = M.Sistema
	ENDIF
	MESSAGEBOX(STRTRAN(m.lcMsg,"\n",CHR(13)),0 + 48,m.lcCaption,10000)
ENDFUNC

****
*
* Pergunta
*
****
FUNCTION Sim
	LPARAMETERS lcMsg, lcCaption
	IF VARTYPE(m.lcMsg)#"C"
		m.lcMsg = TRANSFORM(m.lcMsg)
	ENDIF
	IF VARTYPE(m.lcCaption)#"C"
		m.lcCaption = M.Sistema
	ENDIF
	RETURN MESSAGEBOX(STRTRAN(m.lcMsg,"\n",CHR(13)),4+32,m.lcCaption)=6
ENDFUNC
****
*
* Mensagem
*
****
FUNCTION Mensagem
	LPARAMETERS lcMsg,lcCaption
	IF VARTYPE(m.lcMsg)#"C"
		m.lcMsg = TRANSFORM(m.lcMsg)
	ENDIF
	IF VARTYPE(m.lcCaption)#"C"
		m.lcCaption = M.Sistema
	ENDIF
	MESSAGEBOX(STRTRAN(m.lcMsg,"\n",CHR(13)),0 + 64,m.lcCaption,10000)
ENDFUNC

