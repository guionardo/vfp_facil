****
*
* Auxiliares de Tabelas e Aliases
*
****

****
*
* PushAlias - Grava Alias/Recno/Order do alias atual em uma lista FIFO
*
****
FUNCTION PushAlias
	IF EMPTY(ALIAS())
		RETURN .F.
	ENDIF
	IF TYPE("M._PUFiFoAlias")="U"
		PUBLIC _PUFiFoAlias
		m._PUFiFoAlias = ""
	ENDIF
	m._PUFiFoAlias = PADR(ALIAS(),20) + BINTOC(RECNO(),4) + STR(TAGNO(),2,0) + m._PUFiFoAlias
	RETURN .T.
ENDFUNC
****
*
* PopAlias - Restaura Alias/Recno/Order do último alias registrado por PushAlias
*
****
FUNCTION PopAlias
	IF TYPE("M._PUFiFoAlias")="U"
		PUBLIC _PUFiFoAlias
		m._PUFiFoAlias = ""
	ENDIF
	IF EMPTY(m._PUFiFoAlias)
		RETURN .F.
	ENDIF
	LOCAL lcAlias, lnRN, lnOr
	m.lcAlias = ALLTRIM(LEFT(m._PUFiFoAlias,20))
	m.lnRN = CTOBIN(SUBSTR(m._PUFiFoAlias,21,4))
	m.lnOr = VAL(SUBSTR(m._PUFiFoAlias,25,2))
	m._PUFiFoAlias = SUBSTR(m._PUFiFoAlias,27)
	IF !USED(m.lcAlias)
		RETURN .F.
	ENDIF
	SELECT (m.lcAlias)
	IF BETWEEN(m.lnRN,1,RECCOUNT(m.lcAlias))
		GOTO (m.lnRN) IN (m.lcAlias)
	ELSE
		GOTO BOTTOM IN (m.lcAlias)
		IF !EOF(m.lcAlias)
			SKIP IN (m.lcAlias)
		ENDIF
	ENDIF
	IF BETWEEN(m.lnOr,1,TAGCOUNT(CDX(1,m.lcAlias),m.lcAlias))
		SET ORDER TO (m.lnOr)
	ELSE
		SET ORDER TO 0
	ENDIF
	RETURN .T.
ENDFUNC
