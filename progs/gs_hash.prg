****
*
* Controle de Hash dos registros
*
****
****
*
* HashOK
*
****
FUNCTION HashOk
	LPARAMETERS lcAlias
	IF VARTYPE(m.lcAlias)#"C"
		m.lcAlias = ALIAS()
	ENDIF
	IF (!USED(m.lcAlias)) OR EMPTY(FIELD("HASHNO",m.lcAlias))
		RETURN .F.
	ENDIF
	LOCAL lcAl,lnHash,llRet
	m.lcAl = ALIAS()
	SELECT (m.lcAlias)
	m.lnHash = VAL(SYS(2017,"HASHNO",1,2))
	m.llRet = (m.lnHash=EVALUATE(m.lcAlias+".HASHNO"))
	IF USED(m.lcAl)
		SELECT (m.lcAl)
	ENDIF
	RETURN m.llRet
ENDFUNC
****
*
* UpdateHash
*
****
FUNCTION HashUpdate
	LPARAMETERS lcAlias
	IF VARTYPE(m.lcAlias)#"C"
		m.lcAlias = ALIAS()
	ENDIF
	IF (!USED(m.lcAlias)) OR EMPTY(FIELD("HASHNO",m.lcAlias))
		RETURN .F.
	ENDIF
	LOCAL lcAl, llRL
	m.lcAl = ALIAS()
	SELECT (m.lcAlias)
	m.llRL = ISRLOCKED(RECNO(m.lcAlias),m.lcAlias)
	IF !m.llRL
		IF !RegLock(m.lcAlias,"Atualização de HASH")
			IF USED(m.lcAl)
				SELECT (m.lcAl)
			ENDIF
			RETURN .F.
		ENDIF
	ENDIF
	REPLACE HASHNO WITH VAL(SYS(2017,"HASHNO",1,2))
	IF !m.llRL
		RegUnlock(m.lcAlias)
	ENDIF
	WAIT WINDOW "Hash de "+m.lcAlias+"#"+TRANSFORM(RECNO(m.lcAlias))+" atualizado." TIMEOUT 5 NOWAIT 
	IF USED(m.lcAl)
		SELECT (m.lcAl)
	ENDIF

	RETURN .T.
ENDFUNC