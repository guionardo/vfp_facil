****
*
* Funções de tratamento de strings, números, etc
*
****

****
*
* STRZERO 		Análogo a STR, com preenchimento de zeros a esquerda
*
****
FUNCTION StrZero
	LPARAMETERS lnVal, lnTam, lnDec
	IF VARTYPE(m.lnDec)#"N" OR !BETWEEN(m.lnDec,0,10)
		m.lnDec = 0
	ENDIF
	IF VARTYPE(m.lnTam)#"N" OR !BETWEEN(m.lnTam,1,20)
		m.lnTam = 10
	ENDIF
	IF VARTYPE(m.lnVal)#"N"
		m.lnVal = 0
	ENDIF
	RETURN STRTRAN(STR(m.lnVal,m.lnTam,m.lnDec)," ","0")
ENDFUNC
****
*
* DefaultTo		Retorna o valor default de um parâmetro
*
* lxVal		Valor
* lxDef		Caso lxVal não seja do tipo de lxDef, este é retornado
*
****
FUNCTION DefaultTo
	LPARAMETERS lxVal, lxDef
	IF VARTYPE(m.lxVal)#VARTYPE(m.lxDef)
		RETURN m.lxDef
	ENDIF
	RETURN m.lxVal
ENDFUNC
****
*
* NetInfo		Retorna //COMPUTADOR/USUÁRIO/Sessão (CONSOLE ou RDP)
*
****
FUNCTION NetInfo
	LOCAL li,lcSessionName, lcClientName
	m.li = AT('#',SYS(0))
	m.lcSessionName = IIF("RDP"$GETENV('SESSIONNAME'),"RDP",GETENV('SESSIONNAME'))
	m.lcClientName = IIF(!EMPTY(GETENV('CLIENTNAME')),'\'+GETENV('CLIENTNAME'),'')
	RETURN '\\'+ALLTRIM(LEFT(SYS(0),m.li-1))+'\'+;
		ALLTRIM(SUBSTR(SYS(0),m.li+1))+'\'+;
		m.lcSessionName+m.lcClientName
ENDFUNC
****
*
* ArqVersao		Retorna a versão gravada nos detalhes do executável
*
* lcArq		Nome do arquivo (default = executável atual)
* Retorno	STRING
*
****
FUNCTION ArqVersao
	LPARAMETERS lcArq
	IF VARTYPE(m.lcArq)#"C"
		m.lcArq = APPLICATION.SERVERNAME
	ENDIF
	IF !FILE(m.lcArq)
		RETURN ""
	ENDIF
	LOCAL ARRAY laVer(1)
	IF AGETFILEVERSION(m.laVer,m.lcArq)>0
		IF TYPE("M.laVer(4)")="C"
			RETURN m.laVer(4)
		ENDIF
	ENDIF

	RETURN ""
ENDFUNC
****
*
* FormAtivo - Retorna se Form de nome lcForm está ativo
*
****
FUNCTION FormAtivo
	LPARAMETERS lcForm
	IF VARTYPE(m.lcForm)#"C"
		RETURN .F.
	ENDIF
	m.lcForm = ALLTRIM(UPPER(m.lcForm))
	FOR EACH loF AS FORM IN _SCREEN.FORMS
		IF UPPER(m.loF.NAME) == m.lcForm
			RETURN .T.
		ENDIF
	ENDFOR
	RETURN .F.
ENDFUNC
