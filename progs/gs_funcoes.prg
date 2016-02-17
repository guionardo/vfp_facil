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
****
*
* Retorna apenas os números de uma string
*
****
FUNCTION SoNumero
	LPARAMETERS lcS
	IF VARTYPE(m.lcS)#"C"
		m.lcS = TRANSFORM(m.lcS)
	ENDIF
	LOCAL lnI, lcR
	m.lcR = ""
	FOR m.lnI = 1 TO LEN(m.lcS)
		IF BETWEEN(SUBSTR(m.lcS,m.lnI,1),"0","9")
			m.lcR = m.lcR + SUBSTR(m.lcS,m.lnI,1)
		ENDIF
	NEXT
	RETURN m.lcR
ENDFUNC
****
*
* AjustaTelefone
*
****
FUNCTION AjustaTelefone
	LPARAMETERS lcNF
	m.lcNF = SoNumero(m.lcNF)
	IF LEN(m.lcNF) < 10
		m.lcNF = PADL(m.lcNF,11)
	ELSE
		m.lcNF = LEFT(m.lcNF,2)+PADL(SUBSTR(m.lcNF,3),9)
	ENDIF
	RETURN m.lcNF
ENDFUNC
****
*
* FormataTelefone
*
****
FUNCTION FormataTelefone
	LPARAMETERS lcNF
	RETURN TRANSFORM(AjustaTelefone(m.lcNF),"@R (##)#####-####")
ENDFUNC
****
*
* Verifica documento CPF/CNPJ de acordo como o tamanho
*
****
FUNCTION ChkCPFCNPJ
	LPARAMETERS cDOC
	m.cDOC = ALLTRIM(TRANSFORM(m.cDOC))
	RETURN ((LEN(m.cDOC)=11) AND CPFOk(m.cDOC)) OR ;
		((LEN(m.cDOC)=14) AND CNPJOk(m.cDOC))
ENDFUNC
****
*
* Retorna o dígito do código informado
*
****
FUNCTION Digito
	LPARAMETERS lcCod, llReturnAsString AS Boolean
	IF VARTYPE(m.lcCod)=="N"
		m.lcCod = TRANSFORM(m.lcCod)
	ELSE
		IF VARTYPE(m.lcCod)#"C"
			m.lcCod = ""
		ENDIF
	ENDIF
	IF EMPTY(m.lcCod)
		RETURN IIF(m.llReturnAsString," ",0)
	ENDIF
	LOCAL lnDiv, lnPos, lnMul, lnSobra
	m.lnMul = 2
	m.lnDiv = 0
	FOR m.lnPos = LEN(m.lcCod) TO 1 STEP -1
		m.lnDiv = m.lnDiv + VAL(SUBSTR(m.lcCod,m.lnPos,1)) * m.lnMul
		m.lnMul = IIF(m.lnMul==7,2,m.lnMul + 1)
	NEXT
	m.lnSobra = m.lnDiv % 11
	m.lnDig = INT(IIF(m.lnSobra>1,11-m.lnSobra,0))
	IF m.llReturnAsString
		RETURN STR(m.lnDig,1,0)
	ENDIF
	RETURN m.lnDig
ENDFUNC
****
*
* Verifica se o dígito do número está correto
*
****
FUNCTION DigitoOk
	LPARAMETERS lcVal
	m.lcVal = TRANSFORM(VAL(TRANSFORM(m.lcVal)))
	IF LEN(m.lcVal)<2
		RETURN .F.
	ENDIF
	RETURN Digito(LEFT(m.lcVal,LEN(m.lcVal)-1),.T.)==RIGHT(m.lcVal,1)
ENDFUNC
****
*
FUNCTION CPFOk
* Parametro : CPF a verificar (C11)
* Retorna : .T. se confirmado
*------------------------------------------------------------
	PARAMETERS wcpf

	wn1 = VAL(SUBS(wcpf,01,1))
	wn2 = VAL(SUBS(wcpf,02,1))
	wn3 = VAL(SUBS(wcpf,03,1))
	wn4 = VAL(SUBS(wcpf,04,1))
	wn5 = VAL(SUBS(wcpf,05,1))
	wn6 = VAL(SUBS(wcpf,06,1))
	wn7 = VAL(SUBS(wcpf,07,1))
	wn8 = VAL(SUBS(wcpf,08,1))
	wn9 = VAL(SUBS(wcpf,09,1))
	wn10 = VAL(SUBS(wcpf,10,1))
	wn11 = VAL(SUBS(wcpf,11,1))

* CALCULO DO 1o digito
* --------------------
	soma1 = wn1*10+wn2*9+wn3*8+wn4*7+wn5*6+wn6*5+wn7*4+wn8*3+wn9*2
	dig1  =  11 - MOD(soma1,11)

	IF dig1 = 10 .OR. dig1 = 11
		dig1 = 0
	ENDIF
	IF dig1 <> wn10
		RETURN .F.
	ENDIF

* CALCULO DO 2o digito
* --------------------
	soma2 = wn1*11+wn2*10+wn3*9+wn4*8+wn5*7+wn6*6+wn7*5+wn8*4+wn9*3+wn10*2
	dig2  =  11 - MOD(soma2,11)

	IF dig2 = 10 .OR. dig2 = 11
		dig2 = 0
	ENDIF

	IF dig2 <> wn11
		RETURN .F.
	ENDIF

	RETURN .T.
ENDFUNC
****
*
* Email válido
*
****
FUNCTION EmailOK
	LPARAMETERS lcEmail
	IF VARTYPE(m.lcEmail)#"C" OR !("@"$m.lcEmail)
		RETURN .F.
	ENDIF
	LOCAL lnI
	IF TYPE("M.EmailValidChars")="U"
		PUBLIC EmailValidChars
		m.EmailValidChars = "0123456789@_-."
		FOR m.lnI = ASC('a') TO ASC('z')
			m.EmailValidChars = m.EmailValidChars + CHR(m.lnI)
		NEXT
	ENDIF
	FOR m.lnI = 1 TO LEN(m.lcEmail)
		IF !SUBSTR(m.lcEmail,m.lnI,1)$m.EmailValidChars
			RETURN .F.
		ENDIF
	NEXT
	IF !"."$GETWORDNUM(m.lcEmail,2,'@')
		RETURN .F.
	ENDIF
	RETURN .T.
ENDFUNC

*------------------------------------------------------------
FUNCTION CNPJOk
* Parametro : CNPJ a verificar (C14)
* Retorna : .T. se confirmado
*------------------------------------------------------------

	PARAMETERS wcgc
	wn1 = VAL(SUBS(wcgc,01,1))
	wn2 = VAL(SUBS(wcgc,02,1))
	wn3 = VAL(SUBS(wcgc,03,1))
	wn4 = VAL(SUBS(wcgc,04,1))
	wn5 = VAL(SUBS(wcgc,05,1))
	wn6 = VAL(SUBS(wcgc,06,1))
	wn7 = VAL(SUBS(wcgc,07,1))
	wn8 = VAL(SUBS(wcgc,08,1))
	wn9 = VAL(SUBS(wcgc,09,1))
	wn10 = VAL(SUBS(wcgc,10,1))
	wn11 = VAL(SUBS(wcgc,11,1))
	wn12 = VAL(SUBS(wcgc,12,1))
	wn13 = VAL(SUBS(wcgc,13,1))
	wn14 = VAL(SUBS(wcgc,14,1))

* CALCULO DO 13o ALGARISMO
* ------------------------
	soma1   =  wn1*5+wn2*4+wn3*3+wn4*2+wn5*9+wn6*8+wn7*7+wn8*6+wn9*5+wn10*4+wn11*3+wn12*2
	dig1	=  11 - MOD(soma1,11)
	IF dig1 = 10 .OR. dig1 = 11
		dig1 = 0
	ENDIF

	IF dig1 <> wn13
		RETURN .F.
	ENDIF

* CALCULO DO 14o ALGARISMO
* ------------------------
	soma2  =  wn1*6+wn2*5+wn3*4+wn4*3+wn5*2+wn6*9+wn7*8+wn8*7+wn9*6+wn10*5+wn11*4+wn12*3+wn13*2
	dig2   =  11 - MOD(soma2,11)
	IF dig2 = 10 .OR. dig2 = 11
		dig2 = 0
	ENDIF

	IF dig2 <> wn14
		RETURN .F.
	ENDIF

	RETURN .T.
ENDFUNC