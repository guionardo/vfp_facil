****
*
* Função de arquivamento de arquivos
* Compacta arquivos em uma pasta, separando por data
*
* lcPasta	Caminho da pasta a ser verificada
* lcOpcoes	Opções, separadas por |
*				ext=XXX		Indica extensão(ões) dos arquivos
*							ext=TXT		ext=TXT,LOG
*				sep=XXX		Indica separação em arquivos (Y=Ano, M=Mês, D=Dia)
*							sep=YMA	sep=YM
*				pref=XXX	Indica prefixo do arquivo zip gerado
*							sep=YMA|pref=ETQ produz ETQ20160226.ZIP
*				dest=XXX	Indica pasta de destino
*				zip=1		Indica o uso de arquivos ZIP ao invés de pastas
*				idade=nnX	Indica a idade mínima para processar
*							1d = 1 dia 	2m = 2 meses	1a = 1 ano
FUNCTION GS_Arquivador
	LPARAMETERS lcPasta, lcOpcoes
	IF (VARTYPE(m.lcPasta)#"C") OR !DIRECTORY(m.lcPasta)
		RETURN .F.
	ENDIF
	
	IF TYPE("M._GSALimite")="U"
	PUBLIC _GSALimite
	ENDIF
	IF VARTYPE(m._GSALimite)#"N" OR m._GSALimite<1
	m._GSALimite = 512
	ENDIF 

	IF VARTYPE(m.lcOpcoes)#"C"
		m.lcOpcoes = ""
	ENDIF
	m.lcPasta = FULLPATH(m.lcPasta)
	LOCAL lcExt, lcSep, lcPref, lcDest, llZip, lcIdade, lnI, lnMovidos
	m.lcExt = "*"
	m.lcSep = "YMD"
	m.lcPref = "arquivo"
	m.lcDest = m.lcPasta
	m.llZip = ""
	m.lcIdade = ""
	m.lnMovidos = 0

	GS_ExtraiOp(@lcOpcoes,"ext",@lcExt)
	GS_ExtraiOp(@lcOpcoes,"sep",@lcSep)
	GS_ExtraiOp(@lcOpcoes,"pref",@lcPref)
	GS_ExtraiOp(@lcOpcoes,"dest",@lcDest)
	GS_ExtraiOp(@lcOpcoes,"zip",@llZip)
	GS_ExtraiOp(@lcOpcoes,"idade",@lcIdade)
	m.llZip = (m.llZip=="1")


	LOCAL laExt[GETWORDCOUNT(m.lcExt,",")]
	FOR m.lnI = 1 TO ALEN(m.laExt)
		m.laExt[m.lnI] = GETWORDNUM(m.lcExt,m.lnI,",")
	NEXT

	IF ALEN(m.laExt)>1
		m.lcOpcoes = "ext="+m.laExt[m.lnI]+'|sep='+m.lcSep+'|pref='+m.lcPref+'|dest='+m.lcDest
		FOR m.lnI = 1 TO ALEN(m.laExt)
			m.lnMovidos = m.lnMovidos + GS_Arquivador(m.lcPasta,m.lcOpcoes)
		NEXT
		RETURN m.lnMovidos
	ENDIF

	LOCAL laD(1)
	IF ADIR(m.laD,ADDBS(m.lcPasta)+"*."+m.laExt[1])=0
		RETURN 0
	ENDIF

	IF ALEN(m.laD,1)>m._GSALimite
		MESSAGEBOX("Foram encontrados "+TRANSFORM(ALEN(m.laD,1))+" arquivos para arquivamento."+CHR(13)+CHR(13)+;
			"Por questão de economia de recursos, serão processados apenas os primeiros "+;
			TRANSFORM(m._GSALimite)+".",48,"*** ATENÇÃO ***")
	ENDIF

	DECLARE INTEGER MoveFile IN Win32API ;
		STRING @MyFileorFolderNameNow, ;
		STRING @TheNewNameIWant

	LOCAL laPastas(1), lnPastas, lcPastaDest
	m.laPastas = ""
	m.lnPastas = 0
	PUBLIC _GSATotArq, _GSAMovArq
	m._GSATotArq = MIN(ALEN(m.laD,1),m._GSALimite)
	m._GSAMovArq = 0
	FOR m.lnI = 1 TO ALEN(m.laD,1)
		m._GSAMovArq = m._GSAMovArq + 1
		IF LEFT(m.laD(m.lnI,1),1)='.'
			LOOP
		ENDIF
		IF GS_MoveArqData(ADDBS(m.lcPasta)+m.laD(m.lnI,1),m.lcPref,m.lcSep,m.lcDest,@lcPastaDest)
			m.lnMovidos = m.lnMovidos + 1

			IF ASCAN(m.laPastas, m.lcPastaDest)=0
				IF ALEN(m.laPastas)<m.lnPastas+1
					DIMENSION m.laPastas[m.lnPastas+10]
				ENDIF
				m.lnPastas = m.lnPastas + 1
				m.laPastas[m.lnPastas] = m.lcPastaDest
			ENDIF
		ENDIF
		IF m.lnMovidos = 512
			EXIT
		ENDIF
	NEXT
	RELEASE _GSATotArq, _GSAMovArq


	IF !m.llZip
		RETURN m.lnMovidos
	ENDIF

	WAIT WINDOW "Arquivos movidos para pastas. Tecle algo para zipar"

	FOR m.lnI = 1 TO m.lnPastas
		WAIT WINDOW NOWAIT "Compactando "+m.laPastas[m.lnI]+"..."
		GSZip_Adiciona(m.laPastas[m.lnI],ADDBS(m.laPastas[m.lnI])+"*.*",.T.)
	NEXT

	RETURN m.lnMovidos
ENDFUNC

****
*
* Extrai opção da string informada, removendo
*
* @lcTxt		String informada
* lcOp		Nome da opção
* @lcValor	Valor extraido
* Retorna .T. em caso de sucesso
*
FUNCTION GS_ExtraiOp
	LPARAMETERS lcTxt, lcOp, lcValor
	IF !((m.lcOp+'=')$m.lcTxt)
		RETURN .F.
	ENDIF

	LOCAL lnP, lcT
	m.lnP = AT(m.lcOp+'=',m.lcTxt)
	m.lcT = SUBSTR(m.lcTxt,m.lnP)
	m.lnP = AT('|',m.lcT)-1
	IF m.lnP<1
		m.lnP = LEN(m.lcT)
	ENDIF
	m.lcT = LEFT(m.lcT,m.lnP)

	m.lcValor = SUBSTR(m.lcT,LEN(m.lcOp)+2)
	m.lcTxt = STRTRAN(m.lcTxt,m.lcT,'')
	m.lcTxt = STRTRAN(m.lcTxt,'||','|')
	RETURN .T.
ENDFUNC


****
*
* Move arquivo para pasta com padrão de data
*
FUNCTION GS_MoveArqData
	LPARAMETERS lcArq, lcPref, lcSep, lcDest, lcPastaDest, lcIdade
	IF !FILE(m.lcArq)
		RETURN .F.
	ENDIF
	LOCAL laD(1)
	ADIR(m.laD,m.lcArq)
	LOCAL lcD
	m.lcD =""
	IF (VARTYPE(m.lcIdade)="C") AND ;
			(m.laD(1,3)>GS_Idade(m.lcIdade))
		RETURN .F.
	ENDIF

	FOR m.lnI = 1 TO LEN(m.lcSep)
		DO CASE
			CASE SUBSTR(m.lcSep,m.lnI,1)="Y"
				m.lcD = m.lcD + STR(YEAR(m.laD(1,3)),4,0)
			CASE SUBSTR(m.lcSep,m.lnI,1)="M"
				m.lcD = m.lcD + STR(MONTH(m.laD(1,3)),2,0)
			CASE SUBSTR(m.lcSep,m.lnI,1)="D"
				m.lcD = m.lcD + STR(DAY(m.laD(1,3)),2,0)
		ENDCASE
	NEXT
	m.lcD = ADDBS(m.lcDest) + m.lcPref + "_" + STRTRAN(m.lcD,' ','0')	&& Pasta de destino
	m.lcPastaDest = m.lcD

	IF !DIRECTORY(m.lcD)
		MKDIR (m.lcD)
		IF !DIRECTORY(m.lcD)
			m.lcPastaDest = ""
			RETURN .F.
		ENDIF
	ENDIF
	LOCAL lnR
	IF MOD(m._GSAMovArq,FLOOR(m._GSATotArq/100))=0
		WAIT WINDOW NOWAIT "Movendo "+m.lcArq+" para "+m.lcD
	ENDIF
	m.lnR = MoveFile(m.lcArq,ADDBS(m.lcD)+JUSTFNAME(m.lcArq))
	WAIT CLEAR
	RETURN m.lnR>0

ENDFUNC

*
FUNCTION GS_Idade
	LPARAMETERS lcIdade
	IF RIGHT(m.lcIdade,1)="d"
		RETURN (DATE()-ABS(VAL(m.lcIdade)))
	ENDIF
	IF RIGHT(m.lcIdade,1)="m"
		RETURN GOMONTH(DATE(),-ABS(VAL(m.lcIdade)))
	ENDIF
	IF RIGHT(m.lcIdade,1)="y" OR RIGHT(m.lcIdade,1)="a"
		RETURN DATE(YEAR(DATE())-ABS(VAL(m.lcIdade)),MONTH(DATE()),DAY(DATE()))
	ENDIF
	RETURN DATE()
ENDFUNC
