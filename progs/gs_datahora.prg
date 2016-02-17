****
*
* Funçõees de data e hora
*
****

****
*
* Retorna a data/hora de um servidor NTP
*
* Depende da declaração prévia de ShellExecute, GetLastError, WinApiErrMsg e DeleteFile
****
FUNCTION DataHoraInternet
	LPARAMETERS lcNTPServer
	IF VARTYPE(m.lcNTPServer)#"C"
		m.lcNTPServer = "200.160.0.8" && a.ntp.br
	ENDIF
	LOCAL lcArq
	m.lcArq = ADDBS(GETENV("windir"))+"SYSTEM32\W32TM.EXE"
	IF !FILE(m.lcArq)
		MESSAGEBOX("W32TM.EXE nÃ£o encontrado!")
		RETURN CTOT("")
	ENDIF
	LOCAL lcCom, lcTmp, lcDT, lcSec, llNeg, lnSec
	m.lcTmp = ADDBS(SYS(2023))+"datahora_internet.txt"
	m.lcCom = m.lcArq+" /stripchart /samples:1 /computer:"+m.lcNTPServer+" /dataonly > "+ m.lcTmp

	WAIT WINDOW NOWAIT "NTP data/hora ("+m.lcNTPServer+")..."

	LOCAL lnE
	m.lnE = ShellExecute(0,"open",'cmd.exe','/c "'+m.lcCom+'"','',0)
	IF m.lnE<=32
		MESSAGEBOX("Erro na execuÃ§Ã£o de W32TM.EXE"+CHR(13)+;
			"Erro: "+TRANSFORM(m.lnE)+CHR(13)+;
			WinApiErrMsg(GetLastError()),48,"NTP data/hora ("+m.lcNTPServer+")",5000)
		RETURN CTOT("")
	ENDIF

	IF !FILE(m.lcTmp)
		LOCAL lnT
		m.lnT = 0
		DO WHILE (!FILE(m.lcTmp)) AND (m.lnT<2)
			WAIT WINDOW "NTP data/hora ("+m.lcNTPServer+") Aguardando retorno..." TIMEOUT 1
			m.lnT = m.lnT + 1
		ENDDO
		IF !FILE(m.lcTmp)
			MESSAGEBOX("Arquivo "+m.lcTmp+" nÃ£o encontrado!",48,"NTP data/hora ("+m.lcNTPServer+")",5000)
			RETURN CTOT("")
		ENDIF
	ENDIF

	m.lcDT= FILETOSTR(m.lcTmp)
	m.lnT = 2

	DO WHILE FILE(m.lcTmp)	AND (m.lnT>0)
		IF DeleteFile(m.lcTmp)=0
			IF FILE(m.lcTmp)
				WAIT WINDOW "Limpando temporÃ¡rios..." TIMEOUT 1
			ENDIF
		ENDIF
		m.lnT = m.lnT - 1
	ENDDO

	IF "ERRO"$UPPER(m.lcTmp)
		RETURN CTOT("")
	ENDIF
	m.lcSec = ALLTRIM(STREXTRACT(m.lcDT,"+","s"))
	IF EMPTY(m.lcSec)
		m.lcSec = ALLTRIM(STREXTRACT(m.lcDT,"-","s",1,4))
		llNeg = .T.
	ENDIF
	m.lcSec =STRTRAN(STRTRAN(m.lcSec,"s",""),".",",")
	m.lnSec = VAL(m.lcSec)
	m.lcDT= CTOT(ALLTRIM(STREXTRACT(m.lcDT,"â€š",".")))+VAL(M.lcSec)


	RETURN m.lcDT


*	w32tm /stripchart /samples:1 /computer:es.pool.ntp.org /dataonly
ENDFUNC


****
*
* Une dois valores DATE e STRING (hh:mm:ss) em um datetime
*
****
FUNCTION DataHora
	LPARAMETERS ldData, lcHora
	IF VARTYPE(m.ldData)#"D"
		RETURN {}
	ENDIF
	IF VARTYPE(m.lcHora)#"C"
		m.lcHora = "00:00:00"
	ENDIF
	RETURN CTOT(DTOC(m.ldData)+" "+m.lcHora)
ENDFUNC
