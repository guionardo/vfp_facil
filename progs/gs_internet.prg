*
* Fun��es de acesso INTERNET
*
*

#DEFINE HTTPREQUEST_SETCREDENTIALS_FOR_SERVER 0
#DEFINE HTTPREQUEST_SETCREDENTIALS_FOR_PROXY 1
#DEFINE HTTPREQUEST_PROXYSETTING_PROXY 2
#DEFINE GENERIC_READ     2147483648
#DEFINE GENERIC_WRITE     1073741824
#DEFINE INTERNET_INVALID_PORT_NUMBER 0
#DEFINE INTERNET_OPEN_TYPE_DIRECT  1
#DEFINE INTERNET_OPEN_TYPE_PROXY  3
#DEFINE INTERNET_DEFAULT_FTP_PORT  21
#DEFINE INTERNET_FLAG_ASYNC    268435456
#DEFINE INTERNET_FLAG_FROM_CACHE  16777216
#DEFINE INTERNET_FLAG_OFFLINE   16777216
#DEFINE INTERNET_FLAG_CACHE_IF_NET_FAIL 65536
#DEFINE INTERNET_OPEN_TYPE_PRECONFIG 0
#DEFINE INTERNET_SERVICE_FTP   1
#DEFINE INTERNET_SERVICE_GOPHER   2
#DEFINE INTERNET_SERVICE_HTTP   3
#DEFINE FTP_TRANSFER_TYPE_ASCII   1
#DEFINE FTP_TRANSFER_TYPE_BINARY  2
#DEFINE FILE_ATTRIBUTE_NORMAL   128
#DEFINE FORMAT_MESSAGE_FROM_SYSTEM 0x1000
#DEFINE PROXY_TEST_URL "http://tbyte.com.br"
#DEFINE FTP_AGENT "VFP9_INTERNET"
#DEFINE FTP_CHUNKSIZE 512
#DEFINE PROXY_CFG 'PROXY.CFG'

*
* Efetua o download do conte�do da URL
*	lcURL		URL
*	@lcRetorno	String com o conte�do
*	@lcMsg		Mensagem de retorno
*	llBinary	Download bin�rio = .t. / Texto = .f.
*	lcMsgAnt	Mensagem fixa do wait
*	RETURN		BOOL	Sucesso
*
FUNCTION DownloadToString(lcURL, lcRetorno, lcMsg, llBinary, lcMsgAnt)
	_DECLARE()
	IF VARTYPE(m.lcURL)#'C'
		RETURN .F.
	ENDIF
	LOCAL loHTTP, loExc AS EXCEPTION, llErro
	m.lcRetorno = ''
	IF VARTYPE(m.lcMsg)#"C"
		m.lcMsg = "Obtendo dados de "+m.lcURL
	ENDIF
	IF VARTYPE(m.lcMsgAnt)="C"
		m.lcMsg = LEFT(ALLTRIM(m.lcMsgAnt)+CHR(13)+m.lcMsg,255)
	ENDIF
	IF VARTYPE(m.llBinary)#"L"
		m.llBinary = .F.
	ENDIF
	WAIT WINDOW NOWAIT m.lcMsg
	ProxyCheck()
	LOCAL lcErro, llProxy
	m.llProxy = PEMSTATUS(_SCREEN,'PROXY',5) AND _SCREEN.PROXY.ACTIVE

	m.lcErro = 'Erro ao acessar a URL:'+CHR(13)+;
		m.lcURL+CHR(13)+;
		IIF(m.llProxy, 'Via proxy '+_SCREEN.PROXY.HOST+':'+TRANSFORM(_SCREEN.PROXY.PORT)+' (USUARIO '+_SCREEN.PROXY.USER+')'+ CHR(13)+CHR(13),'')
	TRY
		m.lcRetorno = ''
		loHTTP = CREATEOBJECT("WinHttp.WinHttpRequest.5.1")
		LOCAL lcMsg
		m.lcMsg = ''

		loHTTP.OPEN("GET", m.lcURL, .F.)
		ProxySet(m.loHTTP)
		loHTTP.SEND()
		m.lcMsg = TRANSFORM(m.loHTTP.STATUS)+' '+m.loHTTP.STATUSTEXT
		IF m.loHTTP.STATUS = 200
			IF m.llBinary
				m.lcRetorno = m.loHTTP.ResponseBody
			ELSE
				m.lcRetorno = m.loHTTP.ResponseText
			ENDIF
		ELSE
			_MSG(m.lcErro+m.lcMsg,.T.)
			m.llErro = .T.
		ENDIF
	CATCH TO loExc
		_MSG(m.lcErro+"\n"+m.loExc.MESSAGE+"\n"+m.loExc.PROCEDURE +"("+m.loExc.LINENO+")" ,.T.)
		m.llErro = .T.
	ENDTRY
	WAIT CLEAR
	RETURN !m.llErro
ENDFUNC

FUNCTION DownloadToFile(lcURL, lcPastaDestino,lcMsg,lcMsgAnt)
	LOCAL lcConteudo
	IF !DownloadToString(m.lcURL,@lcConteudo,@lcMsg,.T.,m.lcMsgAnt) OR EMPTY(m.lcConteudo)
		RETURN .F.
	ENDIF

*!*		IF RIGHT(LOWER(m.lcURL),4)=='.zip'
*!*			LOCAL lcTmp
*!*			m.lcTmp = ADDBS(SYS(2023))+SYS(3)+'.zip'
*!*			STRTOFILE(m.lcConteudo,m.lcTmp)
*!*			AZ_Extrair(m.lcTmp,m.lcPastaDestino)
*!*		ELSE
	STRTOFILE(m.lcConteudo,m.lcPastaDestino)
*!*		ENDIF

*
* Download de um arquivo via FTP
* lcFTP		HostFTP
* lcUser	Usu�rio
* lcPass	Senha
* lcPath	Caminho do arquivo no servidor
* lcDestino	Nome do arquivo local
* RETURN	BOOL
*
FUNCTION FTPDownload(lcFTP,lcUser,lcPass,lcPath,lcDestino)
	_Log("FTP Download "+m.lcFTP+'\'+m.lcPath+" -> "+m.lcDestino)
	WAIT WINDOW NOWAIT "FTP DOWNLOAD"+CHR(13)+m.lcFTP+CHR(13)+m.lcPath
	IF FILE(m.lcDestino) AND DeleteFile(m.lcDestino)=0
		_Log("  !ERRO AO EXCLUIR "+m.lcDestino)
		RETURN .F.
	ENDIF

	_DECLARE()
	ProxyCheck()

	LOCAL lnHandler,lnSession, llSucesso
	IF !_FTPConnect(m.lcFTP,m.lcUser,m.lcPass,@lnHandler,@lnSession)
		RETURN .F.
	ENDIF

	LOCAL hLocal
	m.hLocal = FCREATE(m.lcDestino)
	IF m.hLocal = -1
		_Log("  !ERRO AO CRIAR "+m.lcDestino)
		RETURN .F.
	ENDIF

	TRY
		WAIT WINDOW NOWAIT "Recebendo arquivo "+m.lcPath+" de FTP "+m.lcFTP

		LOCAL lcErro, lnFileSize, hTarget, lnFS2
		hTarget = FtpOpenFile(m.lnSession, m.lcPath, GENERIC_READ, 2, 0)
		IF m.hTarget > 0
			m.lnFileSize = 0
			m.lnFS2 = 0
			m.lnFileSize = FtpGetFileSize(m.hTarget, @lnFS2)

			lnTotBytesReaden = 0
			lnBytesReaden = 0
			m.lcBuffer = REPLICATE(CHR(0), FTP_CHUNKSIZE)
			DO WHILE InternetReadFile(hTarget, @lcBuffer, FTP_CHUNKSIZE, @lnBytesReaden)#0 AND ;
					m.lnBytesReaden > 0
				m.lnTotBytesReaden = m.lnTotBytesReaden + m.lnBytesReaden
				WAIT WINDOW NOWAIT "Download de "+m.lcFTP+'/'+m.lcPath+CHR(13)+;
					TRANSFORM(100*m.lnTotBytesReaden/m.lnFileSize,"999.9%")+CHR(13)+;
					_BYTES(m.lnTotBytesReaden)+' / '+_BYTES(m.lnFileSize)
				FWRITE(m.hLocal,m.lcBuffer,m.lnBytesReaden)
			ENDDO
			FCLOSE(m.hLocal)
			m.hLocal = 0
			InternetCloseHandle(hTarget)

			_Log("FTPDownload "+M.lcPath+" OK ("+TRANSFORM(m.lnTotBytesReaden)+"B)")
			m.llSucesso = .T.
		ELSE
			m.lcErro = _WinApiErrMsg()
			_Log("!FTPDownload = ERRO "+m.lcErro)
			= FCLOSE (hSource)

		ENDIF
	CATCH TO loExc
		_Log(" !ERRO "+m.loExc.MESSAGE)
		_MSG("ERRO: "+m.loExc.MESSAGE)
	FINALLY
		IF m.hLocal > 0
			FCLOSE(m.hLocal)
		ENDIF
		InternetCloseHandle(m.lnSession)
		InternetCloseHandle(m.lnHandler)
		WAIT CLEAR
	ENDTRY

	RETURN m.llSucesso
ENDFUNC

*
* Download de um arquivo via FTP
* lcFTP		HostFTP
* lcUser	Usu�rio
* lcPass	Senha
* lcPath	Caminho do arquivo no servidor
* lcOrigem	Nome do arquivo local
* RETURN	BOOL
*
FUNCTION FTPUpload(lcFTP,lcUser,lcPass,lcPath,lcOrigem)
	_Log("FTP Upload "+m.lcOrigem+ " -> "+m.lcFTP+'\'+m.lcPath)
	WAIT WINDOW NOWAIT "FTP UPLOAD"+CHR(13)+m.lcFTP+CHR(13)+m.lcPath
	IF !FILE(m.lcOrigem)
		_Log("  !ORIGEM INEXISTENTE "+m.lcOrigem)
		RETURN .F.
	ENDIF
	LOCAL hLocal, lnTamanho, laF(1)
	IF ADIR(m.laF,m.lcOrigem)=0
		_Log("  !ERRO AO ACESSAR ARQUIVO "+m.lcOrigem)
		RETURN .T.
	ENDIF
	m.lnTamanho = m.laF(1,2)

	m.hLocal = FOPEN(m.lcOrigem)
	IF m.hLocal = -1
		_Log("  !ERRO AO ABRIR "+m.lcOrigem)
		RETURN .F.
	ENDIF

	LOCAL lnHandler,lnSession, llSucesso, hTarget, lnBytesWritten
	IF !_FTPConnect(m.lcFTP,m.lcUser,m.lcPass,@lnHandler,@lnSession)
		FCLOSE(m.hLocal)
		RETURN .F.
	ENDIF
	TRY
		m.hTarget = FtpOpenFile(m.lnSession, m.lcPath, GENERIC_WRITE, 2, 0)
		IF m.hTarget > 0
			lnBytesWritten = 0
			DO WHILE !FEOF(m.hLocal)
				m.lcBuffer = FREAD (m.hLocal, FTP_CHUNKSIZE)
				m.lnLength = LEN(lcBuffer)
				IF m.lnLength # 0
					WAIT WINDOW NOWAIT "Enviando arquivo "+m.lcPath+" para FTP "+m.lcFTP+CHR(13)+;
						TRANSFORM(100*m.lnBytesWritten/m.lnTamanho,"999.9%")+CHR(13)+;
						_BYTES(m.lnBytesWritten)+" / "+_BYTES(m.lnTamanho)
					IF InternetWriteFile (hTarget, @lcBuffer, lnLength, @lnLength) = 1
						lnBytesWritten = lnBytesWritten + lnLength
					ELSE
						EXIT
					ENDIF
				ELSE
					EXIT
				ENDIF
			ENDDO
			m.llSucesso = .T.
			InternetCloseHandle(hTarget)
		ELSE
			m.lcErro = _WinApiErrMsg()
			_Log("!FTPUpload = ERRO "+m.lcErro)
		ENDIF
	CATCH TO loExc
		_Log(" !ERRO "+m.loExc.MESSAGE)
		_MSG("ERRO: "+m.loExc.MESSAGE)
	FINALLY
		InternetCloseHandle(m.lnSession)
		InternetCloseHandle(m.lnHandler)
		FCLOSE(m.hLocal)
		WAIT CLEAR
	ENDTRY
	RETURN m.llSucesso
ENDFUNC

*
* Abre conex�o com servidor FTP
* lcHost	Servidor FTP
* lcUser	Usu�rio
* lcPass	Senha
* @lnHandler	Handler de conex�o
* @lnSession	Handler de sess�o
* RETURN 	BOOL
*
FUNCTION _FTPConnect(lcHost, lcUser, lcPass, lnHandler, lnSession)
	WAIT WINDOW NOWAIT "Conectando ao FTP "+m.lcHost
	_Log("FTPConnect "+m.lcHost)
	LOCAL sAgent, sProxyName, sProxyBypass, lFlags, lpszRemoteFile, lpszNewFile, fFailIfExists, dwContext, llTransferType, lnResult, lcErro
	sProxyName = CHR(0)
	sProxyBypass = CHR(0)
	lFlags = 0
	m.lnSession = 0
	IF !_SCREEN.PROXY.ACTIVE
		m.lnHandler = InternetOpen(FTP_AGENT, INTERNET_OPEN_TYPE_DIRECT, sProxyName, sProxyBypass, lFlags)
	ELSE
		sProxyName = _SCREEN.PROXY.HOST+':'+TRANSFORM(_SCREEN.PROXY.PORT)
		m.lnHandler = InternetOpen (sAgent, INTERNET_OPEN_TYPE_PROXY, sProxyName, sProxyBypass, lFlags)
		InternetSetOption(m.lnH, INTERNET_OPTION_PROXY_PASSWORD, _SCREEN.PROXY.PASS, LEN(_SCREEN.PROXY.PASS))
		InternetSetOption(m.lnH, INTERNET_OPTION_PROXY_USERNAME, _SCREEN.PROXY.USER, LEN(_SCREEN.PROXY.USER))
	ENDIF

	IF m.lnHandler=0
		_Log("!FTPCONNECT => ERRO WININET.DLL")
		_MSG("Sem acesso a WinInet.dll",.T.)
		RETURN .F.
	ENDIF
	m.lnSession = InternetConnect(m.lnHandler, m.lcHost, 0, lcUser, lcPass, 1, 0, 0)
	IF m.lnSession = 0
		LOCAL lcErro
		m.lcErro = _WinApiErrMsg()
		_Log("!FTP.Connect => "+m.lcErro)
		InternetCloseHandle(m.lnHandler)
		m.lnHandler = 0
		_MSG("Erro de conex�o: "+m.lcErro,.T.)
		RETURN .F.
	ENDIF

	_Log("FTPCONNECT OK")
	RETURN .T.
ENDFUNC

FUNCTION ProxyCheck(llApenasTesta)
	IF !PEMSTATUS(_SCREEN,"PROXY",5)
		IF m.llApenasTesta
			RETURN .F.
		ENDIF
		ADDPROPERTY(_SCREEN,'PROXY')
		LOCAL loP
		m.loP = CREATEOBJECT('EMPTY')
		ADDPROPERTY(m.loP,'HOST','')
		ADDPROPERTY(m.loP,'PORT',0)
		ADDPROPERTY(m.loP,'USER','')
		ADDPROPERTY(m.loP,'PASS','')
		ADDPROPERTY(m.loP,'ACTIVE',.F.)
		_SCREEN.PROXY = m.loP
	ENDIF
	IF !(PEMSTATUS(_SCREEN.PROXY,'HOST',5) AND ;
			PEMSTATUS(_SCREEN.PROXY,'PORT',5) AND ;
			PEMSTATUS(_SCREEN.PROXY,'USER',5) AND ;
			PEMSTATUS(_SCREEN.PROXY,'PASS',5) AND ;
			PEMSTATUS(_SCREEN.PROXY,'ACTIVE',5))
		IF m.llApenasTesta
			RETURN .F.
		ENDIF
		REMOVEPROPERTY(_SCREEN,'PROXY')
		RETURN ProxyCheck()
	ENDIF

	IF VARTYPE(_SCREEN.PROXY.ACTIVE)#"L"
		_SCREEN.PROXY.ACTIVE = .F.
	ENDIF
	IF VARTYPE(_SCREEN.PROXY.HOST)#'C'
		_SCREEN.PROXY.HOST = ''
		_SCREEN.PROXY.ACTIVE = .F.
	ENDIF
	IF VARTYPE(_SCREEN.PROXY.PORT)#'N'
		_SCREEN.PROXY.PORT = 0
		_SCREEN.PROXY.ACTIVE = .F.
	ENDIF
	IF VARTYPE(_SCREEN.PROXY.USER)#'C'
		_SCREEN.PROXY.USER = ''
		_SCREEN.PROXY.ACTIVE = .F.
	ENDIF
	IF VARTYPE(_SCREEN.PROXY.PASS)#'C'
		_SCREEN.PROXY.PASS = ''
		_SCREEN.PROXY.ACTIVE = .F.
	ENDIF
ENDFUNC

*
* Testa conectividade com proxy
*
FUNCTION ProxyTest
	IF !ProxyCheck(.T.)
* N�o h� informa��o de Proxy, nem configura��o carregada
		_MSG("N�o h� configura��o de proxy!")
		RETURN .F.
	ENDIF

	ProxyCheck()
	IF EMPTY(_SCREEN.PROXY.HOST)
		_MSG("N�o h� host para o proxy!")
		RETURN .F.
	ENDIF
	_SCREEN.PROXY.ACTIVE = .T.

	LOCAL lcLog
	m.lcLog = "ProxyTest "+;
		_SCREEN.PROXY.HOST+":"+TRANSFORM(_SCREEN.PROXY.PORT)+" "+;
		_SCREEN.PROXY.USER+":"+_SCREEN.PROXY.PASS+" "+IIF(_SCREEN.PROXY.ACTIVE,"ATIVO","INATIVO")+" : "
	LOCAL lcRetorno, llSucesso
	IF !DownloadToString(PROXY_TEST_URL,@lcRetorno)
		_SCREEN.PROXY.ACTIVE = .F.
		IF !DownloadToString(PROXY_TEST_URL,@lcRetorno)
			_Log('!'+m.lcLog+"TESTE FALHOU")
			_MSG("Teste de proxy falhou!",.T.)
		ELSE
			_Log('!'+m.lcLog+"CONEX�O SEM NECESSIDADE DE PROXY")
			_MSG("Conex�o OK sem necessidade de proxy!")
			m.llSucesso = .T.
		ENDIF
	ELSE
		_Log(m.lcLog+"TESTE DE PROXY OK")
		_MSG("Teste de proxy OK!")
		m.llSucesso = .T.
	ENDIF
	RETURN m.llSucesso
ENDFUNC


FUNCTION ProxyLoad(lcArqConfig)
	ProxyCheck()

	IF VARTYPE(m.lcArqConfig)#"C"
		m.lcArqConfig= ADDBS(JUSTPATH(APPLICATION.SERVERNAME))+PROXY_CFG
	ENDIF
	IF DIRECTORY(m.lcArqConfig)
		m.lcArqConfig = ADDBS(m.lcArqConfig)+PROXY_CFG
	ENDIF
	IF !FILE(m.lcArqConfig)
		_Log("!ProxyLoad("+m.lcArqConfig+") INEXISTENTE")
		RETURN .F.
	ENDIF

	LOCAL lcH
	m.lcH = FILETOSTR(m.lcArqConfig)

	IF GETWORDCOUNT(m.lcH,'|')#5
		_Log("!ProxyLoad("+m.lcArqConfig+") PARAMETROS INVALIDOS")
		_MSG("N�mero de par�metros inv�lidos na configura��o do proxy em "+m.lcArqConfig)
		RETURN .F.
	ENDIF
	LOCAL lcHost,lnPort,lcUser,lcPass,llActive
	m.llActive = GETWORDNUM(m.lcH,1,'|')='1'
	m.lcHost = GETWORDNUM(m.lcH,2,'|')
	m.lnPort = VAL(GETWORDNUM(m.lcH,3,'|'))
	m.lcUser = GETWORDNUM(m.lcH,4,'|')
	m.lcPass = GETWORDNUM(m.lcH,5,'|')

	IF EMPTY(m.lcHost) OR ;
			!BETWEEN(m.lnPort,1,65535) OR ;
			EMPTY(m.lcUser) OR ;
			EMPTY(m.lcPass)
		_Log("!ProxyLoad("+m.lcArqConfig+") PARAMETROS INVALIDOS")
		_MSG("Par�metro inv�lido em "+m.lcArqConfig)
		RETURN .F.
	ENDIF

	_SCREEN.PROXY.ACTIVE	= m.llActive
	_SCREEN.PROXY.HOST		= m.lcHost
	_SCREEN.PROXY.PORT		= m.lnPort
	_SCREEN.PROXY.USER		= m.lcUser
	_SCREEN.PROXY.PASS		= m.lcPass
	_Log("ProxyLoad("+m.lcArqConfig+") "+m.lcHost+":"+TRANSFORM(m.lnPort)+" "+m.lcUser+":"+m.lcPass+" "+IIF(m.llActive,"ATIVO","INATIVO"))
	RETURN .T.
ENDFUNC

FUNCTION ProxySave(lcArqConfig)
	ProxyCheck()
	IF VARTYPE(m.lcArqConfig)#"C"
		m.lcArqConfig= ADDBS(JUSTPATH(APPLICATION.SERVERNAME))+PROXY_CFG
	ELSE
		IF DIRECTORY(m.lcArqConfig)
			m.lcArqConfig = ADDBS(m.lcArqConfig)+PROXY_CFG
		ENDIF
	ENDIF

	LOCAL lnH, lcH
	IF FILE(m.lcArqConfig) AND DeleteFile(m.lcArqConfig)=0
		_MSG("N�o foi poss�vel excluir o arquivo de configura��o do proxy em "+m.lcArqConfig,.T.)
		RETURN .F.
	ENDIF

	IF STRTOFILE(IIF(_SCREEN.PROXY.ACTIVE,'1','0')+'|'+_SCREEN.PROXY.HOST+'|'+TRANSFORM(_SCREEN.PROXY.PORT)+'|'+_SCREEN.PROXY.USER+'|'+_SCREEN.PROXY.PASS,m.lcArqConfig) > 0
		_Log("ProxySave("+m.lcArqConfig+") "+;
			_SCREEN.PROXY.HOST+":"+TRANSFORM(_SCREEN.PROXY.PORT)+" "+;
			_SCREEN.PROXY.USER+":"+_SCREEN.PROXY.PASS+" "+IIF(_SCREEN.PROXY.ACTIVE,"ATIVO","INATIVO"))
		RETURN .T.
	ELSE
		_Log("!ProxySave("+m.lcArqConfig+") "+;
			_SCREEN.PROXY.HOST+":"+TRANSFORM(_SCREEN.PROXY.PORT)+" "+;
			_SCREEN.PROXY.USER+":"+_SCREEN.PROXY.PASS+" "+IIF(_SCREEN.PROXY.ACTIVE,"ATIVO","INATIVO"))
		RETURN .F.
	ENDIF

ENDFUNC

*
* Define configura��es de proxy para objeto WinHttp.WinHttpRequest.5.1
*
FUNCTION ProxySet(loHTTP)
	IF m.loHTTP#"O"
		RETURN .F.
	ENDIF
	ProxyCheck()
	LOCAL llProxy, llErro
	m.llProxy = PEMSTATUS(_SCREEN,'PROXY',5) AND _SCREEN.PROXY.ACTIVE

	IF !m.llProxy
		RETURN .T.
	ENDIF

	TRY
		m.loHTTP.SetProxy(HTTPREQUEST_PROXYSETTING_PROXY,_SCREEN.PROXY.HOST+":"+TRANSFORM(_SCREEN.PROXY.PORT))
		m.loHTTP.SetCredentials(_SCREEN.PROXY.USER,_SCREEN.PROXY.PASS,HTTPREQUEST_SETCREDENTIALS_FOR_PROXY)
	CATCH
		m.llErro = .T.
	ENDTRY
	RETURN !m.llErro
ENDFUNC





FUNCTION _MSG(lcMsg,llErro)
	MESSAGEBOX(m.lcMsg,IIF(m.llErro,16,64),'Conex�o internet')
ENDFUNC

FUNCTION _DECLARE
	PUBLIC plInternetInitDeclare
	IF m.plInternetInitDeclare
		RETURN
	ENDIF
	m.plInternetInitDeclare = .T.

	DECLARE INTEGER InternetSetOption IN WININET.DLL INTEGER, INTEGER,STRING @, INTEGER

	DECLARE INTEGER InternetOpen IN WININET.DLL;
		STRING  sAgent,;
		INTEGER lAccessType,;
		STRING  sProxyName,;
		STRING  sProxyBypass,;
		STRING  lFlags
*
	DECLARE INTEGER InternetCloseHandle IN WININET.DLL;
		INTEGER hInet
*
	DECLARE INTEGER InternetConnect IN WININET.DLL;
		INTEGER hInternetSession,;
		STRING  sServerName,;
		INTEGER nServerPort,;
		STRING  sUsername,;
		STRING  sPassword,;
		INTEGER lService,;
		INTEGER lFlags,;
		INTEGER lContext
*
	DECLARE INTEGER FtpGetFile IN WININET.DLL;
		INTEGER hFtpSession,;
		STRING  lpszRemoteFile,;
		STRING  lpszNewFile,;
		INTEGER fFailIfExists,;
		INTEGER dwFlagsAndAttributes,;
		INTEGER dwFlags,;
		INTEGER dwContext
*
	DECLARE INTEGER FtpGetFileSize IN wininet;
		INTEGER   hFile,;
		INTEGER @ lpdwFileSizeHigh
*
	DECLARE INTEGER FtpDeleteFile IN WININET.DLL ;
		INTEGER nFile,;
		STRING  lpszFileName
*
	DECLARE INTEGER InternetWriteFile IN WININET.DLL;
		INTEGER   hFile,;
		STRING  @ sBuffer,;
		INTEGER   lNumBytesToWrite,;
		INTEGER @ dwNumberOfBytesWritten

	DECLARE LONG InternetReadFile IN wininet.DLL ;
		LONG     hFtpSession, ;
		STRING  @lpBuffer, ;
		LONG     dwNumberOfBytesToRead, ;
		LONG    @lpNumberOfBytesRead
*
	DECLARE INTEGER FtpOpenFile IN WININET.DLL;
		INTEGER hFtpSession,;
		STRING  sFileName,;
		INTEGER lAccess,;
		INTEGER lFlags,;
		INTEGER lContext

	DECLARE LONG GetLastError IN WIN32API

	DECLARE LONG FormatMessage IN kernel32 ;
		LONG dwFlags, LONG lpSource, LONG dwMessageId, ;
		LONG dwLanguageId, STRING @lpBuffer, ;
		LONG nSize, LONG Arguments

	#DEFINE FORMAT_MESSAGE_FROM_SYSTEM 0x1000

	DECLARE INTEGER GetActiveWindow ;
		IN Win32API
	DECLARE INTEGER GetWindow IN Win32API ;
		INTEGER HWND, INTEGER nType
	DECLARE INTEGER GetWindowText IN Win32API ;
		INTEGER HWND, STRING @cText, INTEGER nType
	DECLARE INTEGER DeleteFile IN kernel32;
		STRING lpFileName

ENDFUNC

FUNCTION _WinApiErrMsg
	LPARAMETERS tnErrorCode
	IF VARTYPE(m.tnErrorCode)#"N"
		m.tnErrorCode = GetLastError()
	ENDIF

	LOCAL lcErrBuffer, lnNewErr, lnFlag, lcErrorMessage
	lnFlag = FORMAT_MESSAGE_FROM_SYSTEM
	lcErrBuffer = REPL(CHR(0),1000)
	lnNewErr = FormatMessage(lnFlag, 0, tnErrorCode, 0, @lcErrBuffer,500,0)
	lcErrorMessage = TRANSFORM(tnErrorCode) + " " + LEFT(lcErrBuffer, AT(CHR(0),lcErrBuffer)- 1 )
	RETURN lcErrorMessage
ENDFUNC

FUNCTION _Log(lcMsg)
	LOCAL lcArq
	IF JUSTSTEM(APPLICATION.SERVERNAME)=='vfp9'
		m.lcArq = "A:\TBYTE_UPDATE\LOG"
	ELSE
		m.lcArq = ADDBS(JUSTPATH(APPLICATION.SERVERNAME))+"\LOG"
	ENDIF
	IF !DIRECTORY(m.lcArq)
		TRY
			MKDIR (m.lcArq)
		CATCH
			m.lcArq = ''
		ENDTRY
	ENDIF
	IF EMPTY(m.lcArq)
		RETURN .F.
	ENDIF
	m.lcArq = ADDBS(m.lcArq)+"INTERNET_"+DTOS(DATE())+".LOG"
	STRTOFILE(TTOC(DATETIME(),1)+' '+PADR(JUSTSTEM(APPLICATION.SERVERNAME),10)+' '+m.lcMsg+CHR(13)+CHR(10),m.lcArq,1)
ENDFUNC

FUNCTION _BYTES(lnB)
	DO CASE
		CASE m.lnB >= 1073741824
			RETURN STR(m.lnB/1073741824,5,1)+"GB"
		CASE m.lnB >= 1048576
			RETURN STR(m.lnB/1048576,5,1)+"MB"
		CASE m.lnB >= 1024
			RETURN STR(m.lnB/1024,5,1)+"KB"
		OTHERWISE
			RETURN STR(m.lnB,4,1)+"B"
	ENDCASE
ENDFUNC


PROCEDURE TesteInternet
	CLEAR
	ProxyLoad()
	LOCAL lcRetorno, lcMsg
	? DownloadToString('https://tbyte.com.br/tbyte/tbyte.txt',@lcRetorno,@lcMsg)
	? m.lcRetorno
	? m.lcMsg

	? FTPDownload('ftp.tbyte.com.br','****','****',"/tbyte.zip","A:\TEMP\TBYTE.zip")
	? FTPUpload('ftp.tbyte.com.br','****','****',"/VFP.EXE","A:\TEMP\VFP.EXE")

ENDPROC

*
* Verifica se há uma conexão com a internet
*
FUNCTION InternetOK
	DECLARE INTEGER InternetGetConnectedState IN WinInet ;
		INTEGER @lpdwFlags, INTEGER dwReserved

	LOCAL lnFlags, lnReserved, lnSuccess
	m.lnFlags=0
	m.lnReserved=0
	m.lnSuccess=InternetGetConnectedState(@lnFlags,m.lnReserved)
	RETURN  (m.lnSuccess=1)
ENDFUNC